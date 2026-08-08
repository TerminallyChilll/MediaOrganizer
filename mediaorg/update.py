"""Self-update: how far behind the clone is, and how to catch up.

Two halves:

* :func:`check` asks git how many commits separate this clone from the
  branch it tracks. It is cached (:data:`CACHE_NAME`, next to the app) so a
  normal launch reads a file instead of hitting the network, and refreshed
  in a background thread so the menu never waits on it.
* :func:`run_update` does the catching up — fast-forward pull, then a
  dependency install if ``requirements.txt`` moved.

Nothing here imports pandas/guessit, so ``python run.py --update`` works on
an install whose dependencies are broken (which is exactly when you want to
pull a fix).

ASCII markers only ([OK], [!], ->) — this prints on Windows consoles that
predate UTF-8.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path

REPO_URL = "https://github.com/TerminallyChilll/MediaOrganizer"
DEFAULT_REMOTE = "origin"
DEFAULT_BRANCH = "main"
CACHE_NAME = ".mediaorg_update_check.json"
DEFAULT_INTERVAL_HOURS = 24.0
FETCH_TIMEOUT = 20      # seconds; a network call
GIT_TIMEOUT = 15        # seconds; local git plumbing
PIP_TIMEOUT = 600       # seconds; a dependency install may compile something

# State values for UpdateStatus.state:
#   'current'  — up to date with the tracked branch
#   'behind'   — an update is available
#   'ahead'    — local commits the remote does not have (a fork/dev clone)
#   'diverged' — both, so a fast-forward is impossible
#   'unknown'  — could not tell (no git, no network, not a clone, ...)
CURRENT, BEHIND, AHEAD, DIVERGED, UNKNOWN = (
    'current', 'behind', 'ahead', 'diverged', 'unknown')


def app_dir() -> Path:
    """The directory holding run.py — the clone we update."""
    return Path(__file__).resolve().parent.parent


def _version() -> str:
    try:
        from . import __version__
        return __version__
    except Exception:              # pragma: no cover - defensive
        return "unknown"


@dataclass
class UpdateStatus:
    """The answer to 'am I behind?', plus enough context to explain it."""

    state: str = UNKNOWN
    behind: int = 0
    ahead: int = 0
    branch: str = ''
    upstream: str = ''
    local: str = ''            # short sha of HEAD
    remote: str = ''           # short sha of the upstream tip
    version: str = ''
    commits: list[str] = field(default_factory=list)  # newest first, "sha subject"
    reason: str = ''           # why state is UNKNOWN
    hint: str = ''             # what the user can do about an UNKNOWN
    checked_at: float = 0.0
    stale: bool = False        # loaded from cache, not just measured

    @property
    def update_available(self) -> bool:
        return self.state == BEHIND and self.behind > 0

    def to_dict(self) -> dict:
        d = dict(self.__dict__)
        d.pop('stale', None)
        return d

    @classmethod
    def from_dict(cls, data: dict) -> "UpdateStatus":
        """Build a status from cache JSON, rejecting anything ill-typed.

        The cache is a file on disk that anything can scribble in, and its
        values are used arithmetically (``behind > 0``) and formatted into
        the banner. A string where an int belongs used to crash the wizard
        on every launch, before the menu was even drawn — so a value of the
        wrong type makes the whole file invalid rather than a live landmine.

        Raises ValueError; the caller treats that as "no usable cache".
        """
        known = {}
        for name, field_def in cls.__dataclass_fields__.items():
            if name == 'stale' or name not in data:
                continue
            value = data[name]
            expected = field_def.type
            if expected == 'str':
                ok = isinstance(value, str)
            elif expected == 'int':
                # bool is an int subclass; True is not a commit count.
                ok = isinstance(value, int) and not isinstance(value, bool)
            elif expected == 'float':
                ok = isinstance(value, (int, float)) and not isinstance(value, bool)
                value = float(value) if ok else value
            elif expected.startswith('list'):
                ok = isinstance(value, list) and all(isinstance(v, str) for v in value)
            else:                               # pragma: no cover - defensive
                ok = True
            if not ok:
                raise ValueError(
                    f"cached field {name!r} should be {expected}, got "
                    f"{type(value).__name__}")
            known[name] = value
        return cls(**known)


# ── text from elsewhere ──────────────────────────────────────────────────

def _printable(text: str, limit: int = 200) -> str:
    """Strip control characters from text we did not write.

    Commit subjects and git's own stderr come from the remote, are printed
    by ``--check-update`` and the doctor, and are cached and replayed later.
    An escape sequence in one of them could repaint the report or forge a
    line of it, so nothing that moves the cursor survives this.
    """
    cleaned = ''.join(' ' if (ch < ' ' or '\x7f' <= ch <= '\x9f') else ch
                      for ch in text)
    cleaned = cleaned.strip()
    return cleaned if len(cleaned) <= limit else cleaned[:limit - 1] + '…'


# ── git plumbing ─────────────────────────────────────────────────────────

# Variables that repoint git at a different repository than the one we are
# standing in. Inheriting any of these would silently undo cwd=app_dir() and,
# with it, the "only ever update this clone" guarantee.
_REPO_REDIRECTING_VARS = (
    'GIT_DIR', 'GIT_WORK_TREE', 'GIT_INDEX_FILE', 'GIT_COMMON_DIR',
    'GIT_OBJECT_DIRECTORY', 'GIT_ALTERNATE_OBJECT_DIRECTORIES',
    'GIT_CEILING_DIRECTORIES', 'GIT_NAMESPACE',
)


def _git_env() -> dict:
    """git, pointed at this clone, with every interactive prompt disabled.

    Assignment, not ``setdefault``: an inherited ``GIT_ASKPASS`` is exactly
    the case that hangs — a GUI credential prompt on a machine nobody is
    watching — so ours has to win.
    """
    env = dict(os.environ)
    for var in _REPO_REDIRECTING_VARS:
        env.pop(var, None)
    env['GIT_TERMINAL_PROMPT'] = '0'
    env['GIT_ASKPASS'] = ''
    env['SSH_ASKPASS'] = ''
    env['GCM_INTERACTIVE'] = 'never'
    return env


def _git(*args: str, timeout: int = GIT_TIMEOUT) -> tuple[int, str, str]:
    """Run git in the app directory. Returns (returncode, stdout, stderr).

    Never raises: a missing git binary, a timeout or an OS error all come
    back as a non-zero return code with the reason in stderr.
    """
    try:
        proc = subprocess.run(
            ["git", *args],
            cwd=str(app_dir()),
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace',
            timeout=timeout,
            env=_git_env(),
        )
    except FileNotFoundError:
        return 127, "", "git is not installed (or not on PATH)"
    except subprocess.TimeoutExpired:
        return 124, "", f"git {args[0] if args else ''} timed out after {timeout}s"
    except OSError as e:                       # pragma: no cover - defensive
        return 1, "", str(e)
    return proc.returncode, (proc.stdout or "").strip(), (proc.stderr or "").strip()


def git_available() -> bool:
    return _git("--version")[0] == 0


def _same_dir(a: Path, b: Path) -> bool:
    """Path equality that survives Windows drive-letter case and symlinks."""
    try:
        a, b = a.resolve(), b.resolve()
    except OSError:                             # pragma: no cover - defensive
        pass
    return os.path.normcase(str(a)) == os.path.normcase(str(b))


def work_tree_root() -> Path | None:
    """The root of the git work tree the app directory sits in, if any.

    Asks git rather than looking for a ``.git`` entry, so a worktree (where
    ``.git`` is a file) and a submodule both answer correctly.
    """
    rc, out, _ = _git("rev-parse", "--show-toplevel")
    if rc != 0 or not out:
        return None
    try:
        return Path(out)
    except (OSError, ValueError):               # pragma: no cover - defensive
        return None


def is_git_checkout() -> bool:
    """Is the app directory itself the root of a git clone?

    Being *inside* a work tree is not enough: a ZIP copy unpacked into some
    other project's checkout ("~/projects/thing/tools/MediaOrganizer") is
    inside that project's repository, and updating there would fetch and
    fast-forward the unrelated project while leaving Media Organizer exactly
    as it was.
    """
    root = work_tree_root()
    return root is not None and _same_dir(root, app_dir())


def _current_branch() -> str:
    rc, out, _ = _git("rev-parse", "--abbrev-ref", "HEAD")
    return out if rc == 0 else ''


def _resolve_upstream(branch: str) -> tuple[str, str]:
    """The ref this clone should compare against, as (upstream, reason).

    Prefers the branch's configured upstream. The ``origin/main`` fallback
    covers a clone whose tracking config was never set up, but only while
    the user is actually on ``main`` — assuming any untracked branch follows
    main would fast-forward somebody's work branch onto main and call it an
    update.

    A detached HEAD is refused before either: we could still measure it
    against origin/main, but an update from there would fast-forward a
    nameless HEAD and leave the user's own branch behind, so the honest
    answer is "get back on a branch first".
    """
    if not branch or branch == 'HEAD':
        return '', "not on a branch (detached HEAD)"
    rc, out, _ = _git("rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{upstream}")
    if rc == 0 and out and out != '@{upstream}':
        return out, ''
    fallback = f"{DEFAULT_REMOTE}/{DEFAULT_BRANCH}"
    if branch == DEFAULT_BRANCH and _git("rev-parse", "--verify", "--quiet",
                                         fallback)[0] == 0:
        return fallback, ''
    return '', f"branch '{branch}' does not track a remote branch"


def split_upstream(upstream: str) -> tuple[str, str] | None:
    """An upstream ref as (remote, branch), or None if it names no remote.

    ``@{upstream}`` can point at another *local* branch. Splitting that on
    '/' yields an empty branch name and a `git pull <local-branch> ""`, so
    the honest answer is "this is not something we can pull from".
    """
    remote, sep, branch = upstream.partition('/')
    if not sep or not remote or not branch:
        return None
    return remote, branch


def _short(rev: str) -> str:
    rc, out, _ = _git("rev-parse", "--short", rev)
    return out if rc == 0 else ''


def head_revision() -> str:
    """The commit currently checked out, short form ('' if not a clone).

    The only honest answer to "did the files on disk just change?" — take it
    either side of an update rather than inferring from a status measured
    before the fetch, which can be reading refs that were already stale.
    """
    return _short('HEAD')


def local_commits_behind(upstream: str) -> tuple[int, int] | None:
    """(ahead, behind) between HEAD and *upstream*, or None if git could not say.

    None rather than (0, 0): a comparison that *failed* is not the same as a
    comparison that found no difference, and reporting the first as the
    second means a broken clone shows "up to date" and never offers the
    update that would fix it.
    """
    rc, out, _ = _git("rev-list", "--left-right", "--count", f"HEAD...{upstream}")
    if rc != 0:
        return None
    parts = out.split()
    if len(parts) != 2:
        return None
    try:
        return int(parts[0]), int(parts[1])
    except ValueError:                          # pragma: no cover - defensive
        return None


def incoming_commits(upstream: str, limit: int = 20) -> list[str]:
    """Subjects of the commits we do not have yet, newest first."""
    rc, out, _ = _git("log", f"--max-count={limit}", "--pretty=format:%h %s",
                      f"HEAD..{upstream}")
    return [_printable(line) for line in out.splitlines()] if rc == 0 and out else []


def working_tree_changes() -> list[str]:
    """Tracked files modified locally. Untracked files are not our business —
    the journal, spreadsheets and config files all live here."""
    rc, out, _ = _git("status", "--porcelain", "--untracked-files=no")
    return out.splitlines() if rc == 0 and out else []


# ── checking ─────────────────────────────────────────────────────────────

def _unknown(reason: str, hint: str = '') -> UpdateStatus:
    return UpdateStatus(state=UNKNOWN, reason=reason, hint=hint,
                        version=_version(), checked_at=time.time())


def check(*, fetch: bool = True) -> UpdateStatus:
    """Measure this clone against the branch it tracks.

    With ``fetch=False`` the comparison uses whatever the last fetch left in
    ``origin/main``, so it is instant and offline — but it can only report an
    update someone already fetched.
    """
    if not git_available():
        return _unknown(
            "git is not installed",
            "Install git from https://git-scm.com/downloads, or re-download "
            f"the project from {REPO_URL}")
    if not is_git_checkout():
        enclosing = work_tree_root()
        reason = "this copy was not installed with 'git clone'"
        if enclosing is not None:
            # Inside someone else's repository. Say so explicitly — silently
            # updating that repository instead is the one outcome nobody wants.
            reason += f" (it sits inside another repository: {enclosing})"
        return _unknown(
            reason,
            "Automatic updates need a git clone. Re-download with:\n"
            f"    git clone {REPO_URL}.git\n"
            "Your journal, spreadsheets and settings live outside the repo, "
            "so nothing is lost.")

    branch = _current_branch()
    upstream, reason = _resolve_upstream(branch)
    if not upstream:
        hint = f"Switch to the tracked branch: git checkout {DEFAULT_BRANCH}"
        if branch and branch != 'HEAD':
            hint += ("\nor tell git what this branch follows:\n"
                     f"    git branch --set-upstream-to={DEFAULT_REMOTE}/{branch}")
        return _unknown(reason, hint)

    parts = split_upstream(upstream)
    if parts is None:
        return _unknown(
            f"branch '{branch}' tracks '{upstream}', which is not a remote branch",
            f"Point it at the remote: "
            f"git branch --set-upstream-to={DEFAULT_REMOTE}/{DEFAULT_BRANCH}")
    remote, _remote_branch = parts

    if fetch:
        rc, _, err = _git("fetch", "--quiet", remote, timeout=FETCH_TIMEOUT)
        if rc != 0:
            return _unknown(
                "could not reach the remote: "
                f"{_printable(err) or 'git fetch failed'}",
                "Check your internet connection, then try again.")

    counts = local_commits_behind(upstream)
    if counts is None:
        return _unknown(
            f"git could not compare this clone with {upstream}",
            "Run 'python run.py --doctor' — the clone may be damaged.")
    ahead, behind = counts
    if behind and ahead:
        state = DIVERGED
    elif behind:
        state = BEHIND
    elif ahead:
        state = AHEAD
    else:
        state = CURRENT

    return UpdateStatus(
        state=state, behind=behind, ahead=ahead, branch=branch,
        upstream=upstream, local=_short('HEAD'), remote=_short(upstream),
        version=_version(),
        commits=incoming_commits(upstream) if behind else [],
        checked_at=time.time(),
    )


# ── cache ────────────────────────────────────────────────────────────────

def cache_path() -> Path:
    return app_dir() / CACHE_NAME


def load_cache() -> UpdateStatus | None:
    try:
        with open(cache_path(), encoding='utf-8') as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError, ValueError):
        return None
    if not isinstance(data, dict):
        return None
    try:
        st = UpdateStatus.from_dict(data)
    except (TypeError, ValueError):
        # Ill-typed or unusable. Drop the file so the next check writes a
        # good one instead of failing the same way on every future launch.
        try:
            cache_path().unlink()
        except OSError:
            pass
        return None
    st.stale = True
    return st


def save_cache(status: UpdateStatus) -> None:
    """Best effort — an unwritable app directory must not break the app.

    An UNKNOWN is never written. "I could not tell" carries no information
    worth remembering, and because it records no commit or branch it would
    pass :func:`describes_current_checkout` and suppress the real check for
    the rest of the interval.
    """
    if status.state == UNKNOWN:
        return
    try:
        with open(cache_path(), 'w', encoding='utf-8') as f:
            json.dump(status.to_dict(), f, indent=2)
    except OSError:
        pass


def describes_current_checkout(status: UpdateStatus) -> bool:
    """Is this cached answer about the commit and branch running right now?

    Two cheap local git calls, no network. Without this, a manual
    ``git pull`` leaves "3 commits behind" on screen for the rest of the
    day — and, worse, a ``git checkout`` of an older commit is reported as
    up to date.
    """
    if not status.local and not status.branch:
        return True             # nothing recorded to contradict (old cache)
    if status.local and status.local != head_revision():
        return False
    if status.branch and status.branch != _current_branch():
        return False
    return True


def checks_disabled() -> bool:
    """MEDIAORG_NO_UPDATE_CHECK=1 turns the startup check off entirely."""
    return os.environ.get('MEDIAORG_NO_UPDATE_CHECK', '').strip().lower() \
        in ('1', 'true', 'yes', 'on')


def check_interval_hours() -> float:
    raw = os.environ.get('MEDIAORG_UPDATE_INTERVAL', '').strip()
    try:
        val = float(raw)
    except ValueError:
        return DEFAULT_INTERVAL_HOURS
    return val if val >= 0 else DEFAULT_INTERVAL_HOURS


# ── background check (used by the wizard) ────────────────────────────────

_lock = threading.Lock()
_status: UpdateStatus | None = None
_started = False
_thread: threading.Thread | None = None
# Set once the thread has published whatever it remembered, before it goes
# anywhere near the network.
_published = threading.Event()


def check_and_cache(*, fetch: bool = True) -> UpdateStatus:
    """A check whose answer primes the startup banner.

    Used by the explicit commands (``--check-update``, the doctor) so that
    asking once means the next launch can answer without the network.
    """
    st = check(fetch=fetch)
    if st.state != UNKNOWN:
        save_cache(st)
        with _lock:
            global _status
            _status = st
    return st


def begin_background_check(*, force: bool = False) -> None:
    """Start the update check. Does no work on the calling thread.

    Everything — reading the cache, checking it still describes this
    checkout (two git subprocesses), and any network call — happens on a
    daemon thread, so launching the app is never slower for having asked.
    Callers who want the answer before drawing something use
    :func:`wait_for_cache` (fast, local) or :func:`wait_for_check` (may
    include the network).
    """
    global _started, _thread
    if checks_disabled():
        return
    with _lock:
        if _started and not force:
            return
        _started = True
        _published.clear()
        thread = threading.Thread(target=_refresh, args=(force,),
                                  name="mediaorg-update-check", daemon=True)
        _thread = thread
    # Started outside the lock: a thread that finishes instantly would
    # otherwise block on a lock its own starter is still holding.
    thread.start()


def _refresh(force: bool = False) -> None:
    """Publish what we remember, then re-measure if that is not enough."""
    global _status
    try:
        cached = load_cache()
        if cached is not None and not describes_current_checkout(cached):
            # The checkout moved under us — a manual `git pull`, a `checkout`,
            # a `reset`. The cached answer is about a commit that is no longer
            # the one running, so it is not merely old, it is wrong: don't
            # publish it even briefly, and re-measure regardless of its age.
            cached = None
        with _lock:
            _status = cached
    except Exception:                           # pragma: no cover - defensive
        cached = None
    finally:
        _published.set()

    interval = check_interval_hours()
    fresh = (cached is not None and interval > 0
             and (time.time() - cached.checked_at) < interval * 3600)
    if fresh and not force:
        return

    try:
        st = check(fetch=True)
    except Exception:                           # pragma: no cover - defensive
        return
    # Don't cache a transient "no network" over a real answer: keep the last
    # useful result so an offline launch still reports what it knew.
    if st.state != UNKNOWN:
        save_cache(st)
        with _lock:
            _status = st
    else:
        with _lock:
            if _status is None:
                _status = st


def wait_for_cache(timeout: float) -> UpdateStatus | None:
    """Wait for the remembered answer only — local work, no network.

    This is the one every launch does: it costs a couple of git calls at
    worst and lets the notice appear on the first menu draw rather than the
    second. With no check running — checks disabled, or none started — there
    is nothing to wait for and it returns at once.
    """
    with _lock:
        running = _thread is not None
    if running:
        _published.wait(timeout)
    return latest_status()


def wait_for_check(timeout: float) -> UpdateStatus | None:
    """Give a running check up to *timeout* seconds to finish, network included.

    Only worth calling when there is no cached answer at all — the first
    launch after install, where waiting a moment is the difference between
    the notice appearing now and appearing next time. A dead network costs
    the timeout, never more.
    """
    with _lock:
        thread = _thread
    if thread is not None and thread.is_alive():
        thread.join(timeout)
    return latest_status()


def latest_status() -> UpdateStatus | None:
    with _lock:
        return _status


def join_background(timeout: float = 30.0) -> bool:
    """Wait for the check thread to exit. Returns True if none is running.

    Tests use this to stop a daemon thread outliving the fixture that
    pointed it at a scratch repository.
    """
    with _lock:
        thread = _thread
    if thread is None:
        return True
    thread.join(timeout)
    return not thread.is_alive()


def reset_background_state() -> None:
    """Test hook: forget the thread and the published status."""
    global _started, _status, _thread
    with _lock:
        _started = False
        _status = None
        _thread = None
        _published.clear()


# ── presentation ─────────────────────────────────────────────────────────

def update_command() -> str:
    """The exact command to type, including the folder to type it in."""
    return f"cd \"{app_dir()}\"\n    python run.py --update"


def banner(status: UpdateStatus | None = None) -> str:
    """The startup notice, or '' when there is nothing to say.

    Deliberately silent unless an update actually exists: a tool that shouts
    on every launch gets ignored on the launch that mattered.
    """
    st = status if status is not None else latest_status()
    if st is None or not st.update_available:
        return ''
    plural = '' if st.behind == 1 else 's'
    lines = [
        "",
        "-" * 70,
        f"  Update available: you are {st.behind} commit{plural} behind "
        f"{st.upstream}.",
        f"  (you have {st.local or '?'}, latest is {st.remote or '?'})",
        "",
        "  To update, run this in a terminal in the MediaOrganizer folder:",
        f"      cd \"{app_dir()}\"",
        "      python run.py --update",
        "",
        "  ...or press [U] here and the wizard will do it for you.",
        "",
        # The only place this is discoverable in the UI: a media server on a
        # restricted network should not have to read the source to find out
        # that launching the app contacts github.com.
        "  (This check contacts github.com once a day. Turn it off with the",
        "   environment variable MEDIAORG_NO_UPDATE_CHECK=1.)",
        "-" * 70,
    ]
    return "\n".join(lines)


def describe(status: UpdateStatus) -> str:
    """Full report for ``--check-update`` / the doctor."""
    st = status
    out = [f"Media Organizer v{st.version or _version()}"]
    if st.local:
        out.append(f"Installed commit: {st.local}"
                   + (f" (branch {st.branch})" if st.branch else ""))
    if st.state == UNKNOWN:
        out.append(f"[!] Could not check for updates: {st.reason}")
        if st.hint:
            out.extend("    " + line for line in st.hint.split("\n"))
        return "\n".join(out)

    age = ""
    if st.stale and st.checked_at:
        hours = max(0.0, (time.time() - st.checked_at) / 3600)
        age = (f" (as of {hours:.0f}h ago)" if hours >= 1
               else " (checked just now)")

    if st.state == CURRENT:
        out.append(f"[OK] Up to date with {st.upstream}{age}.")
    elif st.state == BEHIND:
        plural = '' if st.behind == 1 else 's'
        out.append(f"[!] {st.behind} commit{plural} behind {st.upstream}{age}"
                   f" — latest is {st.remote}.")
        if st.commits:
            out.append("")
            out.append("What you are missing:")
            out.extend(f"  {c}" for c in st.commits)
        out.append("")
        out.append("To update, run this in the MediaOrganizer folder:")
        out.append(f"    cd \"{app_dir()}\"")
        out.append("    python run.py --update")
    elif st.state == AHEAD:
        plural = '' if st.ahead == 1 else 's'
        out.append(f"[OK] Up to date, plus {st.ahead} local commit{plural} "
                   f"not on {st.upstream}.")
    elif st.state == DIVERGED:
        out.append(f"[!] Diverged from {st.upstream}: {st.ahead} local "
                   f"commit(s), {st.behind} remote commit(s).")
        out.append("    A plain update cannot fast-forward past your own "
                   "commits. Either")
        out.append("    push/keep them on a branch, or discard them with:")
        out.append(f"        git reset --hard {st.upstream}")
    return "\n".join(out)


# ── updating ─────────────────────────────────────────────────────────────

def _can_prompt() -> bool:
    """Is there a terminal on the other end to answer a question?"""
    try:
        return sys.stdin is not None and sys.stdin.isatty()
    except (AttributeError, ValueError):        # closed or replaced stream
        return False


def _in_virtualenv() -> bool:
    return sys.prefix != getattr(sys, 'base_prefix', sys.prefix)


def _pip_install(req: Path, timeout: int = PIP_TIMEOUT) -> bool:
    """Install requirements, retrying with --user on a permissions failure.

    stdin is closed and the call is bounded: pip can ask questions, and an
    unattended ``--update --yes`` that stops for one waits forever. ``--user``
    is skipped inside a virtualenv, where pip rejects it outright, and on
    PEP 668 installs it is the retry that has a chance of working.
    """
    attempts = [[]] if _in_virtualenv() else [[], ["--user"]]
    for extra in attempts:
        try:
            subprocess.run(
                [sys.executable, "-m", "pip", "install", *extra,
                 "-r", str(req), "--quiet", "--disable-pip-version-check"],
                check=True, stdin=subprocess.DEVNULL, timeout=timeout,
            )
            return True
        except subprocess.TimeoutExpired:
            print(f"[!] pip install exceeded {timeout}s and was stopped.")
            return False
        except (subprocess.CalledProcessError, OSError):
            continue
    return False


def _requirements_fingerprint() -> str | None:
    """The contents of requirements.txt, or None if there is no such file.

    Compared either side of the pull to decide whether dependencies need
    reinstalling. Reading the file beats diffing a commit range: it needs no
    successful `git diff`, and it gets a rename or a deletion right.
    """
    try:
        return (app_dir() / "requirements.txt").read_text(encoding='utf-8')
    except OSError:
        return None


def run_update(*, assume_yes: bool = False, dry_run: bool = False) -> int:
    """Fast-forward this clone onto its tracked branch. Returns an exit code.

    0 = updated or already current, 1 = could not update (with the reason and
    the commands to fix it printed).
    """
    print("=" * 70)
    print("Media Organizer - Update")
    print("=" * 70)

    st = check(fetch=True)
    if st.state == UNKNOWN:
        print(describe(st))
        return 1
    if st.state == AHEAD or st.state == CURRENT:
        print(describe(st))
        if st.state == CURRENT:
            print("\nNothing to do.")
        return 0
    if st.state == DIVERGED:
        print(describe(st))
        return 1

    print(describe(st))

    dirty = working_tree_changes()
    if dirty:
        print("\n[!] You have local changes to files the update would "
              "overwrite:")
        for line in dirty[:20]:
            print(f"      {line}")
        if len(dirty) > 20:
            print(f"      ... and {len(dirty) - 20} more")
        # `git checkout -- .` restores the work tree from the index, so it
        # leaves *staged* changes exactly where they were — and those show up
        # in `git status --porcelain` too. Advising it would send the user
        # round the loop to the identical refusal.
        print("\n    Keep them:       git stash")
        print("    Throw them away: git reset --hard")
        print("    Then run:        python run.py --update")
        return 1

    parts = split_upstream(st.upstream)
    if parts is None:                           # pragma: no cover - check() guards
        print(f"\n[!] {st.upstream} does not name a remote branch to pull from.")
        return 1
    remote, branch = parts

    if dry_run:
        print(f"\n[dry-run] Would run: git pull --ff-only {remote} {branch}")
        return 0

    if not assume_yes:
        if not _can_prompt():
            # Every other flow in this app confirms before it changes
            # anything. Silently proceeding because nobody happens to be
            # attached would make --yes decorative rather than the opt-in it
            # is documented to be.
            print("\n[!] Nothing is attached to answer the confirmation "
                  "(stdin is not a terminal).")
            print("    Re-run and opt in explicitly:")
            print("        python run.py --update --yes")
            print("    ...or see what it would do first:")
            print("        python run.py --update --dry-run")
            return 1
        try:
            answer = input("\nUpdate now? (y/n) [y]: ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            # Ctrl-C / Ctrl-D at a confirmation means "no", not a traceback.
            print("\nCancelled - nothing was changed.")
            return 0
        if answer and answer not in ('y', 'yes'):
            print("Cancelled - nothing was changed.")
            return 0

    before = head_revision()
    requirements_before = _requirements_fingerprint()
    print(f"\n-> git pull --ff-only {remote} {branch}")
    rc, out, err = _git("pull", "--ff-only", remote, branch,
                        timeout=FETCH_TIMEOUT * 3)
    if out:
        print(_printable(out, limit=4000))
    if rc != 0:
        print(f"\n[!] Update failed: "
              f"{_printable(err) or 'git pull returned ' + str(rc)}")
        print("    Try manually, from the MediaOrganizer folder:")
        print(f"        cd \"{app_dir()}\"")
        print(f"        git pull --ff-only {remote} {branch}")
        return 1

    after = head_revision()

    # Only reinstall when the dependency list actually moved: pip is slow and
    # this would otherwise run on every update.
    deps_failed = False
    requirements_after = _requirements_fingerprint()
    if requirements_after is not None and requirements_after != requirements_before:
        req = app_dir() / "requirements.txt"
        print("\n-> requirements.txt changed, updating dependencies...")
        if _pip_install(req):
            print("[OK] Dependencies updated.")
        else:
            deps_failed = True
            print("[!] Dependency install failed. Run it yourself:")
            print(f"        python -m pip install -r \"{req}\"")

    save_cache(check(fetch=False))
    print("\n" + "=" * 70)
    if before and after and before == after:    # pragma: no cover - defensive
        print("[OK] Already up to date - nothing changed.")
        print("=" * 70)
        return 0
    moved = f"{before} -> {after}" if before and after else (after or "the latest commit")
    if deps_failed:
        # Not "[OK]": the code is new but its dependencies are not, which is
        # exactly the state that fails confusingly on the next launch.
        print(f"[!] Updated the code ({moved}), but the dependency install "
              f"failed.")
        print("     Finish it with the pip command above, then restart:")
    else:
        print(f"[OK] Updated {moved}.")
        print("     Restart Media Organizer to use the new version:")
    print("         python run.py")
    print("=" * 70)
    return 1 if deps_failed else 0


#: Flags run.py handles before it imports anything heavy.
CLI_FLAGS = ('--update', '--check-update', '--version', '-V')


def dispatch_cli(argv: list[str]) -> int | None:
    """Handle the update/version flags, or return None if none were given.

    Lives here rather than in run.py so there is one parser for them instead
    of a hand-rolled ``in sys.argv`` test that drifts from the real one.
    Abbreviations are off, so ``--upd`` is an error in both parsers rather
    than being ignored by one and honoured by the other.
    """
    if not any(flag in argv for flag in CLI_FLAGS):
        return None
    parser = argparse.ArgumentParser(prog="run.py", add_help=False,
                                     allow_abbrev=False)
    parser.add_argument('--update', action='store_true')
    parser.add_argument('--check-update', action='store_true')
    parser.add_argument('--version', '-V', action='store_true')
    parser.add_argument('--yes', '-y', action='store_true')
    parser.add_argument('--dry-run', action='store_true')
    args, extra = parser.parse_known_args(argv)
    if extra:
        print(f"[!] Unrecognized argument(s): {' '.join(extra)}")
        print(f"    Usage: python run.py [{' | '.join(CLI_FLAGS[:3])}] "
              f"[--yes] [--dry-run]")
        return 2
    if args.version:
        print_version()
        return 0
    if args.check_update:
        print(describe(check_and_cache(fetch=True)))
        return 0
    return run_update(assume_yes=args.yes, dry_run=args.dry_run)


def print_version() -> None:
    line = f"Media Organizer v{_version()}"
    if is_git_checkout():
        sha, branch = _short('HEAD'), _current_branch()
        if sha:
            line += f" ({sha}" + (f" on {branch}" if branch else "") + ")"
    print(line)
