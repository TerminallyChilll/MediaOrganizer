"""Execute plans with a crash-safe JSONL journal; undo by reverse replay.

Two invariants make undo trustworthy:

1. **Intent before action.** An ``intent`` entry is written and fsync'd
   *before* each mutation, and a completion entry after it. Undo replays only
   completions, so its behaviour is unchanged — but an ``intent`` with no
   completion is a crash scar, and :func:`recover` can clean it up (a partial
   cross-device copy, or a stranded ``.mediaorg_tmp`` from a case-only rename).
   Without this, a mutation that died halfway was invisible to undo *and*
   permanently blocked retries, because the collision check would see the
   orphan and skip forever.

2. **The journal records the paths actually used.** Operational paths are made
   absolute lexically (:func:`_op_path`) and never ``resolve()``-d. Resolving
   was previously the last step of path validation, which on Windows
   canonicalises to the on-disk casing — silently breaking every case-only
   rename ("wall-e" -> "WALL-E") and rewriting mapped drives to UNC. Symlink
   containment is now checked separately, against the plan's own roots.
"""

import errno
import json
import os
import shutil
import stat
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path

from .parse import TMP_SUFFIX, is_junk_dir, is_junk_name
from .plan import Op, Plan

JOURNAL_FILE = "mediaorg_journal.jsonl"
JOURNAL_VERSION = 2
TRASH_DIR = ".mediaorg_trash"

# Backoff for files held open by another program (Plex/Jellyfin/VLC streaming
# an episode is the single most common real-world failure). Tests set this to
# () to keep the suite fast.
RETRY_DELAYS = (0.5, 1.5, 3.0)


# --- Journal location -------------------------------------------------------

def journal_path() -> Path:
    """Where the journal lives, independent of the current directory.

    A cwd-relative journal meant that launching from anywhere else — double
    clicking the Windows launcher, or `cd`-ing elsewhere — silently started a
    fresh journal and reported "nothing to undo" for real changes.

    Order: ``$MEDIAORG_JOURNAL`` -> next to the app -> adopt a pre-existing
    journal in the cwd (so upgrading users keep their undo history; we point
    at it rather than moving it behind their back).
    """
    env = os.environ.get("MEDIAORG_JOURNAL")
    if env:
        return Path(env).expanduser()
    app = Path(__file__).resolve().parent.parent / JOURNAL_FILE
    if not app.exists():
        legacy = Path.cwd() / JOURNAL_FILE
        if legacy.exists() and legacy != app:
            return legacy
    return app


@dataclass
class ExecResult:
    done: list[Op] = field(default_factory=list)
    failed: list[tuple[Op, str]] = field(default_factory=list)
    dry_run: bool = False
    # The journal run this result came from, so a caller can reverse exactly
    # what it just did rather than guessing at "the newest pending run" — which
    # is a different run the moment anything else writes to the journal first.
    # None whenever no run was opened: a dry run, a plan with nothing to do,
    # and undo results. Callers must check before passing it to undo_run().
    run_id: str | None = None

    @property
    def ok(self) -> bool:
        return not self.failed


# --- Path handling ----------------------------------------------------------

def _op_path(p: Path) -> Path:
    """The path an operation actually uses: lexically absolute, case intact.

    Deliberately NOT ``resolve()`` — see the module docstring.
    """
    return Path(os.path.abspath(p))


def _resolve_for_check(p: Path) -> Path:
    try:
        return Path(os.path.abspath(p)).resolve()
    except OSError:
        return Path(os.path.abspath(p))


def _resolved_roots(roots) -> list[Path]:
    out = []
    for r in roots or ():
        out.append(_resolve_for_check(Path(r)))
    return out


def _implicit_roots(ops: list[Op]) -> list[Path]:
    """Containment bound derived from the plan itself.

    A plan must never touch anything outside the tree it was planned for, so
    the common ancestor of every op's directory is a sound default root. This
    replaces the old blanket "reject any symlinked path component" rule, which
    refused legitimate layouts wholesale: Windows NTFS junctions and volume
    mount points, macOS `/tmp` and `/var` (both symlinks), and Linux
    symlink-farm / mergerfs media trees.
    """
    dirs: list[str] = []
    for op in ops:
        if op.src:
            dirs.append(str(_op_path(op.src).parent))
        dirs.append(str(_op_path(op.dst).parent))
    if not dirs:
        return []
    try:
        common = os.path.commonpath(dirs)
    except ValueError:
        return []  # mixed drives — fall back to '..'-rejection only
    return [_resolve_for_check(Path(common))]


def _validate_op_path(p: Path, roots: list[Path]) -> None:
    """Reject traversal and any symlink that escapes the allowed roots."""
    if '..' in p.parts:
        raise ValueError(f"path traversal rejected: {p}")
    if not roots:
        return
    resolved = _resolve_for_check(p)
    for root in roots:
        try:
            if resolved == root or resolved.is_relative_to(root):
                return
        except (OSError, ValueError):
            continue
    raise ValueError(f"path escapes the library root: {p} (resolves to {resolved})")


def _same_file(a: Path, b: Path) -> bool:
    try:
        return a.samefile(b)
    except OSError:
        return False


def _missing_dirs(dst: Path) -> list[Path]:
    """`dst` and its ancestors that don't exist yet, shallowest first.

    The loop must terminate on a self-parent, not only on ``exists()``: a
    Windows filesystem root is its own parent (``PureWindowsPath('Z:/').parent
    == 'Z:\\'``) and can be non-existent, so an unavailable mapped drive or
    unreachable UNC share used to spin here forever, before any error handling
    could see it. ``Path('.')`` behaves the same way if the cwd is deleted.
    """
    missing: list[Path] = []
    p = dst
    while not p.exists():
        missing.append(p)
        parent = p.parent
        if parent == p:
            break
        p = parent
    return list(reversed(missing))


def _missing_parents(dst: Path) -> list[Path]:
    """Ancestors of dst that don't exist yet, shallowest first."""
    return _missing_dirs(dst.parent)


# --- Moving -----------------------------------------------------------------

def _is_case_only_rename(src: Path, dst: Path) -> bool:
    """Same directory, same name apart from letter case.

    The old test was ``samefile(src, dst)``, which is also true for
    *hardlinks* — so a hardlink pair went down the two-step rename path and
    POSIX ``os.rename`` destroyed one of the two directory entries, with
    nothing journaled to undo it.
    """
    return (src.parent == dst.parent
            and src.name != dst.name
            and src.name.casefold() == dst.name.casefold())


def _free_tmp_path(src: Path, limit: int = 1000) -> Path:
    candidate = src.with_name(src.name + TMP_SUFFIX)
    if not os.path.lexists(candidate):
        return candidate
    for n in range(1, limit + 1):
        candidate = src.with_name(f"{src.name}{TMP_SUFFIX}.{n}")
        if not os.path.lexists(candidate):
            return candidate
    raise OSError(errno.EEXIST,
                  f"no free temp name for {src} after {limit} attempts")


def _fsync_path(path: Path) -> None:
    try:
        fd = os.open(path, os.O_RDONLY)
    except OSError:
        return
    try:
        os.fsync(fd)
    except OSError:
        pass  # some filesystems/handles refuse; the copy is still verified
    finally:
        os.close(fd)


def _move_across(src: Path, dst: Path) -> None:
    """Rename, falling back to a *verified* copy when crossing devices.

    ``shutil.move``'s fallback leaves a truncated destination if the copy dies
    partway (disk full, network share drops, USB unplugged), which then blocks
    every future retry. Here the size is checked before the source is dropped,
    and the caller's ``intent`` journal entry lets :func:`recover` remove the
    partial file if we never get that far.
    """
    try:
        os.rename(src, dst)
        return
    except OSError as exc:
        if exc.errno != errno.EXDEV:
            raise

    if src.is_dir() and not src.is_symlink():
        shutil.copytree(src, dst, symlinks=True)
        shutil.rmtree(src)
        return

    # lstat on both sides: copy2 is told not to follow symlinks, so for a
    # symlinked file it copies the link itself. stat() would compare the
    # target's size against the link's and fail a perfectly good copy.
    expected = src.lstat().st_size
    shutil.copy2(src, dst, follow_symlinks=False)
    _fsync_path(dst)
    actual = dst.lstat().st_size
    if actual != expected:
        raise OSError(errno.EIO,
                      f"copy verification failed: {dst} is {actual} bytes, "
                      f"expected {expected}")
    os.unlink(src)


def _do_move(src: Path, dst: Path, on_tmp=None) -> None:
    if _is_case_only_rename(src, dst) and (
            not os.path.lexists(dst) or _same_file(src, dst)):
        # A case-insensitive filesystem sees src and dst as the same name, so
        # go through a temp name. Journal the temp name first (via on_tmp) so a
        # crash between the two renames is recoverable, and roll back if the
        # second rename fails so we never strand an orphan.
        tmp = _free_tmp_path(src)
        if on_tmp is not None:
            on_tmp(tmp)
        os.rename(src, tmp)
        try:
            os.rename(tmp, dst)
        except OSError:
            os.rename(tmp, src)
            raise
        return
    if os.path.lexists(dst):
        raise FileExistsError(f"target already exists: {dst}")
    dst.parent.mkdir(parents=True, exist_ok=True)
    _move_across(src, dst)


def _retrying(fn, delays=None):
    """Retry a filesystem action while another program holds the file open."""
    delays = RETRY_DELAYS if delays is None else delays
    for delay in (*delays, None):
        try:
            return fn()
        except PermissionError:
            if delay is None:
                raise
            time.sleep(delay)


def _describe(exc: BaseException) -> str:
    if isinstance(exc, PermissionError):
        return (f"{exc} - the file may be open in another program "
                f"(close Plex/Jellyfin/VLC and retry) or you lack permission")
    return str(exc)


# --- Junk quarantine --------------------------------------------------------

def _junk_children(directory: Path) -> list[Path]:
    """OS metadata entries that would make an otherwise-empty rmdir fail.

    Junk *directories* count too: a folder holding only Synology's `@eaDir`
    still fails rmdir, and treating any subdirectory as "real content" left
    those stuck forever.
    """
    try:
        entries = list(os.scandir(directory))
    except OSError:
        return []
    junk = []
    for entry in entries:
        if entry.is_dir(follow_symlinks=False):
            if not is_junk_dir(entry.name):
                return []  # real content: not "empty but for junk"
        elif not is_junk_name(entry.name):
            return []
        junk.append(Path(entry.path))
    return sorted(junk)


def _trash_dir(target: Path, roots: list[Path]) -> Path:
    for root in roots:
        try:
            if target == root or _resolve_for_check(target).is_relative_to(root):
                return Path(os.path.abspath(root)) / TRASH_DIR
        except (OSError, ValueError):
            continue
    return target.parent / TRASH_DIR


def _free_name(path: Path, limit: int = 1000) -> Path:
    """An unused name near `path`. Bounded so it can never spin forever."""
    if not os.path.lexists(path):
        return path
    for n in range(1, limit + 1):
        candidate = path.with_name(f"{path.stem}.{n}{path.suffix}")
        if not os.path.lexists(candidate):
            return candidate
    raise OSError(errno.EEXIST,
                  f"no free name for {path} after {limit} attempts")


# --- Journal writing --------------------------------------------------------

class _Writer:
    """Append-only journal writer. Every line is flushed and fsync'd."""

    def __init__(self, handle):
        self._h = handle
        self.seq = 0

    def log(self, entry: dict) -> None:
        entry.setdefault("ts", time.time())
        self._h.write(json.dumps(entry, ensure_ascii=False) + "\n")
        self._h.flush()
        os.fsync(self._h.fileno())


def _stat_fields(path: Path) -> dict:
    """Identity of the file we just placed, so undo can detect replacement.

    Directories deliberately get no size/mtime: a directory's mtime changes
    whenever its contents do, which happens legitimately later in the same run
    (renaming the episodes inside a season folder) — recording it would make
    undo refuse to reverse the folder rename.
    """
    try:
        st = os.stat(path, follow_symlinks=False)
    except OSError:
        return {}
    if stat.S_ISDIR(st.st_mode):
        return {"dir": True}
    return {"size": st.st_size, "mtime": round(st.st_mtime, 3)}


def execute(plan: Plan, journal: Path, dry_run: bool = False, *,
            roots=None, label: str | None = None,
            session: str | None = None) -> ExecResult:
    """Apply plan ops in order, journaling each success immediately.

    `roots` bounds where ops may land; when omitted it is derived from the
    plan itself (see :func:`_implicit_roots`).
    """
    result = ExecResult(dry_run=dry_run)
    if dry_run or not plan.ops:
        result.done = list(plan.ops)
        return result

    check_roots = _resolved_roots(roots) or _implicit_roots(plan.ops)
    run_id = uuid.uuid4().hex[:12]
    # Stamped before the loop, not after it: a run whose ops all failed still
    # opened a run in the journal, and the caller still needs to name it.
    result.run_id = run_id

    with open(journal, "a", encoding="utf-8") as handle:
        jw = _Writer(handle)
        jw.log({"op": "begin_run", "id": run_id, "v": JOURNAL_VERSION,
                "session": session or run_id, "label": label})

        for op in plan.ops:
            jw.seq += 1
            seq = jw.seq
            created: list[Path] = []
            quarantined: list[tuple[Path, Path]] = []
            try:
                if op.kind == "move":
                    src, dst = _op_path(op.src), _op_path(op.dst)
                    _validate_op_path(op.src, check_roots)
                    _validate_op_path(op.dst, check_roots)
                    created = _missing_parents(dst)

                    # Whether the destination was already occupied BEFORE we
                    # touched anything. The intent has to be journaled before
                    # the mutation, which means it is also written when the
                    # move then fails on a pre-existing destination — and
                    # without this flag recovery would mistake that innocent
                    # file for a partial copy of ours and delete it.
                    dst_existed = os.path.lexists(dst)

                    def on_tmp(tmp: Path, _seq=seq, _src=src, _dst=dst,
                               _existed=dst_existed) -> None:
                        jw.log({"op": "intent", "kind": "move", "seq": _seq,
                                "src": str(_src), "dst": str(_dst),
                                "tmp": str(tmp), "dst_existed": _existed})

                    jw.log({"op": "intent", "kind": "move", "seq": seq,
                            "src": str(src), "dst": str(dst),
                            "dst_existed": dst_existed})
                    _retrying(lambda: _do_move(src, dst, on_tmp=on_tmp))
                    entry = {"op": "move", "src": str(src), "dst": str(dst),
                             "seq": seq, **_stat_fields(dst)}

                elif op.kind == "mkdir":
                    dst = _op_path(op.dst)
                    _validate_op_path(op.dst, check_roots)
                    # Journal every directory we create, not just the leaf, or
                    # undo leaves the intermediate ones behind.
                    created = _missing_dirs(dst)[:-1]
                    jw.log({"op": "intent", "kind": "mkdir", "seq": seq,
                            "dst": str(dst)})
                    dst.mkdir(parents=True, exist_ok=True)
                    entry = {"op": "mkdir", "src": None, "dst": str(dst),
                             "seq": seq}

                elif op.kind == "rmdir":
                    dst = _op_path(op.dst)
                    _validate_op_path(op.dst, check_roots)
                    # A leftover .DS_Store / Thumbs.db makes rmdir fail, which
                    # used to leave every organize run reporting failures and
                    # undo permanently un-completable. Move the junk aside as
                    # ordinary journaled moves so the whole thing stays
                    # reversible - no delete primitive is introduced.
                    junk = _junk_children(dst)
                    if junk:
                        trash = _trash_dir(dst, check_roots)
                        trash_created = _missing_dirs(trash)
                        trash.mkdir(parents=True, exist_ok=True)
                        for d in trash_created:
                            jw.log({"op": "mkdir", "src": None, "dst": str(d),
                                    "seq": seq})
                        for j in junk:
                            jdst = _free_name(trash / j.name)
                            _retrying(lambda a=j, b=jdst: os.rename(a, b))
                            jw.log({"op": "move", "src": str(j),
                                    "dst": str(jdst), "seq": seq,
                                    **_stat_fields(jdst)})
                            quarantined.append((j, jdst))
                    jw.log({"op": "intent", "kind": "rmdir", "seq": seq,
                            "dst": str(dst)})
                    _retrying(dst.rmdir)
                    entry = {"op": "rmdir", "src": None, "dst": str(dst),
                             "seq": seq}
                else:
                    raise ValueError(f"unknown op kind: {op.kind}")

            except (OSError, ValueError) as e:
                result.failed.append((op, _describe(e)))
                continue

            # Implicitly-created parents are journaled BEFORE the op entry so
            # reverse replay undoes the op first, then rmdirs the dirs.
            for d in created:
                jw.log({"op": "mkdir", "src": None, "dst": str(d), "seq": seq})
            jw.log(entry)
            result.done.append(op)

        jw.log({"op": "end_run", "id": run_id})
    return result


# --- Journal reading --------------------------------------------------------

def _read_journal(journal: Path) -> list[dict]:
    entries = []
    with open(journal, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                try:
                    entries.append(json.loads(line))
                except json.JSONDecodeError:
                    continue  # torn write from a crash — ignore
    return entries


def list_runs(journal: Path) -> list[dict]:
    """Every run in the journal, oldest first.

    Each record: ``{id, session, label, ts, v, ops, intents, undone, open}``.
    """
    if not journal.exists():
        return []
    runs: list[dict] = []
    current: dict | None = None
    for e in _read_journal(journal):
        kind = e.get("op")
        if kind == "begin_run":
            current = {"id": e.get("id"), "session": e.get("session") or e.get("id"),
                       "label": e.get("label"), "ts": e.get("ts"),
                       "v": e.get("v", 1), "ops": [], "intents": [],
                       "undone": False, "open": True}
            runs.append(current)
        elif kind == "end_run":
            if current is not None:
                current["open"] = False
            current = None
        elif kind == "undone_run":
            rid = e.get("id")
            for r in reversed(runs):
                if (r["id"] == rid) if rid else (not r["undone"]):
                    r["undone"] = True
                    break
        elif kind == "intent":
            if current is not None:
                current["intents"].append(e)
        elif current is not None:
            current["ops"].append(e)
    return runs


def pending_runs(journal: Path) -> list[dict]:
    """Runs that have not been undone, oldest first."""
    return [r for r in list_runs(journal) if not r["undone"]]


def last_run_ops(journal: Path) -> list[dict]:
    """Ops of the most recent run that hasn't been undone yet."""
    pending = pending_runs(journal)
    return pending[-1]["ops"] if pending else []


def _mark_undone(journal: Path, run_id) -> None:
    with open(journal, "a", encoding="utf-8") as jf:
        jf.write(json.dumps({"op": "undone_run", "id": run_id,
                             "ts": time.time()}) + "\n")
        jf.flush()
        os.fsync(jf.fileno())


# --- Recovery of half-finished mutations ------------------------------------

def recover(journal: Path, dry_run: bool = False) -> list[str]:
    """Clean up mutations that started but never completed.

    Returns human-readable descriptions of what was (or would be) done.
    """
    if not journal.exists():
        return []
    notes: list[str] = []
    for run in list_runs(journal):
        completed = {(e.get("seq"), e.get("op")) for e in run["ops"]}
        for intent in run["intents"]:
            seq, kind = intent.get("seq"), intent.get("kind")
            if (seq, kind) in completed:
                continue
            src = Path(intent["src"]) if intent.get("src") else None
            dst = Path(intent["dst"]) if intent.get("dst") else None
            tmp = Path(intent["tmp"]) if intent.get("tmp") else None

            if tmp is not None and os.path.lexists(tmp):
                target = src if src is not None and not os.path.lexists(src) else None
                if target is None:
                    notes.append(f"stranded temp file (source occupied): {tmp}")
                    continue
                notes.append(f"restore stranded temp file {tmp} -> {target}")
                if not dry_run:
                    os.rename(tmp, target)
                continue

            if (kind == "move" and src is not None and dst is not None
                    and os.path.lexists(src) and os.path.lexists(dst)
                    # ONLY when we positively recorded that the destination did
                    # not exist when we started — i.e. anything there now was
                    # created by us. A failed move onto a pre-existing file
                    # also leaves an unfinished intent, and deleting *that*
                    # destroys a file we never owned. Journals without the
                    # flag (v1, or v2 before this was added) fall through to
                    # "leave it alone", which is the safe default.
                    and intent.get("dst_existed") is False):
                # The move never completed and the source is intact, so what
                # sits at dst is a partial copy we created. Removing it also
                # unblocks retries: the collision check would otherwise see it
                # and skip this item forever.
                notes.append(f"remove incomplete copy: {dst}")
                if not dry_run:
                    try:
                        if dst.is_dir() and not dst.is_symlink():
                            shutil.rmtree(dst)
                        else:
                            os.unlink(dst)
                    except OSError as exc:
                        notes[-1] = f"could not remove incomplete copy {dst}: {exc}"
                continue

            if (kind == "move" and src is not None and dst is not None
                    and not os.path.lexists(src) and os.path.lexists(dst)):
                notes.append(
                    f"note: {src} -> {dst} completed but was not journaled "
                    f"(crash before the write); it cannot be undone automatically")
    return notes


# --- Undo -------------------------------------------------------------------

def _reverse_ops(entries: list[dict]) -> list[Op]:
    reverse: list[Op] = []
    for e in reversed(entries):
        if e["op"] == "move":
            reverse.append(Op("move", Path(e["dst"]), Path(e["src"])))
        elif e["op"] == "mkdir":
            reverse.append(Op("rmdir", None, Path(e["dst"])))
        elif e["op"] == "rmdir":
            reverse.append(Op("mkdir", None, Path(e["dst"])))
    return reverse


def _identity_by_path(entries: list[dict]) -> dict[str, dict]:
    out = {}
    for e in entries:
        if e["op"] == "move" and ("size" in e or "mtime" in e):
            out[os.path.normcase(e["dst"])] = e
    return out


def _identity_matches(path: Path, recorded: dict) -> tuple[bool, str]:
    st = _stat_fields(path)
    if not st:
        return True, ""  # can't tell; let the move itself decide
    if "size" in recorded and st["size"] != recorded["size"]:
        return False, (f"{path} is {st['size']} bytes but was {recorded['size']} "
                       f"when moved - it looks like it was replaced since")
    if "mtime" in recorded and abs(st["mtime"] - recorded["mtime"]) > 2:
        return False, (f"{path} was modified after the move "
                       f"- refusing to move it back")
    return True, ""


def _apply_reverse(reverse: list[Op], identities: dict[str, dict],
                   *, force: bool, lenient: bool) -> ExecResult:
    """Apply reversed ops.

    `lenient` allows "the source is already gone, so an earlier partial undo
    must have reverted this" — true only when undoing the NEWEST pending run.
    Out of order it is a lie: with runs `a->b` then `b->c`, reversing the older
    run looks for `b`, which is now `c`, and treating that as done would retire
    the run and strand `a` permanently.
    """
    result = ExecResult()
    roots = _implicit_roots(reverse)
    for op in reverse:
        try:
            if op.kind == "move":
                src, dst = _op_path(op.src), _op_path(op.dst)
                _validate_op_path(op.src, roots)
                _validate_op_path(op.dst, roots)
                recorded = identities.get(os.path.normcase(str(src)))
                if recorded and not force and os.path.lexists(src):
                    ok, why = _identity_matches(src, recorded)
                    if not ok:
                        result.failed.append((op, f"{why} (use --force to override)"))
                        continue
                _retrying(lambda s=src, d=dst: _do_move(s, d))
            elif op.kind == "mkdir":
                _validate_op_path(op.dst, roots)
                _op_path(op.dst).mkdir(parents=True, exist_ok=True)
            elif op.kind == "rmdir":
                _validate_op_path(op.dst, roots)
                _retrying(_op_path(op.dst).rmdir)
        except OSError as e:
            # A reverse move whose source is already gone was reverted by an
            # earlier partial undo — count it done so the run can complete.
            if (lenient and op.kind == "move"
                    and not os.path.lexists(_op_path(op.src))):
                result.done.append(op)
                continue
            result.failed.append((op, _describe(e)))
            continue
        except ValueError as e:
            result.failed.append((op, str(e)))
            continue
        result.done.append(op)
    return result


def _undo_record(run: dict, journal: Path, *, dry_run: bool,
                 force: bool) -> ExecResult:
    reverse = _reverse_ops(run["ops"])
    if dry_run:
        return ExecResult(done=reverse, dry_run=True)
    # Leniency is decided here, once, so every caller gets it right: it is
    # only safe when this run is the newest one still pending.
    pending = pending_runs(journal)
    lenient = bool(pending) and pending[-1]["id"] == run["id"]
    result = _apply_reverse(reverse, _identity_by_path(run["ops"]),
                            force=force, lenient=lenient)
    if result.ok:
        _mark_undone(journal, run["id"])
    return result


def undo_last_run(journal: Path, dry_run: bool = False, *,
                  force: bool = False) -> ExecResult:
    """Reverse-replay the last un-undone run. Repeatable for earlier runs."""
    if dry_run:
        # Mirror what the real path does: it retires runs with no entries to
        # reverse and carries on, so the preview must skip them too or it
        # disagrees with the operation it is previewing.
        for run in reversed(pending_runs(journal)):
            if run["ops"]:
                return _undo_record(run, journal, dry_run=True, force=force)
        return ExecResult(dry_run=True)

    while True:
        pending = pending_runs(journal)
        if not pending:
            return ExecResult()
        run = pending[-1]
        if run["ops"]:
            return _undo_record(run, journal, dry_run=False, force=force)
        # A run in which every op failed has no entries to reverse. Returning
        # here without marking it undone (the old behaviour) made every
        # earlier run unreachable forever, so retire it and carry on.
        _mark_undone(journal, run["id"])


def _op_paths(run: dict) -> set[str]:
    out = set()
    for e in run["ops"]:
        for key in ("src", "dst"):
            val = e.get(key)
            if val:
                out.add(os.path.normcase(val))
    return out


def undo_run(journal: Path, run_id: str, dry_run: bool = False, *,
             force: bool = False) -> tuple[ExecResult | None, str]:
    """Undo one specific run. Returns (result, error message)."""
    runs = list_runs(journal)
    matches = [r for r in runs if r["id"] and r["id"].startswith(run_id)]
    if not matches:
        return None, f"no run matching '{run_id}' in {journal}"
    if len(matches) > 1:
        return None, f"'{run_id}' is ambiguous ({len(matches)} runs match)"
    run = matches[0]
    if run["undone"]:
        return None, f"run {run['id']} has already been undone"

    if not force:
        index = runs.index(run)
        mine = _op_paths(run)
        for newer in runs[index + 1:]:
            if newer["undone"]:
                continue
            overlap = mine & _op_paths(newer)
            if overlap:
                return None, (
                    f"run {newer['id']} is newer and touched the same paths "
                    f"(e.g. {sorted(overlap)[0]}) - undo it first, "
                    f"or use --force")
    return _undo_record(run, journal, dry_run=dry_run, force=force), ""


def undo_session(journal: Path, session_id: str | None = None,
                 dry_run: bool = False, *, force: bool = False) -> ExecResult:
    """Undo every run of one logical action, newest first.

    `--action full` performs several execute() calls; they share a session id
    so one command reverses the whole thing.
    """
    pending = pending_runs(journal)
    if not pending:
        return ExecResult(dry_run=dry_run)
    target = session_id or pending[-1]["session"]
    combined = ExecResult(dry_run=dry_run)
    for run in reversed([r for r in pending if r["session"] == target]):
        result = _undo_record(run, journal, dry_run=dry_run, force=force)
        combined.done.extend(result.done)
        combined.failed.extend(result.failed)
        if result.failed:
            break  # stop at the first blockage so state stays predictable
    return combined


def undo_last(journal: Path, count: int = 1, dry_run: bool = False, *,
              force: bool = False) -> ExecResult:
    """Undo the newest `count` runs, newest first."""
    combined = ExecResult(dry_run=dry_run)
    if count < 1:
        return combined

    if dry_run:
        # A dry run never writes `undone_run`, so the journal never advances —
        # calling undo_last_run in a loop would preview the SAME run `count`
        # times. Walk the pending runs directly instead.
        with_ops = [r for r in pending_runs(journal) if r["ops"]]
        for run in reversed(with_ops[-count:]):
            combined.done.extend(_reverse_ops(run["ops"]))
        return combined

    for _ in range(count):
        result = undo_last_run(journal, force=force)
        combined.done.extend(result.done)
        combined.failed.extend(result.failed)
        if result.failed or not result.done:
            break
    return combined
