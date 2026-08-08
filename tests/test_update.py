"""Tests for the self-update check and the update itself.

Every test builds a real repository pair on disk — a bare "origin" and a
clone standing in for the user's install — because the whole feature is an
opinion about what git says, and a mocked git would only test the opinion.
"""

import copy
import json
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path

import pytest

from mediaorg import update


def _git(cwd, *args):
    subprocess.run(
        ["git", "-c", "user.name=Test", "-c", "user.email=test@example.com",
         *args],
        cwd=str(cwd), check=True, capture_output=True, text=True,
    )


@pytest.fixture(autouse=True)
def isolated_background_state():
    """The background check keeps module-level state; don't leak it."""
    update.reset_background_state()
    yield
    update.reset_background_state()


def _quiesce():
    """Wait for the check thread to exit before the scratch repo goes away.

    The thread is a daemon, so nothing else stops it: without this an
    in-flight check outlives the fixture that pointed ``app_dir`` at a
    temporary clone and runs its next git command — a real fetch — against
    the developer's own repository.
    """
    assert update.join_background(30), "background update check did not finish"


@pytest.fixture
def repos(tmp_path, monkeypatch):
    """A bare origin, a clone of it, and the clone wired up as the app dir.

    Returns (origin, clone). Committing to `origin` is done through a second
    working copy so the bare repo stays bare.
    """
    if shutil.which("git") is None:
        pytest.skip("git is not installed")

    origin = tmp_path / "origin.git"
    seed = tmp_path / "seed"
    clone = tmp_path / "clone"

    _git(tmp_path, "init", "--bare", "--initial-branch=main", str(origin))
    _git(tmp_path, "clone", str(origin), str(seed))
    (seed / "run.py").write_text("print('v1')\n")
    (seed / "requirements.txt").write_text("pandas>=2.2\n")
    _git(seed, "add", "-A")
    _git(seed, "commit", "-m", "initial")
    _git(seed, "push", "-u", "origin", "main")
    _git(tmp_path, "clone", str(origin), str(clone))

    monkeypatch.setattr(update, "app_dir", lambda: clone)
    monkeypatch.delenv("MEDIAORG_NO_UPDATE_CHECK", raising=False)
    monkeypatch.delenv("MEDIAORG_UPDATE_INTERVAL", raising=False)
    yield seed, clone
    # Teardown runs before monkeypatch unwinds app_dir, which is the whole
    # point: the thread must be gone while it still points at the clone.
    _quiesce()


def _push_commit(seed, message="new work", path="run.py", body=None):
    (seed / path).write_text(body if body is not None else f"# {message}\n")
    _git(seed, "add", "-A")
    _git(seed, "commit", "-m", message)
    _git(seed, "push", "origin", "main")


# --- checking ----------------------------------------------------------------

def test_fresh_clone_is_current(repos):
    st = update.check(fetch=True)
    assert st.state == update.CURRENT
    assert st.behind == 0
    assert st.upstream == "origin/main"
    assert not st.update_available


def test_counts_commits_behind_and_lists_them(repos):
    seed, _ = repos
    _push_commit(seed, "fix the renamer")
    _push_commit(seed, "add a menu entry")

    st = update.check(fetch=True)
    assert st.state == update.BEHIND
    assert st.behind == 2
    assert st.ahead == 0
    assert st.update_available
    # Newest first, and readable enough to decide whether to update.
    assert [c.split(" ", 1)[1] for c in st.commits] == [
        "add a menu entry", "fix the renamer"]


def test_local_commits_read_as_ahead_not_behind(repos):
    _, clone = repos
    (clone / "notes.txt").write_text("mine")
    _git(clone, "add", "-A")
    _git(clone, "commit", "-m", "local tweak")

    st = update.check(fetch=True)
    assert st.state == update.AHEAD
    assert st.ahead == 1 and st.behind == 0
    assert not st.update_available


def test_both_sides_moved_is_diverged(repos):
    seed, clone = repos
    _push_commit(seed, "upstream work")
    (clone / "notes.txt").write_text("mine")
    _git(clone, "add", "-A")
    _git(clone, "commit", "-m", "local tweak")

    st = update.check(fetch=True)
    assert st.state == update.DIVERGED
    assert st.ahead == 1 and st.behind == 1
    # Diverged is not "an update is available" — a pull cannot fast-forward.
    assert not st.update_available
    assert "git reset --hard origin/main" in update.describe(st)


def test_check_without_fetch_uses_what_is_already_known(repos):
    seed, _ = repos
    _push_commit(seed, "unfetched work")
    # No fetch, so the new commit is invisible until someone asks the network.
    assert update.check(fetch=False).state == update.CURRENT
    assert update.check(fetch=True).state == update.BEHIND


def test_non_git_install_explains_how_to_get_one(tmp_path, monkeypatch):
    plain = tmp_path / "downloaded-zip"
    plain.mkdir()
    monkeypatch.setattr(update, "app_dir", lambda: plain)
    st = update.check(fetch=True)
    assert st.state == update.UNKNOWN
    assert "git clone" in st.hint
    assert "[!] Could not check for updates" in update.describe(st)


def test_re_clone_advice_names_the_files_that_do_not_come_with_it(
        tmp_path, monkeypatch):
    """This hint used to end "so nothing is lost", which was false.

    Every one of these resolves to the app directory — the folder being
    replaced — so a re-clone into a fresh folder silently leaves the undo
    history and the word list behind unless the user is told to bring them.
    """
    plain = tmp_path / "downloaded-zip"
    plain.mkdir()
    monkeypatch.setattr(update, "app_dir", lambda: plain)
    hint = update.check(fetch=True).hint

    for untracked in ("mediaorg_journal", "custom_strip_patterns.json",
                      ".media_llm_config.json", ".media_renamer_config.json"):
        assert untracked in hint
    assert "nothing is lost" not in hint


def test_detached_head_says_which_branch_to_return_to(repos):
    _, clone = repos
    _git(clone, "checkout", "--detach", "HEAD")
    st = update.check(fetch=True)
    assert st.state == update.UNKNOWN
    assert "detached HEAD" in st.reason
    assert "git checkout main" in st.hint


# --- the banner --------------------------------------------------------------

def test_banner_is_silent_when_up_to_date(repos):
    assert update.banner(update.check(fetch=True)) == ""


def test_banner_names_the_command_and_the_folder(repos):
    seed, clone = repos
    _push_commit(seed, "something new")
    text = update.banner(update.check(fetch=True))
    assert "1 commit behind origin/main" in text
    assert "python run.py --update" in text
    assert str(clone) in text          # ...and *where* to type it


def test_banner_names_the_off_switch(repos):
    """The banner is the only place a user on a restricted network finds out
    that launching the app talks to github.com, and how to stop it."""
    seed, _ = repos
    _push_commit(seed, "something new")
    text = update.banner(update.check(fetch=True))
    assert "github.com" in text
    assert "MEDIAORG_NO_UPDATE_CHECK=1" in text


def test_banner_pluralises(repos):
    seed, _ = repos
    _push_commit(seed, "one")
    _push_commit(seed, "two")
    assert "2 commits behind" in update.banner(update.check(fetch=True))


# --- cache -------------------------------------------------------------------

def test_cache_round_trip(repos):
    seed, clone = repos
    _push_commit(seed, "cached work")
    st = update.check(fetch=True)
    update.save_cache(st)

    loaded = update.load_cache()
    assert loaded is not None
    assert loaded.state == update.BEHIND
    assert loaded.behind == 1
    assert loaded.stale is True        # marked as "remembered", not "measured"


def test_corrupt_cache_is_ignored_not_fatal(repos):
    _, clone = repos
    (clone / update.CACHE_NAME).write_text("{not json")
    assert update.load_cache() is None


def test_cache_from_a_newer_version_is_tolerated(repos):
    _, clone = repos
    (clone / update.CACHE_NAME).write_text(json.dumps(
        {"state": "behind", "behind": 3, "some_future_field": True}))
    loaded = update.load_cache()
    assert loaded is not None and loaded.behind == 3


def test_fresh_cache_skips_the_network(repos, monkeypatch):
    seed, _ = repos
    update.save_cache(update.check(fetch=True))          # CURRENT
    _push_commit(seed, "invisible until the cache expires")

    called = []
    monkeypatch.setattr(update, "check",
                        lambda **kw: called.append(kw) or update.UpdateStatus())
    update.begin_background_check()
    _quiesce()                       # deterministic: no wall-clock guessing
    assert called == []
    assert update.latest_status().state == update.CURRENT


def test_stale_cache_triggers_a_refresh(repos):
    seed, _ = repos
    stale = update.check(fetch=True)
    stale.checked_at = time.time() - 999 * 3600
    update.save_cache(stale)
    _push_commit(seed, "found on refresh")

    update.begin_background_check()
    assert update.wait_for_check(30).state == update.BEHIND


def test_offline_check_does_not_overwrite_a_good_cached_answer(repos, monkeypatch):
    seed, _ = repos
    _push_commit(seed, "known update")
    good = update.check(fetch=True)
    good.checked_at = time.time() - 999 * 3600
    update.save_cache(good)

    monkeypatch.setattr(update, "check", lambda **kw: update.UpdateStatus(
        state=update.UNKNOWN, reason="could not reach the remote"))
    update.begin_background_check()
    _quiesce()
    # The remembered "you are behind" survives a flight-mode launch.
    assert update.load_cache().state == update.BEHIND
    assert update.latest_status().state == update.BEHIND


def test_check_can_be_disabled(repos, monkeypatch):
    monkeypatch.setenv("MEDIAORG_NO_UPDATE_CHECK", "1")
    update.begin_background_check()
    assert update.latest_status() is None
    assert update.banner() == ""


# --- updating ----------------------------------------------------------------

def test_update_fast_forwards(repos, capsys):
    seed, clone = repos
    _push_commit(seed, "the fix", body="print('v2')\n")

    assert update.run_update(assume_yes=True) == 0
    assert (clone / "run.py").read_text() == "print('v2')\n"
    assert update.check(fetch=False).state == update.CURRENT
    assert "Restart Media Organizer" in capsys.readouterr().out


def test_update_when_current_changes_nothing(repos, capsys):
    _, clone = repos
    before = (clone / "run.py").read_text()
    assert update.run_update(assume_yes=True) == 0
    assert (clone / "run.py").read_text() == before
    assert "Nothing to do." in capsys.readouterr().out


def test_dry_run_reports_without_pulling(repos, capsys):
    seed, clone = repos
    _push_commit(seed, "the fix", body="print('v2')\n")

    assert update.run_update(dry_run=True) == 0
    assert (clone / "run.py").read_text() == "print('v1')\n"
    assert "[dry-run]" in capsys.readouterr().out


def test_local_edits_block_the_update_and_are_not_lost(repos, capsys):
    seed, clone = repos
    _push_commit(seed, "the fix", body="print('v2')\n")
    (clone / "run.py").write_text("print('my own edit')\n")

    assert update.run_update(assume_yes=True) == 1
    assert (clone / "run.py").read_text() == "print('my own edit')\n"
    out = capsys.readouterr().out
    assert "git stash" in out                 # told how to keep them
    assert "run.py" in out                    # told which files


def test_untracked_files_do_not_block_the_update(repos):
    seed, clone = repos
    _push_commit(seed, "the fix", body="print('v2')\n")
    # The journal, spreadsheets and configs all live next to the app.
    (clone / "mediaorg_journal.jsonl").write_text('{"op":"begin_run"}\n')
    (clone / "media_library.xlsx").write_text("not really xlsx")

    assert update.run_update(assume_yes=True) == 0
    assert (clone / "run.py").read_text() == "print('v2')\n"
    assert (clone / "mediaorg_journal.jsonl").exists()


def test_diverged_clone_refuses_rather_than_clobbering(repos, capsys):
    seed, clone = repos
    _push_commit(seed, "upstream work")
    (clone / "notes.txt").write_text("mine")
    _git(clone, "add", "-A")
    _git(clone, "commit", "-m", "local tweak")

    assert update.run_update(assume_yes=True) == 1
    assert (clone / "notes.txt").read_text() == "mine"
    assert "Diverged" in capsys.readouterr().out


def test_dependencies_reinstalled_only_when_requirements_move(repos, monkeypatch):
    seed, _ = repos
    calls = []
    monkeypatch.setattr(update, "_pip_install", lambda req: calls.append(req) or True)

    _push_commit(seed, "no dependency change")
    assert update.run_update(assume_yes=True) == 0
    assert calls == []

    _push_commit(seed, "needs a newer pandas", path="requirements.txt",
                 body="pandas>=2.3\n")
    assert update.run_update(assume_yes=True) == 0
    assert len(calls) == 1


def test_update_refreshes_the_cache(repos):
    seed, _ = repos
    _push_commit(seed, "the fix")
    update.save_cache(update.check(fetch=True))
    assert update.load_cache().state == update.BEHIND

    assert update.run_update(assume_yes=True) == 0
    # A stale "update available" banner after updating would be maddening.
    assert update.load_cache().state == update.CURRENT


def test_update_of_a_non_git_copy_explains_the_alternative(tmp_path, monkeypatch, capsys):
    plain = tmp_path / "zip-install"
    plain.mkdir()
    monkeypatch.setattr(update, "app_dir", lambda: plain)
    assert update.run_update(assume_yes=True) == 1
    assert "git clone" in capsys.readouterr().out


def test_first_launch_can_wait_for_the_answer(repos):
    """With no cache there is nothing to show yet, so the caller is allowed
    to wait a moment rather than let the first launch say nothing."""
    seed, _ = repos
    _push_commit(seed, "the very first thing you should hear about")
    update.begin_background_check()
    st = update.wait_for_check(30)
    assert st is not None and st.state == update.BEHIND


def test_waiting_when_no_check_is_running_returns_immediately(repos):
    update.save_cache(update.check(fetch=True))
    update.begin_background_check()               # cache is fresh: no thread
    started = time.monotonic()
    assert update.wait_for_check(30).state == update.CURRENT
    assert time.monotonic() - started < 1


# --- what counts as "this install" -------------------------------------------

def test_a_copy_inside_another_repo_is_not_treated_as_a_clone(tmp_path, monkeypatch):
    """A ZIP unpacked into some other project's checkout is inside *that*
    repository. Updating there would fetch and fast-forward the unrelated
    project and leave Media Organizer untouched."""
    outer = tmp_path / "someone-elses-project"
    inner = outer / "tools" / "MediaOrganizer"
    inner.mkdir(parents=True)
    _git(tmp_path, "init", "--initial-branch=main", str(outer))
    (outer / "their_code.py").write_text("theirs")
    _git(outer, "add", "-A")
    _git(outer, "commit", "-m", "their work")

    monkeypatch.setattr(update, "app_dir", lambda: inner)
    assert update.is_git_checkout() is False
    st = update.check(fetch=True)
    assert st.state == update.UNKNOWN
    assert "sits inside another repository" in st.reason
    assert "git clone" in st.hint


def test_updating_a_copy_inside_another_repo_touches_nothing(tmp_path, monkeypatch):
    outer = tmp_path / "someone-elses-project"
    inner = outer / "tools" / "MediaOrganizer"
    inner.mkdir(parents=True)
    _git(tmp_path, "init", "--initial-branch=main", str(outer))
    (outer / "their_code.py").write_text("theirs")
    _git(outer, "add", "-A")
    _git(outer, "commit", "-m", "their work")
    before = subprocess.run(["git", "rev-parse", "HEAD"], cwd=str(outer),
                            capture_output=True, text=True).stdout

    monkeypatch.setattr(update, "app_dir", lambda: inner)
    assert update.run_update(assume_yes=True) == 1
    after = subprocess.run(["git", "rev-parse", "HEAD"], cwd=str(outer),
                           capture_output=True, text=True).stdout
    assert before == after
    assert (outer / "their_code.py").read_text() == "theirs"


def test_the_clone_root_itself_is_a_valid_install(repos):
    assert update.is_git_checkout() is True


# --- which branch we compare against -----------------------------------------

def test_an_untracked_work_branch_is_not_assumed_to_follow_main(repos):
    """Treating any untracked branch as tracking origin/main would let
    --update fast-forward somebody's work branch onto main."""
    seed, clone = repos
    _push_commit(seed, "main moved on")
    _git(clone, "checkout", "-b", "my-tweaks")

    st = update.check(fetch=True)
    assert st.state == update.UNKNOWN
    assert "does not track a remote branch" in st.reason
    assert "git branch --set-upstream-to=origin/my-tweaks" in st.hint


def test_an_untracked_work_branch_is_never_pulled(repos):
    seed, clone = repos
    _push_commit(seed, "main moved on")
    _git(clone, "checkout", "-b", "my-tweaks")
    head = update.head_revision()

    assert update.run_update(assume_yes=True) == 1
    assert update.head_revision() == head          # their branch is untouched


def test_main_without_tracking_config_still_falls_back(repos):
    """A clone whose upstream config was lost is still updatable on main."""
    seed, clone = repos
    _push_commit(seed, "a fix")
    _git(clone, "branch", "--unset-upstream")

    st = update.check(fetch=True)
    assert st.state == update.BEHIND
    assert st.upstream == "origin/main"


# --- the cache tracks the checkout, not just the clock ------------------------

def test_a_cache_from_a_different_commit_is_not_shown(repos):
    """After a manual `git pull` the cached 'you are behind' is not merely
    old, it is about a commit that is no longer running."""
    seed, _ = repos
    _push_commit(seed, "work the user pulled by hand")
    behind = update.check(fetch=True)
    update.save_cache(behind)                      # fresh timestamp, stale sha
    _git(repos[1], "pull", "--ff-only", "origin", "main")   # the manual pull

    update.begin_background_check()
    # Never published, not even briefly, on the way to the refreshed answer.
    assert (update.latest_status() or update.UpdateStatus()).state != update.BEHIND
    assert update.wait_for_check(30).state == update.CURRENT


def test_a_cache_from_a_different_branch_is_not_shown(repos):
    _, clone = repos
    update.save_cache(update.check(fetch=True))
    _git(clone, "checkout", "-b", "elsewhere")
    assert update.describes_current_checkout(update.load_cache()) is False


def test_a_cache_matching_the_checkout_is_still_used(repos):
    update.save_cache(update.check(fetch=True))
    assert update.describes_current_checkout(update.load_cache()) is True


# --- confirmation -------------------------------------------------------------

def test_no_terminal_and_no_yes_refuses_rather_than_assuming(repos, monkeypatch,
                                                             capsys):
    """`python run.py --update < /dev/null` must not change the checkout
    without the confirmation --yes is documented to stand in for."""
    seed, _ = repos
    _push_commit(seed, "the fix", body="print('v2')\n")
    monkeypatch.setattr(update, "_can_prompt", lambda: False)
    head = update.head_revision()

    assert update.run_update() == 1
    assert update.head_revision() == head
    assert "--update --yes" in capsys.readouterr().out


def test_no_terminal_with_yes_proceeds(repos, monkeypatch):
    seed, clone = repos
    _push_commit(seed, "the fix", body="print('v2')\n")
    monkeypatch.setattr(update, "_can_prompt", lambda: False)

    assert update.run_update(assume_yes=True) == 0
    assert (clone / "run.py").read_text() == "print('v2')\n"


def test_declining_the_confirmation_changes_nothing(repos, monkeypatch):
    seed, clone = repos
    _push_commit(seed, "the fix", body="print('v2')\n")
    monkeypatch.setattr(update, "_can_prompt", lambda: True)
    monkeypatch.setattr("builtins.input", lambda *a: "n")

    assert update.run_update() == 0
    assert (clone / "run.py").read_text() == "print('v1')\n"


# --- what a failed git call is allowed to mean --------------------------------

def test_a_failed_comparison_is_not_reported_as_up_to_date(repos, monkeypatch):
    """(0, 0) means "no difference"; a failed rev-list means "I don't know".
    Conflating them hides the very update that would repair a broken clone."""
    monkeypatch.setattr(update, "local_commits_behind", lambda upstream: None)
    st = update.check(fetch=True)
    assert st.state == update.UNKNOWN
    assert "could not compare" in st.reason
    assert update.banner(st) == ""


def test_rev_list_failure_surfaces_as_none(repos):
    assert update.local_commits_behind("origin/no-such-branch-here") is None


def test_git_env_wins_over_an_inherited_askpass(monkeypatch):
    """An inherited GIT_ASKPASS is the case that hangs — a GUI prompt on an
    unattended box — so ours must overwrite it, not defer to it."""
    monkeypatch.setenv("GIT_ASKPASS", "/usr/bin/some-gui-prompter")
    monkeypatch.setenv("SSH_ASKPASS", "/usr/bin/some-gui-prompter")
    env = update._git_env()
    assert env["GIT_ASKPASS"] == ""
    assert env["SSH_ASKPASS"] == ""
    assert env["GIT_TERMINAL_PROMPT"] == "0"


def test_git_env_drops_repo_redirecting_variables(monkeypatch):
    """GIT_DIR/GIT_WORK_TREE override cwd, which would send every command —
    including the pull — at a repository that is not this clone."""
    monkeypatch.setenv("GIT_DIR", "/somewhere/else/.git")
    monkeypatch.setenv("GIT_WORK_TREE", "/somewhere/else")
    monkeypatch.setenv("GIT_INDEX_FILE", "/somewhere/else/index")
    env = update._git_env()
    assert "GIT_DIR" not in env
    assert "GIT_WORK_TREE" not in env
    assert "GIT_INDEX_FILE" not in env


def test_an_inherited_git_dir_does_not_redirect_the_check(repos, tmp_path,
                                                          monkeypatch):
    elsewhere = tmp_path / "elsewhere"
    _git(tmp_path, "init", "--initial-branch=main", str(elsewhere))
    monkeypatch.setenv("GIT_DIR", str(elsewhere / ".git"))
    monkeypatch.setenv("GIT_WORK_TREE", str(elsewhere))
    # Still measuring the clone, not the repository the environment names.
    assert update.check(fetch=True).state == update.CURRENT


# --- a cache is a file anything can scribble in -------------------------------

def test_an_ill_typed_cache_is_rejected_not_believed(repos):
    """`{"behind": "5"}` used to crash the wizard on launch, before the menu
    was drawn, on every single run until the file was deleted by hand."""
    _, clone = repos
    (clone / update.CACHE_NAME).write_text(json.dumps(
        {"state": "behind", "behind": "5", "upstream": "origin/main"}))
    assert update.load_cache() is None
    # The value that used to reach banner() cannot be built at all now.
    with pytest.raises(ValueError):
        update.UpdateStatus.from_dict({"state": "behind", "behind": "5"})


def test_a_rejected_cache_is_removed_so_it_cannot_recur(repos):
    _, clone = repos
    (clone / update.CACHE_NAME).write_text(json.dumps({"checked_at": None}))
    update.load_cache()
    assert not (clone / update.CACHE_NAME).exists()


def test_a_true_is_not_a_commit_count(repos):
    _, clone = repos
    (clone / update.CACHE_NAME).write_text(json.dumps({"behind": True}))
    assert update.load_cache() is None


def test_an_integer_timestamp_is_accepted_as_a_float(repos):
    _, clone = repos
    (clone / update.CACHE_NAME).write_text(json.dumps(
        {"state": "current", "checked_at": 1700000000}))
    loaded = update.load_cache()
    assert loaded is not None and loaded.checked_at == 1700000000.0


def test_an_unknown_is_never_written_to_the_cache(repos):
    _, clone = repos
    update.save_cache(update.UpdateStatus(state=update.UNKNOWN, reason="offline"))
    assert not (clone / update.CACHE_NAME).exists()


# --- the launch thread does no work -------------------------------------------

def test_begin_background_check_does_not_run_git_on_the_caller(repos, monkeypatch):
    """Everything — cache read, checkout comparison, network — belongs to the
    thread. The launch path promises to return immediately, so it must."""
    callers = []
    real_git = update._git

    def recording_git(*args, **kw):
        callers.append(threading.current_thread().name)
        return real_git(*args, **kw)

    monkeypatch.setattr(update, "_git", recording_git)
    update.begin_background_check()
    caller_ran_git = [n for n in callers if n == threading.current_thread().name]
    assert caller_ran_git == []
    _quiesce()
    assert callers, "the thread should have done the work instead"


def test_waiting_for_the_cache_returns_at_once_when_nothing_is_running(repos):
    started = time.monotonic()
    assert update.wait_for_cache(30) is None
    assert time.monotonic() - started < 1


# --- pip ----------------------------------------------------------------------

def test_pip_is_bounded_and_cannot_ask_questions(repos, monkeypatch):
    seen = {}

    def fake_run(cmd, **kwargs):
        seen.update(kwargs)
        seen['cmd'] = cmd
        return subprocess.CompletedProcess(cmd, 0)

    monkeypatch.setattr(update.subprocess, "run", fake_run)
    assert update._pip_install(Path("requirements.txt")) is True
    assert seen['stdin'] is subprocess.DEVNULL     # pip cannot stop for input
    # A deadline for the whole call, so two attempts cannot cost 2x the bound.
    assert 0 < seen['timeout'] <= update.PIP_TIMEOUT


def test_pip_timeout_is_reported_not_retried(repos, monkeypatch, capsys):
    calls = []

    def fake_run(cmd, **kwargs):
        calls.append(cmd)
        raise subprocess.TimeoutExpired(cmd, kwargs.get('timeout', 0))

    monkeypatch.setattr(update.subprocess, "run", fake_run)
    monkeypatch.setattr(update, "_in_virtualenv", lambda: False)
    assert update._pip_install(Path("requirements.txt")) is False
    assert len(calls) == 1                          # no --user retry after a hang
    assert "ran out of its" in capsys.readouterr().out


def test_no_user_retry_inside_a_virtualenv(repos, monkeypatch):
    calls = []

    def fake_run(cmd, **kwargs):
        calls.append(cmd)
        raise subprocess.CalledProcessError(1, cmd)

    monkeypatch.setattr(update.subprocess, "run", fake_run)
    monkeypatch.setattr(update, "_in_virtualenv", lambda: True)
    assert update._pip_install(Path("requirements.txt")) is False
    assert len(calls) == 1
    assert not any("--user" in c for c in calls)    # pip refuses it in a venv


def test_dependencies_reinstall_is_decided_by_the_file_not_a_diff(repos, monkeypatch):
    """Reading requirements.txt either side of the pull works even when the
    commit range cannot be computed, and gets a deletion right."""
    seed, _ = repos
    calls = []
    monkeypatch.setattr(update, "_pip_install", lambda req: calls.append(req) or True)
    monkeypatch.setattr(update, "_git", _failing_diff(update._git))

    _push_commit(seed, "needs a newer pandas", path="requirements.txt",
                 body="pandas>=2.3\n")
    assert update.run_update(assume_yes=True) == 0
    assert len(calls) == 1


def _failing_diff(real):
    def wrapper(*args, **kw):
        if args and args[0] == "diff":
            return 1, "", "simulated: git diff unavailable"
        return real(*args, **kw)
    return wrapper


def test_a_failed_dependency_install_is_not_reported_as_success(repos, monkeypatch,
                                                                capsys):
    seed, _ = repos
    monkeypatch.setattr(update, "_pip_install", lambda req: False)
    _push_commit(seed, "needs a newer pandas", path="requirements.txt",
                 body="pandas>=2.3\n")

    code = update.run_update(assume_yes=True)
    out = capsys.readouterr().out
    assert code == 1                       # the caller can tell something broke
    assert "[OK] Updated" not in out       # ...and so can the user
    assert "dependency install failed" in out


# --- the confirmation ---------------------------------------------------------

def test_ctrl_d_at_the_confirmation_is_a_no_not_a_traceback(repos, monkeypatch,
                                                            capsys):
    seed, clone = repos
    _push_commit(seed, "the fix", body="print('v2')\n")
    monkeypatch.setattr(update, "_can_prompt", lambda: True)

    def eof(*a):
        raise EOFError

    monkeypatch.setattr("builtins.input", eof)
    assert update.run_update() == 0
    assert (clone / "run.py").read_text() == "print('v1')\n"
    assert "Cancelled" in capsys.readouterr().out


def test_ctrl_c_at_the_confirmation_is_a_no_not_a_traceback(repos, monkeypatch):
    seed, clone = repos
    _push_commit(seed, "the fix", body="print('v2')\n")
    monkeypatch.setattr(update, "_can_prompt", lambda: True)

    def interrupt(*a):
        raise KeyboardInterrupt

    monkeypatch.setattr("builtins.input", interrupt)
    assert update.run_update() == 0
    assert (clone / "run.py").read_text() == "print('v1')\n"


def test_the_dirty_tree_advice_actually_clears_a_staged_change(repos, capsys):
    """`git checkout -- .` restores from the index, so it leaves staged
    changes in place — and `git status --porcelain` still calls them dirty."""
    seed, clone = repos
    _push_commit(seed, "the fix", body="print('v2')\n")
    (clone / "run.py").write_text("print('mine')\n")
    _git(clone, "add", "run.py")                    # staged, not just modified

    assert update.run_update(assume_yes=True) == 1
    out = capsys.readouterr().out
    assert "git checkout -- ." not in out           # would not have worked
    assert "git reset --hard" in out

    _git(clone, "reset", "--hard")                  # the advice, followed
    assert update.working_tree_changes() == []
    assert update.run_update(assume_yes=True) == 0  # and now it goes through


# --- upstreams that name no remote --------------------------------------------

def test_an_upstream_on_another_local_branch_is_refused(repos):
    """`git pull --ff-only main ""` is what splitting that on '/' produces."""
    _, clone = repos
    _git(clone, "branch", "release")
    _git(clone, "checkout", "-b", "work")
    _git(clone, "branch", "--set-upstream-to=release")

    st = update.check(fetch=True)
    assert st.state == update.UNKNOWN
    assert "not a remote branch" in st.reason
    assert update.run_update(assume_yes=True) == 1


def test_split_upstream_rejects_what_it_cannot_pull_from():
    assert update.split_upstream("origin/main") == ("origin", "main")
    assert update.split_upstream("upstream/feature/x") == ("upstream", "feature/x")
    assert update.split_upstream("main") is None
    assert update.split_upstream("origin/") is None
    assert update.split_upstream("") is None


# --- text from the remote -----------------------------------------------------

def test_commit_subjects_cannot_repaint_the_terminal(repos):
    """Subjects are remote-controlled, printed by --check-update and the
    doctor, and replayed from the cache long afterwards."""
    seed, _ = repos
    _push_commit(seed, "innocent\x1b[2J\x1b[H FORGED: up to date")

    st = update.check(fetch=True)
    assert st.commits
    assert "\x1b" not in st.commits[0]
    assert "\x1b" not in update.describe(st)


def test_a_hostile_subject_cannot_pad_the_report_indefinitely(repos):
    seed, _ = repos
    _push_commit(seed, "x" * 5000)
    st = update.check(fetch=True)
    assert len(st.commits[0]) <= 210


# --- the CLI ------------------------------------------------------------------

def test_dispatch_declines_argv_that_is_not_its_business():
    assert update.dispatch_cli(["--action", "scan"]) is None
    assert update.dispatch_cli([]) is None


def test_dispatch_rejects_junk_rather_than_ignoring_it(capsys):
    assert update.dispatch_cli(["--version", "--nonsense"]) == 2
    assert "Unrecognized argument" in capsys.readouterr().out


def test_version_reports_the_installed_commit(repos, capsys):
    assert update.dispatch_cli(["--version"]) == 0
    out = capsys.readouterr().out
    sha = update.head_revision()
    assert sha and sha in out          # '' in out would pass no matter what


# --- the kill switch is a promise ---------------------------------------------

def test_the_doctor_respects_the_no_check_switch(repos, monkeypatch):
    """MEDIAORG_NO_UPDATE_CHECK=1 is documented as "nothing is sent". A
    diagnostic that fetches anyway breaks the guarantee it reports on."""
    from mediaorg import doctor
    seed, _ = repos
    _push_commit(seed, "not to be discovered over the network")
    monkeypatch.setenv("MEDIAORG_NO_UPDATE_CHECK", "1")

    fetches = []
    real_git = update._git

    def recording_git(*args, **kw):
        if args and args[0] == "fetch":
            fetches.append(args)
        return real_git(*args, **kw)

    monkeypatch.setattr(update, "_git", recording_git)
    status, message = doctor.check_version()
    assert fetches == []
    assert "MEDIAORG_NO_UPDATE_CHECK=1" in message


def test_locally_visible_problems_survive_the_switch(repos, monkeypatch):
    """Turning the network check off changes what is measured, not which
    problems get reported: a diverged clone is diverged either way."""
    from mediaorg import doctor
    seed, clone = repos
    _push_commit(seed, "upstream work")
    update.check(fetch=True)                     # get the remote ref locally
    (clone / "notes.txt").write_text("mine")
    _git(clone, "add", "-A")
    _git(clone, "commit", "-m", "local tweak")
    monkeypatch.setenv("MEDIAORG_NO_UPDATE_CHECK", "1")

    status, message = doctor.check_version()
    assert status == "WARN"
    assert "Diverged" in message
    assert "MEDIAORG_NO_UPDATE_CHECK=1" in message


def test_a_broken_install_is_reported_even_with_checks_off(tmp_path, monkeypatch):
    from mediaorg import doctor
    plain = tmp_path / "downloaded-zip"
    plain.mkdir()
    monkeypatch.setattr(update, "app_dir", lambda: plain)
    monkeypatch.setenv("MEDIAORG_NO_UPDATE_CHECK", "1")

    status, message = doctor.check_version()
    assert status == "WARN"                      # not a silent [OK]
    assert "git clone" in message


def test_offline_up_to_date_does_not_claim_a_fresh_comparison(repos, monkeypatch):
    from mediaorg import doctor
    monkeypatch.setenv("MEDIAORG_NO_UPDATE_CHECK", "1")
    status, message = doctor.check_version()
    assert status == "OK"
    assert "last fetched" in message


def test_the_doctor_still_checks_when_the_switch_is_off(repos, monkeypatch):
    from mediaorg import doctor
    seed, _ = repos
    _push_commit(seed, "a real update")
    monkeypatch.delenv("MEDIAORG_NO_UPDATE_CHECK", raising=False)
    status, message = doctor.check_version()
    assert status == "WARN"
    assert "1 commit behind" in message


# --- the cache validates values, not just types -------------------------------

def test_validation_survives_dropping_the_future_import(repos, monkeypatch):
    """`field.type` is a string only under PEP 563. If that import ever goes,
    every check must still run rather than silently accepting everything."""
    real = update.UpdateStatus.__dataclass_fields__
    patched = {}
    for name, field_def in real.items():
        clone_field = copy.copy(field_def)
        clone_field.type = {'str': str, 'int': int, 'float': float,
                            'list[str]': list, 'bool': bool}.get(field_def.type,
                                                                 field_def.type)
        patched[name] = clone_field
    monkeypatch.setattr(update.UpdateStatus, "__dataclass_fields__", patched)
    with pytest.raises(ValueError):
        update.UpdateStatus.from_dict({"behind": "5"})
    assert update.UpdateStatus.from_dict({"behind": 5}).behind == 5


def test_infinity_cannot_retire_the_update_check(repos):
    """time.time() - inf is -inf, which is younger than any interval — the
    check would look fresh forever and never run again."""
    _, clone = repos
    (clone / update.CACHE_NAME).write_text('{"state": "current", "checked_at": Infinity}')
    assert update.load_cache() is None


def test_nan_is_rejected_too(repos):
    _, clone = repos
    (clone / update.CACHE_NAME).write_text('{"state": "current", "checked_at": NaN}')
    assert update.load_cache() is None


def test_an_overflowing_timestamp_does_not_escape_as_arithmetic_error(repos):
    _, clone = repos
    (clone / update.CACHE_NAME).write_text(
        '{"state": "current", "checked_at": ' + "9" * 400 + '}')
    assert update.load_cache() is None        # not an OverflowError


def test_a_future_timestamp_does_not_count_as_fresh(repos):
    """A skewed clock, or a cache copied from another machine, must not
    suppress the check for the rest of time."""
    seed, _ = repos
    future = update.check(fetch=True)
    future.checked_at = time.time() + 5000 * 3600
    update.save_cache(future)
    _push_commit(seed, "would never be found")

    update.begin_background_check()
    assert update.wait_for_check(30).state == update.BEHIND


def test_a_state_that_is_not_a_state_is_rejected(repos):
    _, clone = repos
    (clone / update.CACHE_NAME).write_text(json.dumps({"state": "banana"}))
    assert update.load_cache() is None


def test_a_tampered_cache_cannot_repaint_the_terminal(repos):
    """The live path sanitises commit subjects; the cache is replayed through
    exactly the same banner on the next launch."""
    _, clone = repos
    (clone / update.CACHE_NAME).write_text(json.dumps({
        "state": "behind", "behind": 1,
        "upstream": "origin/main\x1b[2J\x1b[H  [OK] Up to date",
        "local": "aaaaaaa", "remote": "bbbbbbb",
        "checked_at": time.time()}))
    loaded = update.load_cache()
    assert loaded is not None
    assert "\x1b" not in loaded.upstream
    assert "\x1b" not in update.banner(loaded)


def test_the_truncation_marker_is_ascii(repos):
    """The module prints on consoles that predate UTF-8."""
    marked = update._printable("y" * 400)
    assert marked.endswith("...")
    marked.encode("ascii")               # would raise if '…' crept back in


# --- two generations of the same check ----------------------------------------

def test_a_superseded_check_cannot_publish_over_a_newer_one(repos):
    """A check that began before an update can finish after it. Its answer is
    about the old commit, so announcing it would offer an update already
    installed."""
    seed, _ = repos
    _push_commit(seed, "the fix")
    old = update.check(fetch=True)
    assert old.state == update.BEHIND

    update.begin_background_check()
    _quiesce()
    update.reset_background_state()      # stands in for "generation moved on"
    assert update._publish(old, generation=-1) is False
    assert update.latest_status() is None


def test_forcing_a_recheck_while_one_runs_does_not_lose_the_thread(repos):
    """join_background is what the test suite relies on to keep a daemon
    fetch from escaping onto the real repository — it has to cover the
    threads a force=True restart superseded, not just the newest."""
    update.begin_background_check()
    update.begin_background_check(force=True)
    with update._lock:
        assert len(update._threads) == 2
    assert update.join_background(30) is True


def test_joining_a_thread_that_never_started_is_not_an_error(repos):
    """There is a window where the handle is published but start() has not
    run — and start() itself can fail under thread exhaustion."""
    never_started = threading.Thread(target=lambda: None)
    with update._lock:
        update._threads.append(never_started)
    assert update.join_background(1) is True     # not RuntimeError


# --- pip ----------------------------------------------------------------------

def test_no_user_retry_under_conda(repos, monkeypatch):
    """~/.local IS on sys.path for conda interpreters, so a --user install
    shadows the conda-managed package for every env of that version."""
    monkeypatch.setenv("CONDA_PREFIX", "/opt/conda/envs/media")
    monkeypatch.setattr(update, "_in_virtualenv", lambda: False)
    calls = []

    def fake_run(cmd, **kwargs):
        calls.append(cmd)
        raise subprocess.CalledProcessError(1, cmd)

    monkeypatch.setattr(update.subprocess, "run", fake_run)
    assert update._pip_install(Path("requirements.txt")) is False
    assert len(calls) == 1
    assert not any("--user" in c for c in calls)


def test_the_pip_bound_covers_the_whole_call_not_each_attempt(repos, monkeypatch):
    timeouts = []
    spent = 0.25            # comfortably above Windows' ~16ms clock resolution

    def fake_run(cmd, **kwargs):
        timeouts.append(kwargs['timeout'])
        time.sleep(spent)   # the first attempt consumes part of the budget
        raise subprocess.CalledProcessError(1, cmd)

    monkeypatch.setattr(update.subprocess, "run", fake_run)
    monkeypatch.setattr(update, "_user_site_usable", lambda: True)
    update._pip_install(Path("requirements.txt"), timeout=5)

    assert len(timeouts) == 2
    assert timeouts[0] <= 5
    # The retry gets what is left of the same budget, not a second full one.
    assert timeouts[1] <= 5 - spent + 0.05


# --- a requirements.txt that is not UTF-8 -------------------------------------

def test_a_non_utf8_requirements_file_does_not_kill_the_update(repos, monkeypatch):
    """The fingerprint is taken after the pull has already succeeded. An
    exception there loses the report, the cache write and — from the wizard —
    the whole session.

    The bad bytes have to arrive *through the pull*: writing them into the
    clone and resetting would just restore the tracked UTF-8 file, and the
    test would pass with the bug still in place.
    """
    seed, clone = repos
    installs = []
    monkeypatch.setattr(update, "_pip_install",
                        lambda req, **kw: installs.append(req) or True)
    (seed / "run.py").write_text("print('v2')\n")
    # UTF-16 with a BOM is what PowerShell 5.1's `pip freeze >` produces. The
    # BOM is the part that matters: bare UTF-16LE of ASCII happens to be valid
    # UTF-8 (the NUL bytes decode), so it would not prove anything.
    (seed / "requirements.txt").write_bytes("pandas>=2.2\n".encode("utf-16"))
    _git(seed, "add", "-A")
    _git(seed, "commit", "-m", "utf-16 requirements")
    _git(seed, "push", "origin", "main")

    assert update.run_update(assume_yes=True) == 0
    assert (clone / "run.py").read_text() == "print('v2')\n"
    # The fingerprint changed, so the reinstall branch ran rather than
    # exploding on the way to it.
    assert len(installs) == 1


def test_a_missing_requirements_file_is_not_an_error(repos):
    _, clone = repos
    (clone / "requirements.txt").unlink()
    assert update._requirements_fingerprint() is None


# --- output the user needs in full --------------------------------------------

def test_the_pull_summary_keeps_its_lines(repos, capsys):
    seed, _ = repos
    _push_commit(seed, "the fix", body="print('v2')\n")
    update.run_update(assume_yes=True)
    out = capsys.readouterr().out
    assert "Fast-forward" in out
    # Sanitising the whole blob at once turned the diffstat into a single
    # run-on line beginning "Updating", so no line started with the file name.
    assert "Updating" in out
    assert any(line.strip().startswith("run.py") for line in out.splitlines())


def test_a_failed_pull_reports_the_whole_error(repos, monkeypatch, capsys):
    seed, _ = repos
    _push_commit(seed, "the fix")
    long_error = "\n".join(f"    path/to/file_{i}.mkv" for i in range(60))
    real_git = update._git

    def failing_pull(*args, **kw):
        if args and args[0] == "pull":
            return 1, "", "error: Your local changes would be overwritten by:\n" + long_error
        return real_git(*args, **kw)

    monkeypatch.setattr(update, "_git", failing_pull)
    assert update.run_update(assume_yes=True) == 1
    out = capsys.readouterr().out
    assert "file_59.mkv" in out            # not truncated at 200 characters


# --- dispatch contract --------------------------------------------------------

def test_a_malformed_flag_returns_a_code_rather_than_exiting():
    """argparse exits the process by default; this function is documented to
    return int | None, and tests and other callers rely on that."""
    assert update.dispatch_cli(["--update", "--dry-run=1"]) == 2
    assert update.dispatch_cli(["-V", "-yz"]) == 2


def test_the_deps_failure_survives_an_unmoved_head(repos, monkeypatch, capsys):
    """Someone else can pull during the confirmation prompt: HEAD does not
    move, requirements.txt did, pip fails — that is not 'nothing changed'."""
    seed, clone = repos
    _push_commit(seed, "newer deps", path="requirements.txt", body="pandas>=2.3\n")
    monkeypatch.setattr(update, "_pip_install", lambda req, **kw: False)
    real_git = update._git

    def pull_then_stand_still(*args, **kw):
        if args and args[0] == "pull":
            # Simulate the file arriving without HEAD appearing to move.
            (clone / "requirements.txt").write_text("pandas>=2.3\n")
            return 0, "Already up to date.", ""
        return real_git(*args, **kw)

    monkeypatch.setattr(update, "_git", pull_then_stand_still)
    code = update.run_update(assume_yes=True)
    out = capsys.readouterr().out
    assert code == 1
    assert "Already up to date - nothing changed" not in out


def test_update_module_pulls_in_no_third_party_packages():
    """run.py imports this before installing dependencies, on every launch."""
    probe = subprocess.run(
        [sys.executable, "-c",
         "import mediaorg.update, sys;"
         "print(sorted(m for m in sys.modules"
         " if m in ('pandas', 'guessit', 'openpyxl', 'tqdm', 'numpy')))"],
        cwd=str(Path(__file__).resolve().parent.parent),
        capture_output=True, text=True, check=True)
    assert probe.stdout.strip() == "[]"


# --- the cache contract with its own future -----------------------------------

def test_a_saved_status_can_be_loaded_back():
    """Fail-closed on unknown annotations is right, but it couples every
    future field to that loop: add one the loop cannot classify and every
    cache this version writes becomes unreadable, silently and with no error
    printed anywhere. This round-trip is what makes that fail in CI instead."""
    st = update.UpdateStatus(
        state=update.BEHIND, behind=1, ahead=0, branch="main",
        upstream="origin/main", local="aaaaaaa", remote="bbbbbbb",
        version="1.0", commits=["aaaaaaa subject"], reason="", hint="",
        checked_at=time.time())
    assert update.UpdateStatus.from_dict(st.to_dict()) == st


def test_a_cache_cannot_smuggle_in_unbounded_output(repos):
    _, clone = repos
    (clone / update.CACHE_NAME).write_text(json.dumps({
        "state": "behind", "behind": 1, "upstream": "origin/main",
        "commits": [f"sha subject {i}" for i in range(100_000)],
        "checked_at": time.time()}))
    loaded = update.load_cache()
    assert loaded is not None
    assert len(loaded.commits) == update.MAX_COMMITS


# --- output stays bounded as well as structured -------------------------------

def test_output_is_capped_by_lines_not_just_by_line_length():
    """Sanitising per line fixed the diffstat but removed the only bound on
    how much gets printed; a pull touching 50k files should not print 50k
    lines, and git's stderr is partly chosen by the remote."""
    block = update._printable_block("\n".join(f"line {i}" for i in range(5000)))
    lines = block.splitlines()
    assert len(lines) == update.MAX_OUTPUT_LINES + 1
    assert lines[-1].strip() == f"... and {5000 - update.MAX_OUTPUT_LINES} more lines"


def test_a_short_block_is_passed_through_whole():
    block = update._printable_block("Updating a1b2c3d..e4f5g6h\nFast-forward\n")
    assert block.splitlines() == ["Updating a1b2c3d..e4f5g6h", "Fast-forward"]


# --- thread registration edges ------------------------------------------------

def test_finished_threads_do_not_accumulate(repos):
    for _ in range(4):
        update.begin_background_check(force=True)
        _quiesce()
    update.begin_background_check(force=True)
    with update._lock:
        assert len(update._threads) <= 2      # the live one, not every one ever
    _quiesce()


def test_a_thread_that_cannot_start_leaves_the_module_usable(repos, monkeypatch):
    """RuntimeError from start() used to wedge things: a dead handle passed
    the alive check, _started blocked the retry, and an Event nobody would
    ever set made every wait_for_cache pay its full timeout."""
    def refuse(self):
        raise RuntimeError("can't start new thread")

    monkeypatch.setattr(threading.Thread, "start", refuse)
    update.begin_background_check()

    started = time.monotonic()
    assert update.wait_for_cache(5) is None
    assert time.monotonic() - started < 1      # not the full timeout
    monkeypatch.undo()

    update.begin_background_check()            # and a later attempt still works
    assert update.wait_for_check(30) is not None


# --- the foreground answer wins ----------------------------------------------

def test_an_explicit_check_retires_a_running_background_one(repos):
    """--check-update is newer than a check already in flight; letting the
    older one land on top is the race generations exist to settle."""
    seed, _ = repos
    _push_commit(seed, "the fix")
    stale = update.check(fetch=True)
    assert stale.state == update.BEHIND

    update.begin_background_check()
    _quiesce()
    with update._lock:
        generation = update._generation
    update.check_and_cache(fetch=False)
    with update._lock:
        assert update._generation > generation
    assert update._publish(stale, generation) is False
