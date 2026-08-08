"""Tests for the self-update check and the update itself.

Every test builds a real repository pair on disk — a bare "origin" and a
clone standing in for the user's install — because the whole feature is an
opinion about what git says, and a mocked git would only test the opinion.
"""

import json
import shutil
import subprocess
import time

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
    return seed, clone


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
    time.sleep(0.2)
    assert called == []
    assert update.latest_status().state == update.CURRENT


def test_stale_cache_triggers_a_refresh(repos):
    seed, _ = repos
    stale = update.check(fetch=True)
    stale.checked_at = time.time() - 999 * 3600
    update.save_cache(stale)
    _push_commit(seed, "found on refresh")

    update.begin_background_check()
    for _ in range(100):                              # the thread is daemonic
        if (update.latest_status() or stale).state == update.BEHIND:
            break
        time.sleep(0.1)
    assert update.latest_status().state == update.BEHIND


def test_offline_check_does_not_overwrite_a_good_cached_answer(repos, monkeypatch):
    seed, _ = repos
    _push_commit(seed, "known update")
    good = update.check(fetch=True)
    good.checked_at = time.time() - 999 * 3600
    update.save_cache(good)

    monkeypatch.setattr(update, "check", lambda **kw: update.UpdateStatus(
        state=update.UNKNOWN, reason="could not reach the remote"))
    update.begin_background_check()
    time.sleep(0.3)
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
