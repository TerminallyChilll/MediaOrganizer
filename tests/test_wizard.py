"""Tests for the wizard-level features: inventory, the custom word list, and
the LLM configuration. These are the parts a user reaches from the menu, and
none of them had any coverage before."""

import json
import re

import pytest

from mediaorg import llm, update, wizard
from mediaorg.parse import load_custom_patterns, save_custom_patterns


# --- Inventory ---------------------------------------------------------------

@pytest.fixture
def library(tmp_path):
    """A folder holding media, companions and unrelated files."""
    root = tmp_path / "Library"
    (root / "Show" / "Season 1").mkdir(parents=True)
    (root / "Show" / "Season 1" / "Show.S01E01.mkv").write_text("video")
    (root / "Show" / "Season 1" / "Show.S01E01.srt").write_text("subs")
    (root / "Show" / "poster.jpg").write_text("art")
    (root / "receipt.pdf").write_text("not media at all")
    (root / "notes.txt").write_text("notes")
    return root


def test_inventory_lists_every_file_not_just_media(library):
    rows, errors = wizard.collect_inventory(library)
    assert errors == []
    assert {r['Path'] for r in rows} == {
        "Show/Season 1/Show.S01E01.mkv",
        "Show/Season 1/Show.S01E01.srt",
        "Show/poster.jpg",
        "receipt.pdf",
        "notes.txt",
    }


def test_inventory_classifies_rather_than_filters(library):
    rows, _ = wizard.collect_inventory(library)
    by_name = {r['File Name']: r for r in rows}
    assert by_name["Show.S01E01.mkv"]['Type'] == "video"
    assert by_name["Show.S01E01.srt"]['Type'] == "companion"
    # A .pdf is in neither media list, and must still be inventoried.
    assert by_name["receipt.pdf"]['Type'] == "other"


def test_inventory_records_size_and_location(library):
    rows, _ = wizard.collect_inventory(library)
    row = next(r for r in rows if r['File Name'] == "Show.S01E01.mkv")
    assert row['Folder'] == "Show/Season 1"
    assert row['Extension'] == ".mkv"
    assert row['Size (bytes)'] == len("video")
    assert row['Modified']


@pytest.mark.parametrize("suffix", [".xlsx", ".csv", ".txt"])
def test_inventory_writes_every_format(library, tmp_path, suffix):
    out = tmp_path / f"inventory{suffix}"
    assert wizard.write_inventory(library, out) is True
    assert out.exists() and out.stat().st_size > 0


def test_inventory_changes_nothing_on_disk(library):
    before = {str(p) for p in library.rglob("*")}
    wizard.collect_inventory(library)
    assert {str(p) for p in library.rglob("*")} == before


def test_inventory_of_an_empty_folder_reports_rather_than_writing(tmp_path):
    empty = tmp_path / "Empty"
    empty.mkdir()
    out = tmp_path / "inventory.csv"
    assert wizard.write_inventory(empty, out) is False
    assert not out.exists()


def test_inventory_dry_run_writes_nothing(library, tmp_path):
    out = tmp_path / "inventory.csv"
    assert wizard.write_inventory(library, out, dry_run=True) is True
    assert not out.exists()


# --- Custom word list --------------------------------------------------------

@pytest.fixture
def words(tmp_path, monkeypatch):
    monkeypatch.setenv("MEDIAORG_PATTERNS", str(tmp_path / "words.json"))
    return tmp_path / "words.json"


def _drive(monkeypatch, answers):
    """Feed the wizard a fixed sequence of prompt answers.

    Feeding answers *is* simulating a terminal, so this also asserts the
    interactive path. Under pytest stdin is not a tty, and without this the
    wizard would correctly take its unattended branch and never read them.
    """
    supplied = iter(answers)
    monkeypatch.setattr("builtins.input", lambda *a: next(supplied))
    monkeypatch.setattr(wizard, "_stdin_is_interactive", lambda: True)


def test_add_a_word(words, monkeypatch, capsys):
    _drive(monkeypatch, ["a", "RARBG", "q"])
    wizard.run_custom_words()
    assert load_custom_patterns() == ["RARBG"]


def test_remove_a_word_by_number(words, monkeypatch):
    save_custom_patterns(["RARBG", "YIFY"])
    _drive(monkeypatch, ["r", "1", "q"])
    wizard.run_custom_words()
    assert load_custom_patterns() == ["YIFY"]


def test_remove_a_word_by_name(words, monkeypatch):
    save_custom_patterns(["RARBG", "YIFY"])
    _drive(monkeypatch, ["r", "YIFY", "q"])
    wizard.run_custom_words()
    assert load_custom_patterns() == ["RARBG"]


def test_clear_the_whole_list(words, monkeypatch):
    save_custom_patterns(["RARBG", "YIFY"])
    _drive(monkeypatch, ["c", "y", "q"])
    wizard.run_custom_words()
    assert load_custom_patterns() == []


def test_clearing_can_be_declined(words, monkeypatch):
    save_custom_patterns(["RARBG"])
    _drive(monkeypatch, ["c", "n", "q"])
    wizard.run_custom_words()
    assert load_custom_patterns() == ["RARBG"]


def test_a_duplicate_is_not_added_twice(words, monkeypatch):
    save_custom_patterns(["RARBG"])
    _drive(monkeypatch, ["a", "RARBG", "q"])
    wizard.run_custom_words()
    assert load_custom_patterns() == ["RARBG"]


def test_an_invalid_pattern_is_offered_escaped(words, monkeypatch):
    _drive(monkeypatch, ["a", "[unclosed", "y", "q"])
    wizard.run_custom_words()
    # Accepted as a literal, so it compiles and matches the text itself.
    assert load_custom_patterns() == ["\\[unclosed"]


def test_a_pattern_matching_whole_names_is_refused(words, monkeypatch):
    """pre_clean declines to apply one of these, so it would sit in the list
    doing nothing at all."""
    _drive(monkeypatch, ["a", ".*", "a", ".+", "q"])
    wizard.run_custom_words()
    assert load_custom_patterns() == []


def test_added_words_actually_strip(words, monkeypatch):
    _drive(monkeypatch, ["a", "RARBG", "q"])
    wizard.run_custom_words()
    from mediaorg.parse import parse_name
    parsed = parse_name("My.Movie.2020.1080p-RARBG.mkv",
                        custom_patterns=load_custom_patterns())
    assert parsed.title == "My Movie" and parsed.year == 2020


def test_removed_words_stop_stripping(words, monkeypatch):
    save_custom_patterns(["Whatever"])
    _drive(monkeypatch, ["r", "1", "q"])
    wizard.run_custom_words()
    assert load_custom_patterns() == []


# --- LLM configuration -------------------------------------------------------

def test_env_vars_supply_the_llm_config(tmp_path, monkeypatch):
    """docker-compose.yml has always passed these, but nothing read them."""
    monkeypatch.setenv("MEDIAORG_LLM_CONFIG", str(tmp_path / "llm.json"))
    monkeypatch.setenv("OPENAI_API_KEY", "sk-from-env")
    monkeypatch.setenv("OLLAMA_URL", "http://nas:11434")
    cfg = llm.load_llm_config()
    assert cfg["openai_key"] == "sk-from-env"
    assert cfg["ollama_url"] == "http://nas:11434"


def test_loading_alone_never_creates_the_config_file(tmp_path, monkeypatch):
    target = tmp_path / "llm.json"
    monkeypatch.setenv("MEDIAORG_LLM_CONFIG", str(target))
    monkeypatch.setenv("GEMINI_API_KEY", "secret")
    llm.load_llm_config()
    assert not target.exists()


def test_an_env_key_is_never_written_to_disk(tmp_path, monkeypatch):
    """The save path is what matters here, not the load path.

    load_llm_config overlays every env value into the dict it returns, so any
    caller that saves that dict back copied those values into plaintext. The
    earlier version of this test only called load_llm_config and checked the
    file was not created — so it passed while the guarantee was being broken
    one line later.
    """
    target = tmp_path / "llm.json"
    monkeypatch.setenv("MEDIAORG_LLM_CONFIG", str(target))
    monkeypatch.setenv("OPENAI_API_KEY", "sk-openai-from-env")
    monkeypatch.setenv("GEMINI_API_KEY", "gk-gemini-from-env")

    llm.save_llm_config(llm.load_llm_config())

    written = json.loads(target.read_text())
    assert "openai_key" not in written
    assert "gemini_key" not in written
    assert "sk-openai-from-env" not in target.read_text()


def test_configuring_ollama_does_not_leak_a_cloud_key(tmp_path, monkeypatch):
    """The exact shape of the wizard's Ollama branch: env keys are present in
    cfg, and it saves the whole dict after adding the url and model."""
    target = tmp_path / "llm.json"
    monkeypatch.setenv("MEDIAORG_LLM_CONFIG", str(target))
    monkeypatch.setenv("OPENAI_API_KEY", "sk-openai-from-env")

    cfg = llm.load_llm_config()
    cfg.update({"ollama_url": "http://localhost:11434",
                "ollama_model": "llama3"})
    llm.save_llm_config(cfg)

    written = json.loads(target.read_text())
    assert written == {"ollama_url": "http://localhost:11434",
                       "ollama_model": "llama3"}


def test_a_typed_key_is_still_saved_when_another_comes_from_the_env(tmp_path, monkeypatch):
    """Only the env-supplied value is dropped — a key the user typed is
    theirs, and must survive so they are not asked for it again."""
    target = tmp_path / "llm.json"
    monkeypatch.setenv("MEDIAORG_LLM_CONFIG", str(target))
    monkeypatch.setenv("OPENAI_API_KEY", "sk-openai-from-env")

    cfg = llm.load_llm_config()
    cfg["gemini_key"] = "gk-typed-by-hand"
    llm.save_llm_config(cfg)

    written = json.loads(target.read_text())
    assert written == {"gemini_key": "gk-typed-by-hand"}


def test_an_interactively_chosen_value_survives_even_if_the_env_sets_it(tmp_path, monkeypatch):
    """Only an exact match is dropped. Picking a different Ollama URL than
    the environment's is a deliberate choice and has to persist."""
    target = tmp_path / "llm.json"
    monkeypatch.setenv("MEDIAORG_LLM_CONFIG", str(target))
    monkeypatch.setenv("OLLAMA_URL", "http://from-env:11434")

    cfg = llm.load_llm_config()
    cfg["ollama_url"] = "http://chosen-by-user:11434"
    llm.save_llm_config(cfg)

    assert json.loads(target.read_text()) == {
        "ollama_url": "http://chosen-by-user:11434"}


def test_a_blank_env_var_does_not_wipe_a_saved_key(tmp_path, monkeypatch):
    """docker-compose passes an unset variable through as an empty string."""
    target = tmp_path / "llm.json"
    monkeypatch.setenv("MEDIAORG_LLM_CONFIG", str(target))
    llm.save_llm_config({"openai_key": "sk-saved"})
    monkeypatch.setenv("OPENAI_API_KEY", "")
    assert llm.load_llm_config()["openai_key"] == "sk-saved"


def test_the_llm_config_is_anchored_to_the_app(tmp_path, monkeypatch):
    """Cwd-relative meant a key saved in one directory was invisible when the
    app was launched from another, so the prompt came back every run."""
    target = tmp_path / "elsewhere.json"
    monkeypatch.setenv("MEDIAORG_LLM_CONFIG", str(target))
    llm.save_llm_config({"openai_key": "sk-1"})
    monkeypatch.chdir(tmp_path)
    assert llm.load_llm_config()["openai_key"] == "sk-1"
    assert json.loads(target.read_text())["openai_key"] == "sk-1"


def test_saved_key_file_is_not_world_readable(tmp_path, monkeypatch):
    import os
    import stat as stat_mod
    if os.name == "nt":
        pytest.skip("POSIX permission bits only")
    target = tmp_path / "llm.json"
    monkeypatch.setenv("MEDIAORG_LLM_CONFIG", str(target))
    llm.save_llm_config({"openai_key": "sk-1"})
    mode = stat_mod.S_IMODE(target.stat().st_mode)
    assert mode & (stat_mod.S_IRGRP | stat_mod.S_IROTH) == 0


def test_llm_response_parsing_tolerates_real_model_output():
    """Local models wrap JSON in fences, use single quotes, add a wrapper key
    and trailing commas. All of that has to survive."""
    names = ["a.mkv"]
    fenced = '```json\n[{"original": "a.mkv", "title": "A", "year": "2001"}]\n```'
    assert llm._parse_llm_response(fenced, names)["a.mkv"]["title"] == "A"

    wrapped = '{"results": [{"original": "a.mkv", "title": "A"}]}'
    assert llm._parse_llm_response(wrapped, names)["a.mkv"]["title"] == "A"

    single = "[{'original': 'a.mkv', 'title': 'A', 'year': '2001',}]"
    assert llm._parse_llm_response(single, names)["a.mkv"]["title"] == "A"

    lone = '{"original": "a.mkv", "title": "A"}'
    assert llm._parse_llm_response(lone, names)["a.mkv"]["title"] == "A"

    # Alternative field names the prompt does not ask for but models emit.
    alt = '[{"filename": "a.mkv", "clean_title": "A", "resolution": "1080p"}]'
    assert llm._parse_llm_response(alt, names)["a.mkv"]["quality"] == "1080p"

    assert llm._parse_llm_response("", names) == {}
    assert llm._parse_llm_response("I cannot help with that", names) == {}


def test_ollama_round_trip_against_a_live_stub_server():
    """Exercises the real HTTP path: model listing, the chat call, response
    parsing, and batching — not just the parser in isolation."""
    import json as _json
    import threading
    from http.server import BaseHTTPRequestHandler, HTTPServer

    seen = []

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass

        def _send(self, payload):
            body = _json.dumps(payload).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self):
            self._send({"models": [{"name": "llama3:8b"}, {"name": "qwen2.5"}]})

        def do_POST(self):
            request = _json.loads(self.rfile.read(
                int(self.headers.get("Content-Length", 0))))
            seen.append(request)
            asked = request["messages"][-1]["content"]
            answer = [{"original": n, "title": "Cleaned " + n.split(".")[0],
                       "year": "1999", "quality": "1080p"}
                      for n in ("one.mkv", "two.mkv") if n in asked]
            self._send({"message": {"content": _json.dumps(answer)}})

    server = HTTPServer(("127.0.0.1", 0), Handler)
    url = f"http://127.0.0.1:{server.server_address[1]}"
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        assert llm.list_ollama_models(url) == ["llama3:8b", "qwen2.5"]
        out = llm.clean_titles_with_llm(["one.mkv", "two.mkv"], "ollama",
                                        model="llama3:8b", ollama_url=url)
    finally:
        server.shutdown()

    assert out["one.mkv"] == {"title": "Cleaned one", "year": "1999",
                              "quality": "1080p"}
    assert out["two.mkv"]["title"] == "Cleaned two"
    assert seen[0]["model"] == "llama3:8b"
    assert seen[0]["format"] == "json", "JSON mode keeps local models on task"
    assert seen[0]["stream"] is False


def test_unreachable_ollama_reports_no_models_rather_than_crashing():
    assert llm.list_ollama_models("http://127.0.0.1:9") == []


def test_an_unwritable_word_list_reports_instead_of_crashing(tmp_path, monkeypatch, capsys):
    """The app directory can legitimately be unwritable (a system-wide
    install, a read-only container mount). An unhandled OSError here unwound
    out of the whole menu.

    A regular file stands in for the unwritable location: chmod is no
    obstacle to root, which is how the container test suite runs.
    """
    blocker = tmp_path / "not-a-directory"
    blocker.write_text("this is a file, so it cannot contain the word list")
    monkeypatch.setenv("MEDIAORG_PATTERNS", str(blocker / "words.json"))
    _drive(monkeypatch, ["a", "RARBG", "q"])
    wizard.run_custom_words()              # must return, not raise
    assert "Could not write" in capsys.readouterr().out


def test_a_missing_parent_directory_is_created(tmp_path, monkeypatch):
    target = tmp_path / "new" / "deeper" / "words.json"
    monkeypatch.setenv("MEDIAORG_PATTERNS", str(target))
    _drive(monkeypatch, ["a", "RARBG", "q"])
    wizard.run_custom_words()
    assert load_custom_patterns() == ["RARBG"]


# --- Update notice in the menu -----------------------------------------------

@pytest.fixture
def quiet_update_check(monkeypatch):
    """The menu kicks off a network check on entry; keep tests offline."""
    monkeypatch.setattr(update, "begin_background_check", lambda **kw: None)
    monkeypatch.setattr(update, "latest_status", lambda: None)


def test_menu_is_silent_when_there_is_no_update(quiet_update_check, monkeypatch, capsys):
    _drive(monkeypatch, ["0"])
    wizard.run_wizard()
    out = capsys.readouterr().out
    assert "Update available" not in out
    assert "[U] Update Media Organizer" in out      # the option is still offered


def test_menu_shows_the_update_command_when_behind(quiet_update_check, monkeypatch, capsys):
    monkeypatch.setattr(update, "latest_status", lambda: update.UpdateStatus(
        state=update.BEHIND, behind=3, upstream="origin/main",
        local="aaaaaaa", remote="bbbbbbb"))
    _drive(monkeypatch, ["0"])
    wizard.run_wizard()
    out = capsys.readouterr().out
    assert "3 commits behind origin/main" in out
    assert "python run.py --update" in out


def _fake_update(monkeypatch, revisions, code=0):
    """Stand in for an update that moves HEAD through *revisions*."""
    seen = iter(revisions)
    monkeypatch.setattr(update, "head_revision", lambda: next(seen))
    monkeypatch.setattr(update, "run_update", lambda **kw: code)


def test_menu_u_runs_the_update_and_exits_when_it_pulled(quiet_update_check,
                                                         monkeypatch, capsys):
    """After a successful pull the files on disk no longer match the modules
    this process imported, so the wizard must hand back to a fresh launch."""
    _fake_update(monkeypatch, ["aaaaaaa", "bbbbbbb"])
    _drive(monkeypatch, ["u"])                       # no "0" — it must exit itself
    wizard.run_wizard()
    assert "Exiting so the new version is loaded" in capsys.readouterr().out


def test_menu_u_returns_to_the_menu_when_already_current(quiet_update_check,
                                                         monkeypatch):
    _fake_update(monkeypatch, ["aaaaaaa", "aaaaaaa"])
    _drive(monkeypatch, ["U", "0"])                  # must still be here for the "0"
    wizard.run_wizard()


def test_menu_u_exits_even_when_the_local_refs_looked_up_to_date(
        quiet_update_check, monkeypatch, capsys):
    """A clone that has never fetched reads as up to date offline. The pull
    still moves HEAD, so the decision to relaunch must come from HEAD, not
    from a status measured before the fetch."""
    monkeypatch.setattr(update, "check",
                        lambda **kw: update.UpdateStatus(state=update.CURRENT))
    _fake_update(monkeypatch, ["aaaaaaa", "bbbbbbb"])
    _drive(monkeypatch, ["u"])
    wizard.run_wizard()
    assert "Exiting so the new version is loaded" in capsys.readouterr().out


def test_menu_u_stays_put_when_the_update_was_refused(quiet_update_check,
                                                      monkeypatch, capsys):
    """A dirty work tree stops the update; nothing was pulled, so the wizard
    keeps running rather than sending the user off to relaunch."""
    _fake_update(monkeypatch, ["aaaaaaa", "aaaaaaa"], code=1)   # HEAD never moved
    _drive(monkeypatch, ["u", "0"])
    wizard.run_wizard()
    assert "Exiting so the new version is loaded" not in capsys.readouterr().out


def test_the_full_notice_appears_once_then_shrinks(quiet_update_check, monkeypatch,
                                                   capsys):
    """Fourteen lines between the header and the menu on every redraw stops
    being information and becomes wallpaper."""
    monkeypatch.setattr(update, "latest_status", lambda: update.UpdateStatus(
        state=update.BEHIND, behind=2, upstream="origin/main",
        local="aaaaaaa", remote="bbbbbbb"))
    _drive(monkeypatch, ["x", "0"])              # invalid choice, then quit
    wizard.run_wizard()
    out = capsys.readouterr().out
    assert out.count("To update, run this in a terminal") == 1
    assert out.count("press [U] to install it") == 1


def test_a_broken_update_check_never_stops_the_app_starting(monkeypatch, capsys):
    def boom(*a, **kw):
        raise RuntimeError("cache is a smoking crater")

    monkeypatch.setattr(update, "begin_background_check", boom)
    monkeypatch.setattr(update, "latest_status", boom)
    _drive(monkeypatch, ["0"])
    wizard.run_wizard()                          # must reach the menu, not raise
    assert "[1] Clean file names" in capsys.readouterr().out


def test_the_launch_never_pays_the_update_budget_twice(monkeypatch):
    """`wait_for_cache` returning None is ambiguous — nothing remembered, or
    not published yet — so the two waits share one deadline rather than
    charging the user for both."""
    import time as _time
    spent = 0.4
    waits = []
    monkeypatch.setattr(update, "begin_background_check", lambda **kw: None)

    def slow_cache_wait(timeout):
        waits.append(timeout)
        _time.sleep(spent)          # the local phase used part of the budget
        return None

    monkeypatch.setattr(update, "wait_for_cache", slow_cache_wait)
    monkeypatch.setattr(update, "wait_for_check", lambda t: waits.append(t))
    started = _time.monotonic()
    wizard._start_update_check()

    assert waits[0] == wizard.LAUNCH_CHECK_BUDGET
    # The second wait gets what is left of the same budget, not a fresh one.
    assert waits[1] <= wizard.LAUNCH_CHECK_BUDGET - spent + 0.05
    assert _time.monotonic() - started < wizard.LAUNCH_CHECK_BUDGET


# ── where the remembered folders live ────────────────────────────────────

def test_remembered_folders_follow_the_app_not_the_cwd(tmp_path, monkeypatch):
    """The config was the last state file still resolved against the cwd, so
    launching from anywhere else silently forgot the folders you picked."""
    monkeypatch.delenv("MEDIAORG_CONFIG", raising=False)
    app = wizard.Path(wizard.__file__).resolve().parent.parent

    monkeypatch.chdir(tmp_path)
    assert wizard.config_path() == app / wizard.CONFIG_FILE


def test_the_config_location_can_be_pinned(tmp_path, monkeypatch):
    pinned = tmp_path / "elsewhere" / "cfg.json"
    monkeypatch.setenv("MEDIAORG_CONFIG", str(pinned))
    assert wizard.config_path() == pinned

    pinned.parent.mkdir(parents=True)
    wizard._save_config({"tv": "/media/TV"})
    assert wizard._load_config() == {"tv": "/media/TV"}


def test_an_existing_cwd_config_is_adopted_not_orphaned(tmp_path, monkeypatch):
    """Upgrading users keep what they already have: a config written by an
    older version sits in the launch directory, and is read rather than
    replaced by an empty one next to the app."""
    monkeypatch.delenv("MEDIAORG_CONFIG", raising=False)
    app_dir = tmp_path / "app"
    (app_dir / "mediaorg").mkdir(parents=True)
    monkeypatch.setattr(wizard, "__file__", str(app_dir / "mediaorg" / "wizard.py"))

    legacy = tmp_path / "launched-from"
    legacy.mkdir()
    (legacy / wizard.CONFIG_FILE).write_text('{"tv": "/old/TV"}', encoding="utf-8")
    monkeypatch.chdir(legacy)

    assert wizard.config_path() == legacy / wizard.CONFIG_FILE
    assert wizard._load_config() == {"tv": "/old/TV"}


def test_a_config_next_to_the_app_wins_over_one_in_the_cwd(tmp_path, monkeypatch):
    monkeypatch.delenv("MEDIAORG_CONFIG", raising=False)
    app_dir = tmp_path / "app"
    (app_dir / "mediaorg").mkdir(parents=True)
    monkeypatch.setattr(wizard, "__file__", str(app_dir / "mediaorg" / "wizard.py"))
    (app_dir / wizard.CONFIG_FILE).write_text('{"tv": "/new/TV"}', encoding="utf-8")

    other = tmp_path / "somewhere-else"
    other.mkdir()
    (other / wizard.CONFIG_FILE).write_text('{"tv": "/stale/TV"}', encoding="utf-8")
    monkeypatch.chdir(other)

    assert wizard._load_config() == {"tv": "/new/TV"}


def test_unreadable_config_is_not_fatal(tmp_path, monkeypatch):
    monkeypatch.setenv("MEDIAORG_CONFIG", str(tmp_path / "cfg.json"))
    (tmp_path / "cfg.json").write_text("{not json", encoding="utf-8")
    assert wizard._load_config() == {}


# --- Review and the accept-or-revert gate ------------------------------------
#
# These cover the promise the whole write path rests on: nothing is applied
# without being shown first, and anything the user declines afterwards is put
# back exactly as it was.

from pathlib import Path

from mediaorg.execute import list_runs
from mediaorg.plan import Op, Plan, check_collisions


@pytest.fixture
def wired(tmp_path, monkeypatch):
    """A library plus a journal, with the retry backoff turned off."""
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    journal = tmp_path / "journal.jsonl"
    monkeypatch.setenv("MEDIAORG_JOURNAL", str(journal))
    root = tmp_path / "TV" / "My Show"
    root.mkdir(parents=True)
    for name in ("My.Show.S01E01.mkv", "My.Show.S01E01.srt",
                 "My.Show.S01E02.mkv", "My.Show.S02E01.mkv"):
        (root / name).write_text(name)
    return tmp_path / "TV", journal


def _tree(root: Path) -> dict:
    return {str(p.relative_to(root)): (p.read_text() if p.is_file() else None)
            for p in sorted(Path(root).rglob("*"))}


def test_declining_afterwards_puts_every_file_back(wired, monkeypatch):
    tv, journal = wired
    before = _tree(tv)
    _drive(monkeypatch, ["Y", "n"])          # apply as listed, then decline

    wizard.run_organize(str(tv))

    assert _tree(tv) == before
    runs = list_runs(journal)
    assert runs, "nothing was applied, so the revert proves nothing"
    assert all(r["undone"] for r in runs)


def test_accepting_afterwards_keeps_them(wired, monkeypatch):
    tv, journal = wired
    before = _tree(tv)
    _drive(monkeypatch, ["Y", "y"])

    wizard.run_organize(str(tv))

    assert _tree(tv) != before
    assert (tv / "My Show" / "Season 1" / "My.Show.S01E01.mkv").exists()
    assert not any(r["undone"] for r in list_runs(journal))


def test_a_bare_enter_at_the_gate_reverts(wired, monkeypatch):
    """The default must be the safe answer, not the destructive one."""
    tv, _ = wired
    before = _tree(tv)
    _drive(monkeypatch, ["Y", ""])

    wizard.run_organize(str(tv))

    assert _tree(tv) == before


def test_back_at_the_gate_reverts_instead_of_escaping(wired, monkeypatch):
    """'back' must not unwind to the menu leaving the library changed."""
    tv, _ = wired
    before = _tree(tv)
    _drive(monkeypatch, ["Y", "back"])

    wizard.run_organize(str(tv))       # must not raise BackNavigation

    assert _tree(tv) == before


def test_an_interrupt_at_the_gate_keeps_and_says_how_to_undo(wired, monkeypatch,
                                                             capsys):
    """Ctrl-C must not silently start a large unasked-for reversal."""
    tv, journal = wired
    answers = iter(["Y"])

    def _input(*a):
        try:
            return next(answers)
        except StopIteration:
            raise KeyboardInterrupt
    monkeypatch.setattr("builtins.input", _input)
    monkeypatch.setattr(wizard, "_stdin_is_interactive", lambda: True)

    with pytest.raises(KeyboardInterrupt):
        wizard.run_organize(str(tv))

    assert (tv / "My Show" / "Season 1" / "My.Show.S01E01.mkv").exists()
    out = capsys.readouterr().out
    run_id = list_runs(journal)[-1]["id"]
    assert f"--undo-run {run_id}" in out


def test_quitting_the_review_changes_nothing(wired, monkeypatch):
    tv, journal = wired
    before = _tree(tv)
    _drive(monkeypatch, ["Q"])

    wizard.run_organize(str(tv))

    assert _tree(tv) == before
    assert not journal.exists()


def test_excluding_an_item_leaves_that_file_alone(wired, monkeypatch):
    tv, _ = wired
    # review -> exclude item 1 -> apply -> keep
    _drive(monkeypatch, ["R", "x 1", "Y", "y"])

    wizard.run_organize(str(tv))

    show = tv / "My Show"
    assert (show / "My.Show.S01E01.mkv").exists()          # excluded, still put
    assert (show / "Season 1" / "My.Show.S01E02.mkv").exists()
    assert (show / "Season 2" / "My.Show.S02E01.mkv").exists()


def test_excluding_a_range(wired, monkeypatch):
    tv, _ = wired
    before = _tree(tv)
    _drive(monkeypatch, ["R", "x 1-4", "Y"])

    wizard.run_organize(str(tv))

    assert _tree(tv) == before      # nothing left to do -> nothing happened


def test_keep_puts_an_excluded_item_back_in(wired, monkeypatch):
    tv, _ = wired
    _drive(monkeypatch, ["R", "x 1-4", "k 1-4", "Y", "y"])

    wizard.run_organize(str(tv))

    assert (tv / "My Show" / "Season 1" / "My.Show.S01E01.mkv").exists()


def test_excluding_a_move_also_drops_the_folder_cleanup(wired, monkeypatch,
                                                        tmp_path):
    """A folder the user kept a file in must not be reported as removable."""
    tv, _ = wired
    nested = tv / "My Show" / "Season 1" / "Disc 1"
    nested.mkdir(parents=True)
    (nested / "My.Show.S01E03.mkv").write_text("ep3")
    (nested / "My.Show.S01E04.mkv").write_text("ep4")
    # Lift both episodes out of Disc 1, then exclude one of them: the rmdir of
    # Disc 1 can no longer succeed and must not be attempted.
    captured = {}
    real = wizard.execute

    def spy(plan, *a, **kw):
        captured['plan'] = plan
        return real(plan, *a, **kw)
    monkeypatch.setattr(wizard, "execute", spy)

    # ... "y" to "continue with the remaining?", then "y" to keep them.
    _drive(monkeypatch, ["R", "A", "x 1", "Y", "y", "y"])
    wizard.run_organize(str(tv))

    kept_rmdirs = [op.dst for op in captured['plan'].ops if op.kind == "rmdir"]
    assert nested not in kept_rmdirs
    assert (nested / "My.Show.S01E03.mkv").exists()   # the excluded one stayed
    assert real is not None


def test_renaming_an_item_carries_its_companions(wired, monkeypatch):
    """Editing a video's name must not orphan its subtitles."""
    tv, _ = wired
    # Item 1 and 2 are the .mkv and its .srt going into Season 1.
    _drive(monkeypatch, ["R", "e 1", "Pilot.mkv", "Y", "y"])

    wizard.run_organize(str(tv))

    season = tv / "My Show" / "Season 1"
    assert (season / "Pilot.mkv").exists()
    assert (season / "Pilot.srt").exists()
    assert (season / "Pilot.srt").read_text() == "My.Show.S01E01.srt"


def test_a_renamed_item_that_collides_is_refused_not_applied(wired, monkeypatch):
    """A hand-typed name is exactly as able to collide as a planned one."""
    tv, _ = wired
    # Rename item 3 (S01E02) onto item 1's destination name.
    _drive(monkeypatch, ["R", "e 3", "My.Show.S01E01.mkv", "Y", "n", "Q"])

    wizard.run_organize(str(tv))

    # The collision was caught before anything ran: declining to continue
    # leaves the library untouched.
    assert (tv / "My Show" / "My.Show.S01E02.mkv").exists()


def test_a_rename_that_would_drop_the_extension_is_questioned(wired, monkeypatch,
                                                              capsys):
    tv, _ = wired
    # 'n' to "really change the extension?" -> the .mkv is restored.
    _drive(monkeypatch, ["R", "e 1", "Pilot", "n", "Y", "y"])

    wizard.run_organize(str(tv))

    assert (tv / "My Show" / "Season 1" / "Pilot.mkv").exists()


def test_a_rename_with_a_path_separator_is_refused(wired, monkeypatch, capsys):
    """Editing changes the name, never the folder a file lands in."""
    tv, _ = wired
    _drive(monkeypatch, ["R", "e 1", "../escape.mkv", "Y", "y"])

    wizard.run_organize(str(tv))

    assert (tv / "My Show" / "Season 1" / "My.Show.S01E01.mkv").exists()
    assert "No slashes" in capsys.readouterr().out


def test_a_scripted_run_applies_without_the_gate(wired, monkeypatch, capsys):
    """--action without --review has nobody to answer the last question."""
    tv, journal = wired
    _drive(monkeypatch, ["Y"])          # only the pre-apply review is answered

    wizard.run_organize(str(tv), accept_gate=False)

    assert (tv / "My Show" / "Season 1" / "My.Show.S01E01.mkv").exists()
    assert not any(r["undone"] for r in list_runs(journal))
    assert "python run.py --undo" in capsys.readouterr().out


def test_a_dry_run_asks_nothing_and_changes_nothing(wired, monkeypatch):
    tv, journal = wired
    before = _tree(tv)
    monkeypatch.setattr("builtins.input",
                        lambda *a: pytest.fail("a dry run must not prompt"))

    wizard.run_organize(str(tv), dry_run=True)

    assert _tree(tv) == before
    assert not journal.exists()


def test_declining_reverses_every_phase_of_one_action(wired, monkeypatch):
    """[4]/--action full: "no" puts back the organize as well as the rename."""
    tv, journal = wired
    before = _tree(tv)
    session = "deadbeefcafe"

    _drive(monkeypatch, ["Y"])
    wizard.run_organize(str(tv), session=session, accept_gate=False)
    assert _tree(tv) != before

    # A second phase under the same session id.
    show = tv / "My Show"
    extra = Plan(ops=[Op("move", show / "Season 2" / "My.Show.S02E01.mkv",
                         show / "Season 2" / "Renamed.mkv")])
    _drive(monkeypatch, ["Y"])
    wizard.confirm_and_execute(extra, journal, label="renames", roots=[tv],
                               session=session, accept_gate=False)
    assert (show / "Season 2" / "Renamed.mkv").exists()

    _drive(monkeypatch, ["n"])
    outcome = wizard.accept_or_revert(None, journal, session=session)
    assert outcome is wizard.GateOutcome.REVERTED

    assert _tree(tv) == before
    assert all(r["undone"] for r in list_runs(journal))


def test_the_word_list_step_offers_the_editor(monkeypatch, tmp_path):
    monkeypatch.setenv("MEDIAORG_PATTERNS", str(tmp_path / "words.json"))
    save_custom_patterns(["YIFY"])
    _drive(monkeypatch, ["y", "a", "RARBG", "q"])

    assert wizard.confirm_word_list() == ["YIFY", "RARBG"]


def test_the_word_list_step_can_be_walked_past(monkeypatch, tmp_path):
    monkeypatch.setenv("MEDIAORG_PATTERNS", str(tmp_path / "words.json"))
    save_custom_patterns(["YIFY"])
    _drive(monkeypatch, [""])

    assert wizard.confirm_word_list() == ["YIFY"]


@pytest.mark.parametrize("arg,total,expected", [
    ("3", 5, [2]),
    ("1,3", 5, [0, 2]),
    ("2-4", 5, [1, 2, 3]),
    ("1,3-5", 5, [0, 2, 3, 4]),
    ("0", 5, None),
    ("6", 5, None),
    ("4-2", 5, None),
    ("x", 5, None),
    ("", 5, None),
])
def test_item_number_parsing(arg, total, expected):
    assert wizard._parse_item_numbers(arg, total) == expected


def test_renaming_a_folder_does_not_drag_a_file_along(tmp_path, monkeypatch):
    """Path.stem of a dotted folder name must not match a neighbouring file.

    Both destinations below have the stem "Show", so a stem-only rule would
    rename the .mkv as though it were the folder's subtitle track.
    """
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    journal = tmp_path / "journal.jsonl"
    root = tmp_path / "Movies"
    (root / "old folder").mkdir(parents=True)
    (root / "stray.mkv").write_text("a file, not the folder's companion")

    plan = Plan(ops=[Op("move", root / "old folder", root / "Show.S01"),
                     Op("move", root / "stray.mkv", root / "Show.mkv")])
    _drive(monkeypatch, ["R", "e 1", "Renamed Folder", "Y", "y"])

    wizard.confirm_and_execute(plan, journal, label="renames", roots=[root])

    assert (root / "Renamed Folder").is_dir()
    assert (root / "Show.mkv").exists()             # followed its own plan
    assert not (root / "Renamed Folder.mkv").exists()


def test_renaming_a_folder_is_not_asked_about_extensions(tmp_path, monkeypatch):
    """A folder has no extension to protect, so it must not be questioned."""
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    journal = tmp_path / "journal.jsonl"
    root = tmp_path / "Movies"
    (root / "old folder").mkdir(parents=True)

    plan = Plan(ops=[Op("move", root / "old folder", root / "The Matrix (1999)")])
    # No answer is supplied for an extension question: if one is asked, the
    # iterator runs dry and the test fails loudly rather than silently passing.
    _drive(monkeypatch, ["R", "e 1", "The Matrix (1999) [1080p]", "Y", "y"])

    wizard.confirm_and_execute(plan, journal, label="renames", roots=[root])

    assert (root / "The Matrix (1999) [1080p]").is_dir()


def test_back_at_the_full_list_prompt_does_not_escape_the_gate(wired,
                                                               monkeypatch):
    """The viewer's prompt is not a step; 'back' there must not skip the gate."""
    tv, _ = wired
    before = _tree(tv)
    real_capped = wizard._print_capped
    # Cap at 1 so the "... and N more / show the full list?" branch is reached.
    monkeypatch.setattr(wizard, "_print_capped",
                        lambda lines, cap=1: real_capped(lines, 1))
    _drive(monkeypatch, ["Y", "back", "n"])   # back at "show full list?", then no

    wizard.run_organize(str(tv))

    assert _tree(tv) == before


def test_going_back_after_the_organize_phase_still_asks(wired, monkeypatch,
                                                        capsys):
    """[4]: 'back' mid-flow must not strand the organize phase unquestioned."""
    tv, journal = wired
    before = _tree(tv)
    monkeypatch.chdir(tv.parent)
    monkeypatch.setenv("MEDIAORG_CONFIG", str(tv.parent / "cfg.json"))
    monkeypatch.setenv("MEDIAORG_PATTERNS", str(tv.parent / "words.json"))
    monkeypatch.setattr(wizard, "browse_for_folder",
                        lambda *a, **k: str(tv) if "TV" in a[0] else None)
    monkeypatch.setattr(wizard, "run_scan",
                        lambda *a, **k: pytest.fail("should not get this far"))
    # menu [4] -> excel name -> organize review [Y] -> 'back' at the word list
    # -> the gate must still be reached -> 'n' reverts the organizing.
    _drive(monkeypatch, ["4", "media_library.xlsx", "Y", "back", "n", "0"])

    wizard.run_wizard()

    runs = list_runs(journal)
    assert runs, "the organize phase never ran, so this proves nothing"
    assert all(r["undone"] for r in runs)
    assert _tree(tv) == before


def test_the_step_counter_does_not_overcount(wired, monkeypatch, capsys):
    """"[3/4]" when there is no fourth step is a small lie, but it is a lie."""
    tv, _ = wired
    monkeypatch.chdir(tv.parent)
    monkeypatch.setenv("MEDIAORG_CONFIG", str(tv.parent / "cfg.json"))
    monkeypatch.setenv("MEDIAORG_PATTERNS", str(tv.parent / "words.json"))
    monkeypatch.setattr(wizard, "browse_for_folder",
                        lambda *a, **k: str(tv) if "TV" in a[0] else None)
    monkeypatch.setattr(wizard, "run_scan", lambda *a, **k: None)
    # [6] scan only: word list, then scan. Decline the word-list editor, exit.
    _drive(monkeypatch, ["6", "media_library.xlsx", "n", "0"])

    wizard.run_wizard()

    printed = capsys.readouterr().out
    steps = re.findall(r"\[(\d+)/(\d+)\]", printed)
    numbered = [(int(a), int(b)) for a, b in steps if int(b) < 10]
    assert numbered, "no step banners were printed"
    total = numbered[0][1]
    assert max(n for n, _ in numbered) == total, numbered


# --- Follow-ups from PR review -----------------------------------------------

def test_renaming_a_subtitle_does_not_drag_its_video(tmp_path, monkeypatch):
    """Companion coupling runs video -> sidecar, never the other way.

    Same-stem alone made it symmetric, so retitling a subtitle track renamed
    the film it belonged to.
    """
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    journal = tmp_path / "journal.jsonl"
    root = tmp_path / "Movies"
    root.mkdir(parents=True)
    (root / "raw.mkv").write_text("video")
    (root / "raw.srt").write_text("subs")

    plan = Plan(ops=[Op("move", root / "raw.mkv", root / "Episode.mkv"),
                     Op("move", root / "raw.srt", root / "Episode.srt")])
    # Item 2 is the subtitle. Renaming it must leave the video's plan alone.
    _drive(monkeypatch, ["R", "e 2", "English.srt", "Y", "y"])

    wizard.confirm_and_execute(plan, journal, label="renames", roots=[root])

    assert (root / "Episode.mkv").exists()      # unchanged by the .srt edit
    assert (root / "English.srt").exists()
    assert not (root / "English.mkv").exists()


def test_renaming_the_video_still_carries_the_subtitle(tmp_path, monkeypatch):
    """The direction that should propagate still does."""
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    journal = tmp_path / "journal.jsonl"
    root = tmp_path / "Movies"
    root.mkdir(parents=True)
    (root / "raw.mkv").write_text("video")
    (root / "raw.srt").write_text("subs")

    plan = Plan(ops=[Op("move", root / "raw.mkv", root / "Episode.mkv"),
                     Op("move", root / "raw.srt", root / "Episode.srt")])
    _drive(monkeypatch, ["R", "e 1", "Pilot.mkv", "Y", "y"])

    wizard.confirm_and_execute(plan, journal, label="renames", roots=[root])

    assert (root / "Pilot.mkv").exists()
    assert (root / "Pilot.srt").exists()


def test_excluding_a_move_drops_the_folder_it_was_going_into(tmp_path):
    """plan_loose_movies emits mkdir + move; excluding the move orphans it."""
    root = tmp_path / "Movies"
    root.mkdir()
    (root / "Film.mkv").write_text("v")
    folder = root / "Film"
    mkdir_op = Op("mkdir", None, folder)
    move_op = Op("move", root / "Film.mkv", folder / "Film.mkv")

    plan = check_collisions([mkdir_op], dropped=[move_op])

    assert plan.ops == []
    assert any("no longer needed" in reason for _, reason in plan.skipped)


def test_a_folder_still_wanted_by_another_move_is_kept(tmp_path):
    """Only fully-orphaned folders go; one surviving move is enough to keep it."""
    root = tmp_path / "Movies"
    root.mkdir()
    (root / "Film.mkv").write_text("v")
    (root / "Film.srt").write_text("s")
    folder = root / "Film"
    mkdir_op = Op("mkdir", None, folder)
    kept = Op("move", root / "Film.srt", folder / "Film.srt")
    gone = Op("move", root / "Film.mkv", folder / "Film.mkv")

    plan = check_collisions([mkdir_op, kept], dropped=[gone])

    assert mkdir_op in plan.ops


def test_a_planners_own_mkdir_is_never_second_guessed(tmp_path):
    """With nothing dropped, this rule must not fire at all."""
    root = tmp_path / "Movies"
    root.mkdir()
    folder = root / "Empty On Purpose"
    plan = check_collisions([Op("mkdir", None, folder)])
    assert [op.kind for op in plan.ops] == ["mkdir"]


def test_an_interrupt_before_the_session_gate_says_what_is_on_disk(wired,
                                                                   monkeypatch,
                                                                   capsys):
    """[4]: Ctrl-C mid-flow must not exit silently on an applied library."""
    tv, journal = wired
    monkeypatch.chdir(tv.parent)
    monkeypatch.setenv("MEDIAORG_CONFIG", str(tv.parent / "cfg.json"))
    monkeypatch.setenv("MEDIAORG_PATTERNS", str(tv.parent / "words.json"))
    monkeypatch.setattr(wizard, "browse_for_folder",
                        lambda *a, **k: str(tv) if "TV" in a[0] else None)
    answers = iter(["4", "media_library.xlsx", "Y"])

    def _input(*a):
        try:
            return next(answers)
        except StopIteration:
            raise KeyboardInterrupt          # interrupt at the word-list step
    monkeypatch.setattr("builtins.input", _input)
    monkeypatch.setattr(wizard, "_stdin_is_interactive", lambda: True)

    wizard.run_wizard()          # the outer handler catches it and returns

    runs = list_runs(journal)
    assert runs, "the organize phase never ran, so this proves nothing"
    assert not any(r["undone"] for r in runs)      # not reverted behind their back
    out = capsys.readouterr().out
    assert "never got the chance to accept" in out
    # Concrete run ids, not a bare --undo-session: that flag takes no argument
    # and resolves to whichever session is newest whenever it is finally run.
    for run in runs:
        assert f"--undo-run {run['id']}" in out


# --- Round 3: findings from the Kilo review ----------------------------------

def test_the_changes_are_listed_before_anything_is_asked(wired, monkeypatch,
                                                         capsys):
    """The whole point of the feature: nothing is applied unseen.

    The screen used to print only a count, so the default [Y] applied every
    change without the before/after list ever appearing.
    """
    tv, _ = wired
    _drive(monkeypatch, ["Q"])          # cancel at the very first prompt

    wizard.run_organize(str(tv))

    out = capsys.readouterr().out
    assert "BEFORE" in out and "AFTER" in out
    assert "My.Show.S01E01.mkv" in out


def test_bare_enter_pages_rather_than_applying(wired, monkeypatch, capsys):
    """Enter must never be the keystroke that commits a library-wide rename."""
    tv, journal = wired
    before = _tree(tv)
    # Two bare Enters on a single-page plan, then quit. If Enter applied, the
    # tree would change and the answers would run out.
    _drive(monkeypatch, ["", "", "Q"])

    wizard.run_organize(str(tv))

    assert _tree(tv) == before
    assert not journal.exists()
    assert "last page" in capsys.readouterr().out


def test_paging_never_splits_a_before_from_its_after(tmp_path, monkeypatch):
    """Folder ops print one line and moves print two, so a line-stride pager
    puts the page break in the middle of a change."""
    root = tmp_path / "Movies"
    root.mkdir()
    items = []
    for n in range(4):
        (root / f"F{n}.mkv").write_text("v")
        items.append(wizard._ReviewItem(Op("mkdir", None, root / f"F{n}"),
                                        Op("mkdir", None, root / f"F{n}")))
        mv = Op("move", root / f"F{n}.mkv", root / f"F{n}" / f"F{n}.mkv")
        items.append(wizard._ReviewItem(mv, mv))

    for page_start in (0, 3, 5):
        lines = wizard._review_lines(items, set(), page_start, page_start + 3)
        # Every BEFORE rendered in a slice is immediately followed by its AFTER.
        for i, line in enumerate(lines):
            if "BEFORE" in line:
                assert i + 1 < len(lines) and "AFTER" in lines[i + 1]


def test_an_unattended_run_applies_without_prompting(wired, monkeypatch,
                                                     capsys):
    """A cron job or Docker service has nobody to answer either question."""
    tv, journal = wired
    monkeypatch.setattr(wizard, "_stdin_is_interactive", lambda: False)
    monkeypatch.setattr("builtins.input",
                        lambda *a: pytest.fail("must not prompt"))

    wizard.run_organize(str(tv))

    assert (tv / "My Show" / "Season 1" / "My.Show.S01E01.mkv").exists()
    out = capsys.readouterr().out
    assert "not a terminal" in out
    assert "--undo-run" in out
    # The plan is still printed: the log is the only record anyone will read.
    assert "My.Show.S01E01.mkv" in out


def test_a_closed_stdin_at_the_gate_reverts_rather_than_tracebacks(wired,
                                                                   monkeypatch):
    tv, _ = wired
    before = _tree(tv)
    answers = iter(["Y"])

    def _input(*a):
        try:
            return next(answers)
        except StopIteration:
            raise EOFError
    monkeypatch.setattr("builtins.input", _input)
    monkeypatch.setattr(wizard, "_stdin_is_interactive", lambda: True)

    wizard.run_organize(str(tv))       # must not raise

    assert _tree(tv) == before


@pytest.mark.parametrize("token", ["b", "back"])
def test_back_in_the_review_is_not_a_command(wired, monkeypatch, token):
    """It used to raise BackNavigation and silently discard the whole review."""
    tv, _ = wired
    before = _tree(tv)
    _drive(monkeypatch, [token, "Q"])

    wizard.run_organize(str(tv))       # must not raise BackNavigation

    assert _tree(tv) == before


def test_a_file_can_be_named_back(tmp_path, monkeypatch):
    """'back' has to be a name you can type, not a control word."""
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    journal = tmp_path / "journal.jsonl"
    root = tmp_path / "Movies"
    root.mkdir()
    (root / "raw.mkv").write_text("v")
    plan = Plan(ops=[Op("move", root / "raw.mkv", root / "Episode.mkv")])
    _drive(monkeypatch, ["R", "e 1", "back.mkv", "Y", "y"])

    wizard.confirm_and_execute(plan, journal, label="renames", roots=[root])

    assert (root / "back.mkv").exists()


@pytest.mark.parametrize("bad", ["²", "①", "9" * 5000])
def test_unicode_digits_do_not_crash_the_review(bad):
    """isdigit() is True for these; int() raises. The ValueError escaped."""
    assert wizard._parse_item_numbers(bad, 10) is None


def test_a_typed_name_keeps_its_dotted_segments(tmp_path, monkeypatch):
    """Path.stem eats only the last dot, so restemming dropped '.1080p'."""
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    journal = tmp_path / "journal.jsonl"
    root = tmp_path / "Movies"
    root.mkdir()
    (root / "raw.mkv").write_text("v")
    plan = Plan(ops=[Op("move", root / "raw.mkv", root / "Episode.mkv")])
    # Type a dotted name with no extension, then decline to change the ext.
    _drive(monkeypatch, ["R", "e 1", "Show.S01E01.1080p", "n", "Y", "y"])

    wizard.confirm_and_execute(plan, journal, label="renames", roots=[root])

    assert (root / "Show.S01E01.1080p.mkv").exists()


def test_the_extension_guard_follows_the_plan_not_the_source(tmp_path,
                                                             monkeypatch):
    """extfix converts .ts -> .mp4; judging by the source undid the conversion."""
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    journal = tmp_path / "journal.jsonl"
    root = tmp_path / "Movies"
    root.mkdir()
    (root / "clip.ts").write_text("v")
    plan = Plan(ops=[Op("move", root / "clip.ts", root / "clip.mp4")])
    # Retype the name as .mp4. That matches the *plan*, so nothing is asked;
    # no answer is supplied for an extension question, so one would fail here.
    _drive(monkeypatch, ["R", "e 1", "Clip.mp4", "Y", "y"])

    wizard.confirm_and_execute(plan, journal, label="extension fixes",
                               roots=[root])

    assert (root / "Clip.mp4").exists()
    assert not (root / "Clip.ts").exists()


def test_a_language_tagged_subtitle_still_follows_its_video(tmp_path,
                                                            monkeypatch):
    """_companion_ops preserves '.en'; exact-stem matching never saw those."""
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    journal = tmp_path / "journal.jsonl"
    root = tmp_path / "TV"
    root.mkdir()
    for name in ("raw.mkv", "raw.en.srt", "raw.fr.srt"):
        (root / name).write_text(name)
    plan = Plan(ops=[
        Op("move", root / "raw.mkv", root / "Episode.mkv"),
        Op("move", root / "raw.en.srt", root / "Episode.en.srt"),
        Op("move", root / "raw.fr.srt", root / "Episode.fr.srt"),
    ])
    _drive(monkeypatch, ["R", "e 1", "Pilot.mkv", "Y", "y"])

    wizard.confirm_and_execute(plan, journal, label="renames", roots=[root])

    assert (root / "Pilot.mkv").exists()
    # Both tails survive: collapsing them onto one name would make
    # check_collisions drop both subtitle tracks as a duplicate target.
    assert (root / "Pilot.en.srt").read_text() == "raw.en.srt"
    assert (root / "Pilot.fr.srt").read_text() == "raw.fr.srt"


def test_renaming_a_video_does_not_drag_a_numbered_sibling(tmp_path,
                                                           monkeypatch):
    """A prefix is only a sidecar match on a non-alphanumeric boundary.

    "Episode10.en" starts with "Episode1", so a bare startswith renamed episode
    ten's subtitle on top of episode one's - the bystander is what makes this
    worse than an orphan, since it moves a file the user never selected.
    """
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    journal = tmp_path / "journal.jsonl"
    root = tmp_path / "TV"
    root.mkdir()
    for name in ("raw1.mkv", "raw1.en.srt", "raw10.en.srt"):
        (root / name).write_text(name)
    plan = Plan(ops=[
        Op("move", root / "raw1.mkv", root / "Episode1.mkv"),
        Op("move", root / "raw1.en.srt", root / "Episode1.en.srt"),
        Op("move", root / "raw10.en.srt", root / "Episode10.en.srt"),
    ])
    _drive(monkeypatch, ["R", "e 1", "Pilot.mkv", "Y", "y"])

    wizard.confirm_and_execute(plan, journal, label="renames", roots=[root])

    assert (root / "Pilot.mkv").exists()
    assert (root / "Pilot.en.srt").read_text() == "raw1.en.srt"
    # Episode 10's subtitle keeps its own planned name and its own content.
    assert (root / "Episode10.en.srt").read_text() == "raw10.en.srt"
    assert not (root / "Pilot0.en.srt").exists()


def test_a_longer_title_in_the_plan_keeps_its_own_sidecar(tmp_path,
                                                          monkeypatch):
    """The review screen settles sidecar ownership the same way the planner does."""
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    journal = tmp_path / "journal.jsonl"
    root = tmp_path / "Movies"
    root.mkdir()
    for name in ("a.mkv", "b.mkv", "b.en.srt"):
        (root / name).write_text(name)
    plan = Plan(ops=[
        Op("move", root / "a.mkv", root / "Film.mkv"),
        Op("move", root / "b.mkv", root / "Film.2010.1080p.mkv"),
        Op("move", root / "b.en.srt", root / "Film.2010.1080p.en.srt"),
    ])
    # Renaming item 1 ("Film") must not claim the subtitle planned for the
    # longer "Film.2010.1080p", even though ".2010.1080p.en" is a valid tail.
    _drive(monkeypatch, ["R", "e 1", "Pilot.mkv", "Y", "y"])

    wizard.confirm_and_execute(plan, journal, label="renames", roots=[root])

    assert (root / "Pilot.mkv").exists()
    assert (root / "Film.2010.1080p.en.srt").read_text() == "b.en.srt"
    assert not (root / "Pilot.2010.1080p.en.srt").exists()


def test_b_at_the_extension_prompt_keeps_the_review(tmp_path, monkeypatch):
    """'b' cancels the one rename, it does not discard every edit made so far.

    prompt_input raises BackNavigation, which used to unwind past the review
    all the way to the menu loop's `except BackNavigation: pass` - silently.
    """
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    journal = tmp_path / "journal.jsonl"
    root = tmp_path / "Movies"
    root.mkdir()
    for name in ("one.mkv", "two.mkv"):
        (root / name).write_text(name)
    # Names differ by more than case: on macOS and Windows a case-only rename
    # is invisible to exists(), so "Two.mkv" would test nothing there.
    plan = Plan(ops=[Op("move", root / "one.mkv", root / "Alpha.mkv"),
                     Op("move", root / "two.mkv", root / "Beta.mkv")])
    # Exclude item 2, then type a name with no extension and answer 'b' to the
    # extension question. The exclusion must survive and [Y] must still apply.
    _drive(monkeypatch, ["R", "x 2", "e 1", "Pilot", "b", "Y", "y"])

    wizard.confirm_and_execute(plan, journal, label="renames", roots=[root])

    assert (root / "Pilot.mkv").exists()     # extension kept, not restemmed
    assert (root / "two.mkv").exists()       # the exclusion survived 'b'
    assert not (root / "Beta.mkv").exists()


def test_b_at_the_sanitize_prompt_cancels_only_that_rename(tmp_path,
                                                           monkeypatch):
    """Same for the other ask_yes_no _edit_destination can reach."""
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    journal = tmp_path / "journal.jsonl"
    root = tmp_path / "Movies"
    root.mkdir()
    (root / "one.mkv").write_text("one")
    # Not a case-only rename: those are invisible to exists() on macOS.
    plan = Plan(ops=[Op("move", root / "one.mkv", root / "Planned.mkv")])
    # A colon does not survive sanitize(), so this asks "Use '...' instead?".
    _drive(monkeypatch, ["R", "e 1", "Pi:lot.mkv", "b", "Y", "y"])

    wizard.confirm_and_execute(plan, journal, label="renames", roots=[root])

    assert (root / "Planned.mkv").exists()   # the planned name, rename cancelled
    assert not (root / "one.mkv").exists()   # the review still applied


def test_ctrl_d_in_the_review_aborts_without_applying(tmp_path, monkeypatch):
    """A closed stdin is a "no", not a traceback - and must not spin the loop."""
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    journal = tmp_path / "journal.jsonl"
    root = tmp_path / "TV"
    root.mkdir()
    (root / "raw.mkv").write_text("video")
    plan = Plan(ops=[Op("move", root / "raw.mkv", root / "Episode.mkv")])

    def _eof(*a):
        raise EOFError
    monkeypatch.setattr("builtins.input", _eof)
    monkeypatch.setattr(wizard, "_stdin_is_interactive", lambda: True)

    entries = wizard.confirm_and_execute(plan, journal, label="renames",
                                         roots=[root])

    assert entries == []
    assert (root / "raw.mkv").exists()          # untouched
    assert not (root / "Episode.mkv").exists()


def test_the_abort_message_does_not_claim_the_disk_is_untouched(tmp_path,
                                                                monkeypatch,
                                                                capsys):
    """review_changes also runs mid-session, after organize is already applied.

    "Nothing has been changed" would be a lie there, and the session gate a
    moment later contradicts it by name.
    """
    root = tmp_path / "TV"
    root.mkdir()
    (root / "raw.mkv").write_text("v")
    plan = Plan(ops=[Op("move", root / "raw.mkv", root / "Episode.mkv")])

    def _eof(*a):
        raise EOFError
    monkeypatch.setattr("builtins.input", _eof)

    assert wizard.review_changes(plan, "renames") is None
    out = capsys.readouterr().out
    assert "None of these changes were applied" in out
    assert "Nothing has been changed" not in out


def test_ctrl_d_at_the_go_to_page_prompt_aborts(tmp_path, monkeypatch):
    """The sub-prompts are covered by the same handler, not just the main one."""
    plan = Plan(ops=[Op("move", tmp_path / f"a{i}", tmp_path / f"b{i}")
                     for i in range(25)])       # 3 pages, so [G] is offered
    answers = iter(["G"])            # then the page prompt hits EOF

    def _maybe_eof(*a):
        try:
            return next(answers)
        except StopIteration:
            raise EOFError
    monkeypatch.setattr("builtins.input", _maybe_eof)

    assert wizard.review_changes(plan, "renames") is None


def test_ctrl_d_at_the_new_name_prompt_aborts(tmp_path, monkeypatch):
    """_ask_new_name reads with a bare input() too."""
    root = tmp_path / "TV"
    root.mkdir()
    (root / "raw.mkv").write_text("v")
    plan = Plan(ops=[Op("move", root / "raw.mkv", root / "Episode.mkv")])
    answers = iter(["R", "e 1"])     # then the name prompt hits EOF

    def _maybe_eof(*a):
        try:
            return next(answers)
        except StopIteration:
            raise EOFError
    monkeypatch.setattr("builtins.input", _maybe_eof)

    assert wizard.review_changes(plan, "renames") is None


def test_the_cross_device_tag_survives_an_edit(tmp_path, monkeypatch):
    """Op is frozen and compared by value, so an edit dropped it from the set."""
    root = tmp_path / "Movies"
    root.mkdir()
    (root / "raw.mkv").write_text("v")
    op = Op("move", root / "raw.mkv", root / "Episode.mkv")
    items = [wizard._ReviewItem(op, op)]
    _drive(monkeypatch, ["Pilot.mkv"])

    wizard._edit_destination(items, 0)

    assert items[0].op.dst.name == "Pilot.mkv"
    # Keyed on .original, so the tag is still found after the replace().
    assert "COPIED across drives" in "".join(
        wizard._review_lines(items, {op}))


def test_a_skipped_move_also_drops_its_folder(tmp_path):
    """The sibling rmdir rule unions plan.skipped; this one must too."""
    root = tmp_path / "Movies"
    root.mkdir()
    folder = root / "Film"
    # The move's source does not exist, so check_collisions skips it itself.
    ops = [Op("mkdir", None, folder),
           Op("move", root / "gone.mkv", folder / "gone.mkv")]

    plan = check_collisions(ops)

    assert plan.ops == []
    assert any("no longer needed" in reason for _, reason in plan.skipped)


def test_a_failed_revert_still_records_what_is_on_disk(wired, monkeypatch,
                                                       capsys):
    """The library is still renamed, so the audit trail must not be blank."""
    tv, journal = wired
    _drive(monkeypatch, ["Y", "n"])
    monkeypatch.setattr(wizard, "undo_run",
                        lambda *a, **k: (None, "run is newer, use --force"))

    entries = wizard.confirm_and_execute(
        Plan(ops=[Op("move", tv / "My Show" / "My.Show.S01E01.mkv",
                     tv / "My Show" / "Renamed.mkv")]),
        journal, label="renames", roots=[tv])

    assert entries, "a failed revert must still report what was applied"
    out = capsys.readouterr().out
    assert "--force" in out
    assert "still in their new locations" in out


def test_rejecting_a_session_warns_the_spreadsheet_is_stale(wired, monkeypatch,
                                                            capsys, tmp_path):
    """organize -> scan -> revert leaves the sheet describing a dead layout."""
    tv, journal = wired
    xlsx = tmp_path / "lib.xlsx"
    xlsx.write_text("pretend spreadsheet")
    session = "cafef00dbeef"

    _drive(monkeypatch, ["Y"])
    wizard.run_organize(str(tv), session=session, accept_gate=False)

    _drive(monkeypatch, ["n"])
    outcome = wizard.accept_or_revert(None, journal, session=session)
    assert outcome is wizard.GateOutcome.REVERTED
    wizard._warn_stale_spreadsheet(xlsx)

    out = capsys.readouterr().out
    assert "no longer exists" in out and "scan again" in out


def test_the_undo_hint_names_real_run_ids(wired, monkeypatch, capsys):
    """--undo-session takes no id and resolves to whichever is newest."""
    tv, journal = wired
    _drive(monkeypatch, ["Y", "y"])

    wizard.run_organize(str(tv))

    run_id = list_runs(journal)[-1]["id"]
    assert f"--undo-run {run_id}" in capsys.readouterr().out


def test_the_planners_skip_reasons_are_shown(tmp_path, monkeypatch, capsys):
    """A bare count said something was dropped without saying what or why."""
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    journal = tmp_path / "journal.jsonl"
    root = tmp_path / "Movies"
    root.mkdir()
    (root / "a.mkv").write_text("a")
    (root / "taken.mkv").write_text("occupied")
    plan = check_collisions([Op("move", root / "a.mkv", root / "taken.mkv")])
    plan.ops.append(Op("move", root / "a.mkv", root / "b.mkv"))
    _drive(monkeypatch, ["Q"])

    wizard.confirm_and_execute(plan, journal, label="renames", roots=[root])

    out = capsys.readouterr().out
    assert "target already exists" in out


def test_a_scripted_rename_still_writes_the_changes_sheet(tmp_path, monkeypatch):
    """Regression guard: `--action rename` passes accept_gate=False, and
    gating the log on that stopped the sheet being written at all."""
    import pandas as pd

    from mediaorg import excel, scan
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    monkeypatch.setenv("MEDIAORG_JOURNAL", str(tmp_path / "journal.jsonl"))
    monkeypatch.setenv("MEDIAORG_CONFIG", str(tmp_path / "cfg.json"))
    monkeypatch.setenv("MEDIAORG_PATTERNS", str(tmp_path / "words.json"))
    movies = tmp_path / "Movies"
    (movies / "the.matrix.1999.1080p").mkdir(parents=True)
    (movies / "the.matrix.1999.1080p" / "the.matrix.1999.1080p.mkv").write_text("v")
    xlsx = tmp_path / "lib.xlsx"
    excel.write_library(xlsx, scan.scan_movies(movies), [], str(movies), None)

    # The CLI's own arguments: no --review, so accept_gate is False.
    _drive(monkeypatch, ["n", "n", "1", "Y"])
    entries = wizard.run_rename(str(movies), None, xlsx, accept_gate=False)

    assert entries, "nothing was renamed, so this proves nothing"
    logged = pd.read_excel(xlsx, sheet_name='Changes')
    assert len(logged) == len(entries)


def test_a_multi_phase_rename_does_not_log_twice(tmp_path, monkeypatch):
    """[4] logs after its own session gate, so run_rename must not also log."""
    import pandas as pd

    from mediaorg import excel, scan
    monkeypatch.setattr("mediaorg.execute.RETRY_DELAYS", ())
    monkeypatch.setenv("MEDIAORG_JOURNAL", str(tmp_path / "journal.jsonl"))
    monkeypatch.setenv("MEDIAORG_CONFIG", str(tmp_path / "cfg.json"))
    monkeypatch.setenv("MEDIAORG_PATTERNS", str(tmp_path / "words.json"))
    movies = tmp_path / "Movies"
    (movies / "the.matrix.1999.1080p").mkdir(parents=True)
    (movies / "the.matrix.1999.1080p" / "the.matrix.1999.1080p.mkv").write_text("v")
    xlsx = tmp_path / "lib.xlsx"
    excel.write_library(xlsx, scan.scan_movies(movies), [], str(movies), None)

    _drive(monkeypatch, ["n", "n", "1", "Y"])
    entries = wizard.run_rename(str(movies), None, xlsx, session="s" * 12,
                                accept_gate=False, log=False)
    wizard.log_changes(xlsx, entries)

    logged = pd.read_excel(xlsx, sheet_name='Changes')
    assert len(logged) == len(entries)
