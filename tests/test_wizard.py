"""Tests for the wizard-level features: inventory, the custom word list, and
the LLM configuration. These are the parts a user reaches from the menu, and
none of them had any coverage before."""

import json

import pytest

from mediaorg import llm, wizard
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
    """Feed the wizard a fixed sequence of prompt answers."""
    supplied = iter(answers)
    monkeypatch.setattr("builtins.input", lambda *a: next(supplied))


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


def test_an_env_key_is_never_written_to_disk(tmp_path, monkeypatch):
    target = tmp_path / "llm.json"
    monkeypatch.setenv("MEDIAORG_LLM_CONFIG", str(target))
    monkeypatch.setenv("GEMINI_API_KEY", "secret")
    llm.load_llm_config()
    assert not target.exists()


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
