"""The plain-text copies of the Markdown docs, and the generator behind them.

The first test is the one that matters: README.txt is a second copy of the
install guide, and a second copy that can drift is worse than no copy, because
the reader who cannot tell it is stale is exactly the reader it was added for.
"""

import re

import pytest

from tools import md_to_txt

pytestmark = pytest.mark.filterwarnings("ignore")


# ── the drift guard ──────────────────────────────────────────────────────

def test_committed_copies_are_current():
    """Regenerating must be a no-op. If this fails, run the generator.

    Note the CI trigger deliberately does *not* ignore '**.md' — with markdown
    excluded from the path filter, a README-only edit would skip the whole
    workflow and this test would never run on the change that broke it.
    """
    stale = [target.name for source in md_to_txt.sources()
             if (target := md_to_txt.target_for(source)).exists()
             and target.read_text(encoding="utf-8")
             != md_to_txt.convert(source.read_text(encoding="utf-8"),
                                  source.name)]
    assert not stale, (
        f"out of date: {', '.join(stale)} — run: python tools/md_to_txt.py")


def test_every_markdown_doc_has_a_plain_text_copy():
    missing = [s.name for s in md_to_txt.sources()
               if not md_to_txt.target_for(s).exists()]
    assert not missing, f"no .txt copy for: {', '.join(missing)}"


def test_check_mode_reports_a_stale_copy(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(md_to_txt, "REPO_ROOT", tmp_path)
    (tmp_path / "DOC.md").write_text("# Title\n\nBody.\n", encoding="utf-8")
    (tmp_path / "DOC.txt").write_text("something else\n", encoding="utf-8")

    assert md_to_txt.main(["--check"]) == 1
    assert "DOC.txt" in capsys.readouterr().err
    # --check must not have "helpfully" fixed it on the way past.
    assert (tmp_path / "DOC.txt").read_text(encoding="utf-8") == "something else\n"

    assert md_to_txt.main([]) == 0
    assert md_to_txt.main(["--check"]) == 0


# ── conversion ───────────────────────────────────────────────────────────

def convert(markdown: str) -> str:
    """Convert, minus the generated banner — which has prose and an indented
    command of its own and would otherwise answer half these assertions."""
    out = md_to_txt.convert(markdown, "DOC.md")
    _, separator, body = out.partition("=" * md_to_txt.WIDTH + "\n")
    assert separator, "banner separator missing"
    return body


def test_headings_are_underlined_not_hashed():
    out = convert("# Top\n\n## Second\n\nText.\n")
    assert "Top\n===" in out
    assert "Second\n------" in out
    assert not re.search(r"^#", out, re.MULTILINE)


def test_code_blocks_are_indented_and_left_verbatim():
    out = convert("Do this:\n\n```bash\ngit reset --hard   # **not** italic_x_\n```\n")
    assert "    git reset --hard   # **not** italic_x_" in out


def test_long_command_lines_are_not_wrapped():
    """Wrapping a command turns a thing you can paste into two broken ones."""
    command = "python run.py --action scan --movies /media/Movies --tv /media/TV --output lib.xlsx"
    out = convert(f"```bash\n{command}\n```\n")
    assert f"    {command}" in out


def test_links_keep_their_address_and_anchors_drop_theirs():
    out = convert("See [guessit](https://example.com/g) and [Updating](#updating).\n")
    assert "guessit (https://example.com/g)" in out
    assert "Updating." in out
    assert "#updating" not in out


def test_a_self_titled_link_is_not_printed_twice():
    assert "README.txt (README.txt)" not in convert("See [README.txt](README.txt).")
    assert "See README.txt." in convert("See [README.txt](README.txt).")


def test_emphasis_and_code_spans_are_stripped():
    out = convert("A **bold** and *italic* `code` and _under_ word.\n")
    assert "A bold and italic code and under word." in out
    for marker in ("**", "`"):
        assert marker not in out


def test_a_lone_asterisk_inside_code_survives():
    """`*.xlsx` and `[U]` are content, not markup — the backticks protect them
    from the emphasis and link patterns, so they must be stripped last."""
    out = convert("Press `[U]`, then open `*.xlsx`.\n")
    assert "Press [U], then open *.xlsx." in out


def test_bullets_are_reflowed_with_a_hanging_indent():
    out = convert("- " + "word " * 40 + "\n")
    lines = [ln for ln in out.splitlines() if ln.startswith("  ")]
    assert lines[0].startswith("  * word")
    assert lines[1].startswith("    word")      # continuation, no marker


def test_blockquotes_are_marked_not_left_with_their_angle_brackets():
    out = convert("> A callout **worth** noticing.\n")
    assert "  | A callout worth noticing." in out
    assert ">" not in out


def test_tables_become_aligned_columns():
    out = convert("| Code | Meaning |\n|---|---|\n| 0 | Fine |\n| 1 | Broken |\n")
    assert "Code  Meaning" in out
    assert "----  -------" in out
    assert re.search(r"^0     Fine$", out, re.MULTILINE)
    assert "|" not in out


def test_a_wide_table_wraps_inside_its_column():
    out = convert("| Code | Meaning |\n|---|---|\n| 0 | " + "word " * 40 + "|\n")
    for line in out.splitlines():
        assert len(line) <= md_to_txt.WIDTH


def test_flags_are_never_broken_across_lines():
    """break_on_hyphens would split '--dry-run', which in a page of flags
    reads as two different flags rather than one wrapped one."""
    out = convert("Text text text " * 4 + "and then --dry-run ends it.\n")
    assert "--dry-run" in out
    assert not re.search(r"-\n", out)


def test_a_url_overflows_rather_than_being_split():
    url = "https://github.com/TerminallyChilll/MediaOrganizer/blob/main/" + "x" * 40
    out = convert(f"Read more at {url} today.\n")
    assert url in out


def test_typography_is_transliterated_to_ascii():
    out = convert("A — dash, an … ellipsis, an → arrow.\n")
    assert "A -- dash, an ... ellipsis, an -> arrow." in out


def test_the_banner_names_the_file_to_edit_instead():
    out = md_to_txt.convert("# T\n", "README.md")
    assert "README.md" in out.splitlines()[0]
    assert "python tools/md_to_txt.py" in out


def test_blank_lines_never_pile_up():
    out = convert("# T\n\n\n\n\nBody.\n\n\n\n## S\n\n\n- a\n")
    assert "\n\n\n" not in out


def test_output_ends_with_exactly_one_newline():
    assert convert("# T\n\nBody.\n\n\n").endswith("Body.\n")


# ── the real documents ───────────────────────────────────────────────────

@pytest.mark.parametrize("source", md_to_txt.sources(),
                         ids=lambda p: p.name)
def test_no_markdown_syntax_survives(source):
    text = md_to_txt.target_for(source).read_text(encoding="utf-8")
    assert not re.search(r"^#{1,6} ", text, re.MULTILINE), "heading hashes"
    assert not re.search(r"\*\*", text), "bold markers"
    assert not re.search(r"\]\(", text), "link syntax"
    assert "```" not in text, "code fences"
    assert not re.search(r"^\s*>", text, re.MULTILINE), "blockquote markers"


@pytest.mark.parametrize("source", md_to_txt.sources(),
                         ids=lambda p: p.name)
def test_prose_fits_the_margin(source):
    """Code blocks are exempt: they are indented four spaces and copied
    verbatim, because a wrapped command is a broken command."""
    for number, line in enumerate(
            md_to_txt.target_for(source).read_text(encoding="utf-8").splitlines(), 1):
        if line.startswith("    "):
            continue
        assert len(line) <= md_to_txt.WIDTH, f"{source.stem}.txt line {number}"
