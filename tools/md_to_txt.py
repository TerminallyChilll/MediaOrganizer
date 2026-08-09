#!/usr/bin/env python3
"""Generate the plain-text copy of each Markdown document at the repo root.

Why this exists: the README *is* the install guide, and the people who need
it most are the ones who downloaded a ZIP, do not know what a .md file is,
and open it in Notepad -- where the markdown syntax is noise rather than
formatting. GitHub renders README.md as the repository landing page, so the
.md cannot simply be renamed; this produces the .txt that ships beside it.

Two files saying the same thing drift, so this is a generator and not a
one-time conversion: ``--check`` re-renders and compares, and the test suite
calls it, so editing README.md without regenerating fails CI.

    python tools/md_to_txt.py            # rewrite every .txt
    python tools/md_to_txt.py --check    # verify they are current (CI)

The output is deliberately plain: no markdown left over, no line past 78
columns, and typographic characters (em dash, ellipsis, arrow) transliterated
to their ASCII spellings, which is both the plain-text convention and the
safest thing to hand an unknown editor.
"""

from __future__ import annotations

import argparse
import re
import sys
import textwrap
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent

#: Hard wrap. 78 leaves a couple of columns of slack in an 80-column console.
WIDTH = 78

#: Typography that carries no meaning a plain-text reader would miss.
TRANSLITERATIONS = {
    "—": "--",    # em dash
    "–": "-",     # en dash
    "…": "...",   # ellipsis
    "→": "->",    # rightwards arrow
    "‘": "'", "’": "'",
    "“": '"', "”": '"',
    " ": " ",     # non-breaking space
}

#: Underline character per heading level. Levels below these get none.
HEADING_RULES = {1: "=", 2: "-", 3: "~", 4: "."}


def _wrap(text: str, width: int = WIDTH, break_long: bool = False,
          **kwargs) -> list[str]:
    """textwrap with the two defaults that are wrong for this document.

    ``break_on_hyphens`` would split ``--dry-run`` across two lines, which in
    a document that is mostly command-line flags reads as two flags. And a
    long URL is worth overflowing the margin for: a wrapped one cannot be
    copied, so ``break_long_words`` is off outside of tables, where it stays
    on because a column that overflows takes the whole table's alignment
    with it.
    """
    return textwrap.wrap(text, width, break_on_hyphens=False,
                         break_long_words=break_long, **kwargs)

BANNER = """\
This is the plain-text copy of {source}, for reading in Notepad or any other
text editor. It is generated: to change anything here, edit {source} and run

    python tools/md_to_txt.py
"""


# ── inline markup ────────────────────────────────────────────────────────

def _transliterate(text: str) -> str:
    for char, ascii_form in TRANSLITERATIONS.items():
        text = text.replace(char, ascii_form)
    return text


def _inline(text: str) -> str:
    """Strip inline markup from a line of prose.

    Code spans are lifted out first and put back last. Stripping their
    backticks at the end instead would only protect a *lone* delimiter --
    ``*.xlsx`` survives because the emphasis patterns need a closing ``*`` --
    while a span with paired markup inside it, ``**x**`` or a link, would be
    eaten by the patterns below before its backticks were ever removed. A
    code span is content by definition, so nothing may rewrite it.

    Among the rest, order matters: links are consumed before emphasis, so a
    bold link keeps its URL.
    """
    text = _transliterate(text)

    spans: list[str] = []

    def stash(match: re.Match) -> str:
        spans.append(match.group(1))
        # NUL cannot occur in the source documents, so this cannot collide
        # with text the patterns below are meant to see.
        return f"\x00{len(spans) - 1}\x00"

    text = re.sub(r"`([^`]+)`", stash, text)
    # Images and links. An anchor link ("#updating") has no address worth
    # printing on paper, so only its text survives.
    text = re.sub(r"!\[([^\]]*)\]\(([^)]+)\)", r"\1 (\2)", text)
    text = re.sub(r"\[([^\]]+)\]\(#[^)]*\)", r"\1", text)
    # A link whose text is already its address ("[README.txt](README.txt)")
    # would otherwise print the name twice.
    text = re.sub(r"\[([^\]]+)\]\(([^)]+)\)",
                  lambda m: m.group(1) if m.group(1) == m.group(2)
                  else f"{m.group(1)} ({m.group(2)})", text)
    # Emphasis. The lookarounds keep intra-word underscores and a lone
    # asterisk (a glob, a footnote marker) from being read as a delimiter.
    text = re.sub(r"\*\*\*([^*\n]+)\*\*\*", r"\1", text)
    text = re.sub(r"\*\*([^*\n]+)\*\*", r"\1", text)
    text = re.sub(r"(?<!\w)__([^_\n]+)__(?!\w)", r"\1", text)
    text = re.sub(r"(?<!\w)\*([^*\n]+)\*(?!\w)", r"\1", text)
    text = re.sub(r"(?<!\w)_([^_\n]+)_(?!\w)", r"\1", text)

    for index, span in enumerate(spans):
        text = text.replace(f"\x00{index}\x00", span)
    return text


# ── block matchers ───────────────────────────────────────────────────────

_FENCE = re.compile(r"^\s*```")
_HEADING = re.compile(r"^(#{1,6})\s+(.*?)\s*#*$")
_HRULE = re.compile(r"^\s*([-*_])(\s*\1){2,}\s*$")
_BULLET = re.compile(r"^(\s*)([-*+])\s+(.*)$")
_NUMBERED = re.compile(r"^(\s*)(\d+)[.)]\s+(.*)$")
_TABLE_ROW = re.compile(r"^\s*\|")
_TABLE_SEP_CELL = re.compile(r"^:?-+:?$")
_QUOTE = re.compile(r"^\s*>\s?(.*)$")


def _is_block_start(line: str) -> bool:
    """Does this line begin something that is not flowing prose?"""
    return bool(_FENCE.match(line) or _HEADING.match(line)
                or _HRULE.match(line) or _TABLE_ROW.match(line)
                or _QUOTE.match(line))


# ── tables ───────────────────────────────────────────────────────────────

def _split_row(line: str) -> list[str]:
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def _is_separator(cells: list[str]) -> bool:
    return bool(cells) and all(_TABLE_SEP_CELL.match(c) for c in cells)


def _fit_columns(rows: list[list[str]]) -> list[int]:
    """Column widths that fit inside WIDTH, shrinking the widest first.

    A table that does not fit is normal here -- the exit-code table's second
    column is a sentence -- so the wide column wraps and the narrow ones stay
    readable, rather than every column being squeezed equally.
    """
    columns = max(len(r) for r in rows)
    widths = [max((len(r[i]) for r in rows if i < len(r)), default=1)
              for i in range(columns)]
    gaps = 2 * (columns - 1)
    floor = 8
    # Shrink the widest column one character at a time. Stop if every column
    # has hit the floor, so a pathologically wide table overflows rather
    # than looping forever.
    while sum(widths) + gaps > WIDTH and max(widths) > floor:
        widths[widths.index(max(widths))] -= 1
    return widths


def _render_table(rows: list[list[str]]) -> list[str]:
    body = [r for r in rows if not _is_separator(r)]
    if not body:
        return []
    widths = _fit_columns(body)
    header, *data = body

    def render(cells: list[str]) -> list[str]:
        wrapped = [_wrap(cells[i], widths[i], break_long=True) or [""]
                   if i < len(cells) else [""]
                   for i in range(len(widths))]
        height = max(len(w) for w in wrapped)
        lines = []
        for row in range(height):
            parts = [(w[row] if row < len(w) else "").ljust(widths[i])
                     for i, w in enumerate(wrapped)]
            lines.append("  ".join(parts).rstrip())
        return lines

    out = render(header)
    out.append("  ".join("-" * w for w in widths))
    for row in data:
        out.extend(render(row))
    return out


# ── prose and lists ──────────────────────────────────────────────────────

def _render_text_block(block: list[str], width: int = WIDTH) -> list[str]:
    """Wrap a run of prose, splitting it into list items where it has them.

    Items are joined before wrapping because the source wraps them at its own
    width; re-flowing to ours is the whole point, and honouring the original
    line breaks would leave a ragged half-width column.
    """
    items: list[tuple[str, str, list[str]]] = []   # (kind, marker, lines)
    for line in block:
        bullet = _BULLET.match(line)
        numbered = _NUMBERED.match(line)
        if bullet:
            indent, _, text = bullet.groups()
            items.append(("list", " " * (len(indent) // 2 * 2) + "* ", [text]))
        elif numbered:
            indent, number, text = numbered.groups()
            items.append(("list", " " * (len(indent) // 2 * 2) + f"{number}. ",
                          [text]))
        elif items and items[-1][0] == "list":
            items[-1][2].append(line.strip())
        elif items and items[-1][0] == "para":
            items[-1][2].append(line.strip())
        else:
            items.append(("para", "", [line.strip()]))

    out: list[str] = []
    for kind, marker, lines in items:
        text = _inline(" ".join(lines).strip())
        if not text:
            continue
        if kind == "para":
            out.extend(_wrap(text, width) or [""])
        else:
            indent = "  " + marker
            out.extend(_wrap(text, width,
                             initial_indent=indent,
                             subsequent_indent=" " * len(indent))
                       or [indent.rstrip()])
    return out


# ── document ─────────────────────────────────────────────────────────────

def convert(markdown: str, source_name: str) -> str:
    lines = markdown.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    out: list[str] = [BANNER.format(source=source_name).rstrip(),
                      "", "=" * WIDTH, ""]
    i = 0
    while i < len(lines):
        line = lines[i]

        if not line.strip():
            i += 1
            continue

        if _FENCE.match(line):
            i += 1
            code: list[str] = []
            while i < len(lines) and not _FENCE.match(lines[i]):
                code.append(lines[i])
                i += 1
            i += 1                       # closing fence (or end of file)
            # Verbatim apart from transliteration and the indent: a code
            # block is the one place a stray backtick or asterisk is content.
            while code and not code[-1].strip():
                code.pop()
            out.append("")
            out.extend("    " + _transliterate(c).rstrip() if c.strip() else ""
                       for c in code)
            out.append("")
            continue

        heading = _HEADING.match(line)
        if heading:
            level = len(heading.group(1))
            title = _inline(heading.group(2))
            if out and out[-1] != "":
                out.append("")
            out.append(title)
            rule = HEADING_RULES.get(level)
            if rule:
                out.append(rule * len(title))
            out.append("")
            i += 1
            continue

        if _HRULE.match(line):
            out.extend(["", "-" * WIDTH, ""])
            i += 1
            continue

        if _QUOTE.match(line):
            quoted = []
            while i < len(lines) and _QUOTE.match(lines[i]):
                quoted.append(_QUOTE.match(lines[i]).group(1))
                i += 1
            # Rendered at a narrower width and then indented, so the callout
            # still reads as set apart once the '>' markers are gone.
            body = _render_text_block([q for q in quoted if q.strip()],
                                      width=WIDTH - 4)
            out.append("")
            out.extend("  | " + entry if entry else "  |" for entry in body)
            out.append("")
            continue

        if _TABLE_ROW.match(line):
            rows = []
            while i < len(lines) and _TABLE_ROW.match(lines[i]):
                rows.append([_inline(c) for c in _split_row(lines[i])])
                i += 1
            out.append("")
            out.extend(_render_table(rows))
            out.append("")
            continue

        block = []
        while i < len(lines) and lines[i].strip() and not _is_block_start(lines[i]):
            block.append(lines[i])
            i += 1
        out.extend(_render_text_block(block))
        out.append("")

    # Collapse the runs of blank lines the block handlers leave behind.
    collapsed: list[str] = []
    for entry in out:
        if entry == "" and (not collapsed or collapsed[-1] == ""):
            continue
        collapsed.append(entry)
    return "\n".join(collapsed).rstrip() + "\n"


def sources(root: Path | None = None) -> list[Path]:
    """Every Markdown document at the repo root, in a stable order.

    Resolved at call time rather than bound as a default, so a test can point
    the whole generator at a scratch directory.
    """
    return sorted((root or REPO_ROOT).glob("*.md"))


def target_for(source: Path) -> Path:
    return source.with_suffix(".txt")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Generate plain-text copies of the Markdown docs.")
    parser.add_argument("--check", action="store_true",
                        help="verify the .txt files are current; change nothing")
    args = parser.parse_args(argv)

    stale = []
    for source in sources():
        target = target_for(source)
        expected = convert(source.read_text(encoding="utf-8"), source.name)
        current = (target.read_text(encoding="utf-8")
                   if target.exists() else None)
        if current == expected:
            continue
        if args.check:
            stale.append(target.name)
        else:
            target.write_text(expected, encoding="utf-8", newline="\n")
            print(f"wrote {target.name}")

    if stale:
        print("Out of date: " + ", ".join(stale), file=sys.stderr)
        print("Run: python tools/md_to_txt.py", file=sys.stderr)
        return 1
    if args.check:
        print("Plain-text copies are up to date.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
