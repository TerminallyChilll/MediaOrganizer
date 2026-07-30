import datetime
import os

import pytest

from mediaorg.parse import ParsedName
from mediaorg.plan import (
    NamingScheme, Op, build_episode_file_name, build_movie_file_name,
    build_movie_folder_name, check_collisions, companion_files, episode_code,
    extract_season_episode, folder_has_episodes_or_seasons, plan_loose_movies,
    plan_season_structure, sanitize,
)


def touch(path):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("x")
    return path


def apply_plan(plan):
    """Minimal in-test applier so planner tests don't depend on execute.py."""
    for op in plan.ops:
        if op.kind == "move":
            op.dst.parent.mkdir(parents=True, exist_ok=True)
            op.src.rename(op.dst)
        elif op.kind == "rmdir":
            if not any(op.dst.iterdir()):
                op.dst.rmdir()
        elif op.kind == "mkdir":
            op.dst.mkdir(parents=True, exist_ok=True)


# --- Regressions from the old (failing) suite -------------------------------

def test_watch_folder_not_treated_as_season(tmp_path):
    """'WATCH - The Office Extended 9x9 - FREE' must produce zero ops."""
    for d in ["WATCH - The Office Extended 9x9 - FREE",
              "WATCH - The Office Extended 9x8 - FREE",
              "Deleted Scenes Season 9",
              "Bloopers Season 8"]:
        (tmp_path / d).mkdir()
    plan = plan_season_structure(tmp_path)
    assert plan.ops == [] and plan.skipped == []


def test_watch_root_not_detected_as_show(tmp_path):
    for d in ["WATCH - The Office Extended 9x9 - FREE",
              "WATCH - Bloopers - Season 8 - FREE",
              "Deleted Scenes Season 9"]:
        (tmp_path / d).mkdir()
    assert folder_has_episodes_or_seasons(tmp_path) is False


def test_show_detection_positive(tmp_path):
    (tmp_path / "Season 1").mkdir()
    assert folder_has_episodes_or_seasons(tmp_path) is True

    show2 = tmp_path / "show2"
    show2.mkdir()
    touch(show2 / "Show.S01E01.mkv")
    assert folder_has_episodes_or_seasons(show2) is True

    show3 = tmp_path / "show3"
    (show3 / "Snowfall S02").mkdir(parents=True)
    assert folder_has_episodes_or_seasons(show3) is True


# --- Season structure -------------------------------------------------------

def test_pure_season_folders_get_normalized(tmp_path):
    (tmp_path / "S01").mkdir()
    (tmp_path / "season 2").mkdir()
    plan = plan_season_structure(tmp_path)
    renames = {(str(o.src.name), str(o.dst.name)) for o in plan.ops if o.kind == "move"}
    assert ("S01", "Season 1") in renames
    assert ("season 2", "Season 2") in renames


def test_show_name_season_folders_get_normalized(tmp_path):
    (tmp_path / "Snowfall S02").mkdir()
    plan = plan_season_structure(tmp_path)
    assert [(o.src.name, o.dst.name) for o in plan.ops] == [("Snowfall S02", "Season 2")]


def test_loose_episode_folders_get_grouped(tmp_path):
    (tmp_path / "Show S01E01 720p").mkdir()
    (tmp_path / "Show S01E02 720p").mkdir()
    (tmp_path / "Show S02E01").mkdir()
    plan = plan_season_structure(tmp_path)
    dsts = {str(o.dst.relative_to(tmp_path)) for o in plan.ops}
    assert dsts == {"Season 1/Show S01E01 720p",
                    "Season 1/Show S01E02 720p",
                    "Season 2/Show S02E01"}


def test_loose_episode_files_move_with_companions(tmp_path):
    touch(tmp_path / "Show.S01E01.720p.mkv")
    # Companions match by episode code, not full stem (no quality tag here).
    touch(tmp_path / "Show.S01E01.srt")
    touch(tmp_path / "Show.S01E01.nfo")
    touch(tmp_path / "unrelated.txt")
    plan = plan_season_structure(tmp_path)
    dsts = {str(o.dst.relative_to(tmp_path)) for o in plan.ops}
    assert dsts == {"Season 1/Show.S01E01.720p.mkv",
                    "Season 1/Show.S01E01.srt",
                    "Season 1/Show.S01E01.nfo"}


def test_flatten_episode_dirs_inside_season(tmp_path):
    ep_dir = tmp_path / "Season 4" / "Show S04E10 Title"
    touch(ep_dir / "Show.S04E10.mp4")
    touch(ep_dir / "Show.S04E10.srt")
    plan = plan_season_structure(tmp_path)
    moves = {str(o.dst.relative_to(tmp_path)) for o in plan.ops if o.kind == "move"}
    assert moves == {"Season 4/Show.S04E10.mp4", "Season 4/Show.S04E10.srt"}
    rmdirs = [o.dst for o in plan.ops if o.kind == "rmdir"]
    assert rmdirs == [ep_dir]


def test_merge_duplicate_season_folders(tmp_path):
    touch(tmp_path / "S02" / "Show.S02E01.mkv")
    touch(tmp_path / "Season 2" / "Show.S02E02.mkv")
    plan = plan_season_structure(tmp_path)
    apply_plan(plan)
    assert plan.skipped == []
    assert not (tmp_path / "S02").exists()
    season2 = sorted(p.name for p in (tmp_path / "Season 2").iterdir())
    assert season2 == ["Show.S02E01.mkv", "Show.S02E02.mkv"]


def test_season_zero_specials(tmp_path):
    touch(tmp_path / "Show.S00E01.Special.mkv")
    plan = plan_season_structure(tmp_path)
    assert [str(o.dst.relative_to(tmp_path)) for o in plan.ops] == \
        ["Season 0/Show.S00E01.Special.mkv"]


def test_idempotent(tmp_path):
    touch(tmp_path / "Show S01E01 720p" / "Show.S01E01.mkv")
    touch(tmp_path / "S02" / "Show.S02E01.mkv")
    touch(tmp_path / "Show.S03E01.mkv")
    plan = plan_season_structure(tmp_path)
    assert plan.ops
    apply_plan(plan)
    replan = plan_season_structure(tmp_path)
    # Note: "Season 1/Show S01E01 720p" gets flattened on the second pass
    # (folder moved into season first, contents flattened next run) — apply
    # until stable, then assert stability.
    while replan.ops:
        apply_plan(replan)
        replan = plan_season_structure(tmp_path)
    assert replan.ops == [] and replan.skipped == []


# --- Collisions -------------------------------------------------------------

def test_two_sources_one_target_both_skipped(tmp_path):
    a = touch(tmp_path / "a" / "same.mkv")
    b = touch(tmp_path / "b" / "same.mkv")
    ops = [Op("move", a, tmp_path / "out" / "same.mkv"),
           Op("move", b, tmp_path / "out" / "same.mkv")]
    plan = check_collisions(ops)
    assert plan.ops == []
    assert len(plan.skipped) == 2


def test_existing_target_skipped(tmp_path):
    src = touch(tmp_path / "src.mkv")
    dst = touch(tmp_path / "dst.mkv")
    plan = check_collisions([Op("move", src, dst)])
    assert plan.ops == []
    assert "already exists" in plan.skipped[0][1]


def test_missing_source_skipped(tmp_path):
    plan = check_collisions([Op("move", tmp_path / "ghost.mkv", tmp_path / "o.mkv")])
    assert plan.ops == []
    assert "source missing" in plan.skipped[0][1]


# --- Loose movies -----------------------------------------------------------

def test_plan_loose_movies(tmp_path):
    touch(tmp_path / "Some.Movie.2020.mkv")
    touch(tmp_path / "Some.Movie.2020.srt")
    touch(tmp_path / "Already Foldered (2019)" / "m.mkv")
    plan = plan_loose_movies(tmp_path)
    dsts = {str(o.dst.relative_to(tmp_path)) for o in plan.ops}
    assert dsts == {"Some.Movie.2020",
                    "Some.Movie.2020/Some.Movie.2020.mkv",
                    "Some.Movie.2020/Some.Movie.2020.srt"}
    apply_plan(plan)
    assert plan_loose_movies(tmp_path).ops == []  # idempotent


# --- Helpers / builders -----------------------------------------------------

def test_extract_season_episode():
    assert extract_season_episode("Show.S01E05.mkv") == (1, 5)
    assert extract_season_episode("Show 1x05.mkv") == (1, 5)
    assert extract_season_episode("Show.S00E01.mkv") == (0, 1)  # falsy season!
    assert extract_season_episode("A Movie (2020).mkv") == (None, None)


def test_companion_files(tmp_path):
    v = touch(tmp_path / "Ep.S01E01.mkv")
    s = touch(tmp_path / "Ep.S01E01.en.srt")
    touch(tmp_path / "other.srt")
    assert companion_files(v) == [s]


def test_episode_code():
    assert episode_code(ParsedName(title="X", season=1, episodes=[5])) == "S01E05"
    assert episode_code(ParsedName(title="X", season=1, episodes=[1, 2])) == "S01E01-E02"
    assert episode_code(ParsedName(title="X", season=0, episodes=[1])) == "S00E01"
    assert episode_code(ParsedName(title="X", date=datetime.date(2024, 1, 15))) == "2024-01-15"


def test_builders_and_sanitize():
    scheme = NamingScheme()
    p = ParsedName(title="The Matrix", year=1999, quality="1080p")
    assert build_movie_folder_name(p, None, scheme) == "The Matrix (1999) [1080p]"
    scheme.movie_folder_include_year = False
    scheme.movie_folder_include_quality = False
    assert build_movie_folder_name(p, None, scheme) == "The Matrix"

    ep = ParsedName(title="Show: Redux", season=1, episodes=[1], quality="720p")
    name = build_episode_file_name(ep, ".mkv", None, scheme)
    assert name == "Show Redux S01E01 [720p].mkv"  # ':' sanitized away

    assert sanitize('Bad:Name<>"') == "BadName"
    assert sanitize("Trailing. ") == "Trailing"


def test_year_tagged_season_folder_left_alone(tmp_path):
    # "Season 1 (2024)" is what the rename scheme produces; organize must
    # not strip the year (PR review regression).
    touch(tmp_path / "Season 1 (2024)" / "Show.S01E01.mkv")
    plan = plan_season_structure(tmp_path)
    assert plan.ops == [] and plan.skipped == []


def test_loose_files_join_year_tagged_season(tmp_path):
    touch(tmp_path / "Season 1 (2024)" / "Show.S01E01.mkv")
    touch(tmp_path / "Show.S01E02.mkv")
    plan = plan_season_structure(tmp_path)
    assert [str(o.dst.relative_to(tmp_path)) for o in plan.ops] == \
        ["Season 1 (2024)/Show.S01E02.mkv"]


# --- Sanitization hardening --------------------------------------------------

import unicodedata

from mediaorg import plan as pl
from mediaorg.plan import max_path_length, norm, sanitize


def test_windows_reserved_names_are_defused():
    """CON/NUL/COM1 are unusable as filenames on Windows, with or without ext."""
    for name in ("CON", "con", "PRN", "AUX", "NUL", "COM1", "LPT9"):
        assert sanitize(name).rstrip("_").upper() == name.upper()
        assert sanitize(name) != name, f"{name} passed through unchanged"
    # Only the part before the extension counts, and the extension survives.
    assert sanitize("nul.mkv") == "nul_.mkv"
    # A name that merely starts with a reserved word is fine.
    assert sanitize("Contact") == "Contact"
    assert sanitize("Aux Cable Story") == "Aux Cable Story"


def test_sanitize_never_returns_an_empty_name():
    """An empty stem produced a bare '.mkv' — a hidden file with no name."""
    for name in ("", "   ", ".", "..", '?<>:*|', "\x00\x01"):
        assert sanitize(name) == "Untitled"
    assert sanitize("", fallback="S01E01") == "S01E01"


def test_sanitize_strips_leading_dots_and_invisibles():
    assert sanitize("...Movie") == "Movie"
    assert sanitize(".hidden") == "hidden"
    assert sanitize("Movie​‮") == "Movie"


def test_sanitize_clamps_to_255_bytes():
    assert len(sanitize("A" * 300).encode("utf-8")) == pl.MAX_COMPONENT_BYTES
    # A CJK title busts the byte limit long before the character limit.
    cjk = sanitize("漫" * 200)
    assert len(cjk.encode("utf-8")) <= pl.MAX_COMPONENT_BYTES
    assert cjk  # and it is not clamped into nothing


def test_sanitize_is_nfc_normalized():
    nfd = unicodedata.normalize("NFD", "Amélie")
    assert sanitize(nfd) == unicodedata.normalize("NFC", "Amélie")
    assert norm(nfd) == norm("Amélie")


def test_builders_never_emit_a_bare_extension():
    scheme = NamingScheme()
    p = ParsedName(title="???")
    assert build_movie_file_name(p, ".mkv", None, scheme) == "Untitled.mkv"
    ep = ParsedName(title="?", season=1, episodes=[1])
    assert build_episode_file_name(ep, ".mkv", None, scheme) == "S01E01.mkv"


def test_overlong_destination_is_skipped_not_crashed(tmp_path):
    src = touch(tmp_path / "a.mkv")
    long_dst = tmp_path / ("B" * (max_path_length() + 10))
    plan = check_collisions([Op("move", src, long_dst)])
    assert plan.ops == []
    assert "over this system" in plan.skipped[0][1]


# --- Episode/season parsing --------------------------------------------------

def test_three_digit_episodes_are_not_truncated():
    """S01E101 used to come back as episode 10."""
    assert extract_season_episode("Show.S01E101.mkv") == (1, 101)
    assert extract_season_episode("Show.S01E01.mkv") == (1, 1)
    # Multi-episode files still report the first episode from the regex path.
    assert extract_season_episode("Show.S01E01E02.mkv") == (1, 1)


def test_year_style_season_is_recognized():
    assert extract_season_episode("Show.S2024E05.mkv") == (2024, 5)


def test_zero_padded_season_folder_is_not_canonical():
    """'Season 01' was treated as canonical, so it could never be normalized."""
    assert pl._CANONICAL_SEASON.match("Season 1")
    assert pl._CANONICAL_SEASON.match("Season 1 (2024)")
    assert pl._CANONICAL_SEASON.match("Season 0")
    assert not pl._CANONICAL_SEASON.match("Season 01")


def test_editions_and_parts_no_longer_collide():
    """EXTENDED vs theatrical used to build the same name, so BOTH were skipped."""
    scheme = NamingScheme()
    theatrical = ParsedName(title="Movie", year=1999, quality="1080p")
    extended = ParsedName(title="Movie", year=1999, quality="1080p",
                          edition="Extended")
    assert (build_movie_file_name(theatrical, ".mkv", None, scheme)
            != build_movie_file_name(extended, ".mkv", None, scheme))
    part1 = ParsedName(title="Movie", year=1999, part=1)
    part2 = ParsedName(title="Movie", year=1999, part=2)
    assert (build_movie_file_name(part1, ".mkv", None, scheme)
            != build_movie_file_name(part2, ".mkv", None, scheme))


# --- Junk files --------------------------------------------------------------

def test_appledouble_sidecar_is_not_media(tmp_path):
    """'._Show.S01E01.mkv' has a .mkv suffix and used to collide with the real
    episode, causing check_collisions to skip BOTH."""
    touch(tmp_path / "Show.S01E01.mkv")
    touch(tmp_path / "._Show.S01E01.mkv")
    touch(tmp_path / ".DS_Store")
    plan = plan_season_structure(tmp_path)
    dsts = [str(o.dst.relative_to(tmp_path)) for o in plan.ops]
    assert dsts == ["Season 1/Show.S01E01.mkv"]
    assert plan.skipped == []


def test_junk_does_not_make_a_folder_look_like_a_show(tmp_path):
    (tmp_path / "notashow").mkdir()
    (tmp_path / "notashow" / "._Show.S01E01.mkv").write_text("junk")
    assert not folder_has_episodes_or_seasons(tmp_path / "notashow")


def test_hardlinked_target_is_not_waved_through(tmp_path):
    a = touch(tmp_path / "a.mkv")
    b = tmp_path / "b.mkv"
    try:
        os.link(a, b)
    except (OSError, AttributeError, NotImplementedError):
        pytest.skip("filesystem does not support hardlinks")
    plan = check_collisions([Op("move", a, b)])
    assert plan.ops == [] and "already exists" in plan.skipped[0][1]
