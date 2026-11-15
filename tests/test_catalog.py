"""Tests for the offline Movie & TV Show catalog (mediaorg.catalog)."""

import gzip
from pathlib import Path

import pandas as pd
import pytest

from mediaorg import catalog, doctor, wizard
from mediaorg.parse import parse_name

SAMPLE_RATINGS = """\
tconst\taverageRating\tnumVotes
tt0133093\t8.7\t2050000
tt1375666\t8.8\t2500000
tt8579674\t8.2\t680000
tt1038988\t7.4\t195000
tt0910970\t8.4\t1200000
tt0114369\t8.6\t1800000
tt1396484\t7.3\t600000
tt0903747\t9.5\t2100000
tt0472954\t8.8\t250000
tt0386676\t9.0\t700000
tt0211915\t8.3\t790000
tt9990001\t5.0\t10
tt9990002\t6.5\t5000
tt9990003\t7.0\t5000
tt0133094\t6.0\t500
"""

SAMPLE_BASICS = """\
tconst\ttitleType\tprimaryTitle\toriginalTitle\tisAdult\tstartYear\tendYear\truntimeMinutes\tgenres
tt0133093\tmovie\tThe Matrix\tThe Matrix\t0\t1999\t\\N\t136\tAction,Sci-Fi
tt0133094\tmovie\tThe Matrix\tThe Matrix\t0\t1999\t\\N\t90\tDocumentary
tt1375666\tmovie\tInception\tInception\t0\t2010\t\\N\t148\tAction,Sci-Fi
tt8579674\tmovie\t1917\t1917\t0\t2019\t\\N\t119\tDrama,War
tt1038988\tmovie\tREC\t[REC]\t0\t2007\t\\N\t78\tHorror
tt0910970\tmovie\tWALL·E\tWALL·E\t0\t2008\t\\N\t98\tAnimation
tt0114369\tmovie\tSe7en\tSe7en\t0\t1995\t\\N\t127\tCrime,Drama
tt1396484\tmovie\tIt\tIt\t0\t2017\t\\N\t135\tHorror
tt0903747\ttvSeries\tBreaking Bad\tBreaking Bad\t0\t2008\t2013\t49\tCrime,Drama
tt0472954\ttvSeries\tIt's Always Sunny in Philadelphia\tIt's Always Sunny in Philadelphia\t0\t2005\t\\N\t22\tComedy
tt0386676\ttvSeries\tThe Office\tThe Office\t0\t2005\t2013\t22\tComedy
tt0211915\tmovie\tAmélie\tLe Fabuleux Destin d'Amélie Poulain\t0\t2001\t\\N\t122\tComedy,Romance
tt9990001\tmovie\tToo Few Votes\tToo Few Votes\t0\t2020\t\\N\t90\tDrama
tt9990002\tmovie\tAdult Film\tAdult Film\t1\t2020\t\\N\t90\tAdult
tt9990003\tshort\tSome Short\tSome Short\t0\t2020\t\\N\t15\tShort
"""


@pytest.fixture
def sample_catalog(tmp_path, monkeypatch) -> Path:
    """Build a sample SQLite catalog from gzipped TSV fixtures."""
    ratings_gz = tmp_path / "title.ratings.tsv.gz"
    basics_gz = tmp_path / "title.basics.tsv.gz"
    with gzip.open(ratings_gz, "wt", encoding="utf-8") as fh:
        fh.write(SAMPLE_RATINGS)
    with gzip.open(basics_gz, "wt", encoding="utf-8") as fh:
        fh.write(SAMPLE_BASICS)

    db_path = tmp_path / ".media_catalog.db"
    monkeypatch.setenv("MEDIAORG_CATALOG", str(db_path))
    catalog.build_catalog(
        db_path,
        basics_source=basics_gz,
        ratings_source=ratings_gz,
        min_movie_votes=100,
        min_tv_votes=50,
    )
    return db_path


def test_normalize_title_handles_punctuation_accents_and_apostrophes():
    assert catalog.normalize_title("WALL·E") == "wall e"
    assert catalog.normalize_title("WALL-E") == "wall e"
    assert catalog.normalize_title("[REC]") == "rec"
    assert catalog.normalize_title("It's Always Sunny in Philadelphia") == (
        "its always sunny in philadelphia"
    )
    assert catalog.normalize_title("Its Always Sunny in Philadelphia") == (
        "its always sunny in philadelphia"
    )
    assert catalog.normalize_title("Amélie") == "amelie"
    assert catalog.normalize_title("Tom & Jerry") == "tom and jerry"
    assert catalog.normalize_title("Æon Flux") == "aeon flux"
    assert catalog.normalize_title("千と千尋の神隠し") == "千と千尋の神隠し"


def test_build_catalog_filters_and_deduplicates(sample_catalog):
    stats = catalog.catalog_stats(sample_catalog)
    assert stats is not None
    assert stats["movie_count"] == 9  # 8 unique + 1 duplicate norm_title/year merged
    assert stats["tv_count"] == 3

    # Low votes, adult titles, and shorts must be excluded.
    assert catalog.lookup_title("Too Few Votes") is None
    assert catalog.lookup_title("Adult Film") is None
    assert catalog.lookup_title("Some Short") is None

    # Duplicate (norm_title, kind, year) keeps the entry with higher votes.
    m = catalog.lookup_title("The Matrix", 1999, kind_hint="movie")
    assert m is not None
    assert m.votes == 2050000
    assert m.match_type == "exact"


def test_original_title_alias_and_fuzzy_year(sample_catalog):
    # Matched via French originalTitle, returns primaryTitle "Amélie".
    m = catalog.lookup_title("Le Fabuleux Destin d'Amelie Poulain", 2001)
    assert m is not None
    assert m.title == "Amélie"
    assert m.year == 2001

    # +-1 year festival vs wide release tolerance.
    m_fuzzy = catalog.lookup_title("Amelie", 2002, kind_hint="movie")
    assert m_fuzzy is not None
    assert m_fuzzy.title == "Amélie"
    assert m_fuzzy.year == 2001
    assert m_fuzzy.match_type == "year_fuzzy"


def test_prefix_lookup_strips_junk_but_never_truncates_real_titles(sample_catalog):
    # Trailing release/edition words are stripped via word-boundary prefix match.
    m = catalog.lookup_title("Inception Special Collector Cut", 2010, kind_hint="movie")
    assert m is not None
    assert m.title == "Inception"
    assert m.match_type == "prefix"

    # A real title ("It Comes at Night") whose first word matches another 2017
    # movie ("It") must NOT be truncated when the tail words are not junk.
    assert catalog.lookup_title("It Comes at Night", 2017, kind_hint="movie") is None

    # Sequel numbers and markers must never be stripped to match the original,
    # even when allow_any_tail=True and multiple words are in the tail.
    assert catalog.lookup_title("Inception 2", 2010, kind_hint="movie") is None
    assert catalog.lookup_title("Inception Part II", 2010, kind_hint="movie") is None
    assert catalog.lookup_title(
        "It Chapter 2", 2017, kind_hint="movie", allow_any_tail=True
    ) is None
    assert catalog.lookup_title(
        "It Story 2", 2017, kind_hint="movie", allow_any_tail=True
    ) is None


def test_parse_name_uses_catalog_as_backup(sample_catalog):
    # 1. Rescues source="raw" ("[REC].2007.1080p.mkv")
    p_rec = parse_name("[REC].2007.1080p.mkv", kind_hint="movie")
    assert p_rec.title == "REC"
    assert p_rec.year == 2007
    assert p_rec.source == "catalog"

    # 2. Strips trailing edition/release junk left by guessit
    p_inc = parse_name("Inception.Special.Collector.Cut.2010.1080p.mkv", kind_hint="movie")
    assert p_inc.title == "Inception"
    assert p_inc.year == 2010
    assert p_inc.source == "catalog"

    # 3. Fills in missing release year for bare numeric movie titles
    p_1917 = parse_name("1917.1080p.BluRay.mkv", kind_hint="movie")
    assert p_1917.title == "1917"
    assert p_1917.year == 2019
    assert p_1917.source == "catalog"

    # 4. Preserves guessit's original casing when guessit parsed cleanly
    assert parse_name("wall-e (2008).mkv").title == "wall-e"
    assert parse_name("SE7EN.1995.mkv").title == "SE7EN"

    # 5. Preserves country disambiguation ("The Office (US)")
    assert parse_name("The Office (US)", kind_hint="episode").title == "The Office (US)"


def test_clean_titles_with_catalog_returns_canonical_titles(sample_catalog):
    cleaned = catalog.clean_titles_with_catalog([
        "Its.Always.Sunny.in.Philadelphia.S14E03.1080p.mkv",
        "[REC].2007.1080p.mkv",
        "1917.1080p.BluRay.mkv",
        "The.Office.US.S01E01.mkv",
        "It.Comes.at.Night.2017.1080p.mkv",
    ])
    assert cleaned["Its.Always.Sunny.in.Philadelphia.S14E03.1080p.mkv"]["title"] == (
        "It's Always Sunny in Philadelphia"
    )
    assert cleaned["[REC].2007.1080p.mkv"] == {
        "title": "REC",
        "year": "2007",
        "quality": "1080p",
    }
    assert cleaned["1917.1080p.BluRay.mkv"] == {
        "title": "1917",
        "year": "2019",
        "quality": "1080p",
    }
    # Country suffix ("(US)") must be preserved in both lookup_title and clean_titles_with_catalog.
    assert catalog.lookup_title("The Office (US)", kind_hint="episode").title == "The Office (US)"
    assert cleaned["The.Office.US.S01E01.mkv"]["title"] == "The Office (US)"
    assert cleaned["The.Office.US.S01E01.mkv"]["year"] == ""
    # Episode lookup does not cross-match a movie ("It" 2017) when kind_hint="episode".
    assert catalog.lookup_title("It", kind_hint="episode") is None
    # Non-catalog movie whose prefix matches "It" (2017) must not be truncated.
    assert "It.Comes.at.Night.2017.1080p.mkv" not in cleaned


def test_cached_conn_closed_when_catalog_missing(sample_catalog, tmp_path):
    assert catalog.lookup_title("Inception", 2010) is not None
    assert catalog._cached_conn is not None
    assert catalog.lookup_title("Inception", 2010, db_path=tmp_path / "missing.db") is None
    assert catalog._cached_conn is None


def test_wizard_renaming_engine_option_4_uses_catalog(sample_catalog, monkeypatch):
    df_m = pd.DataFrame([{"Folder Name": "[REC].2007", "Video Files": "[REC].2007.1080p.mkv"}])
    df_t = pd.DataFrame([{"Show Folder": "Its Always Sunny in Philadelphia",
                          "Episode File": "Season 1/Its.Always.Sunny.in.Philadelphia.S01E01.mkv"}])
    answers = iter(["4", "y"])
    monkeypatch.setattr("builtins.input", lambda *a: next(answers))
    monkeypatch.setattr(wizard, "_stdin_is_interactive", lambda: True)

    res = wizard._ask_llm_results(df_m, df_t, [])
    assert res["[REC].2007"]["title"] == "REC"
    assert res["Its Always Sunny in Philadelphia"]["title"] == (
        "It's Always Sunny in Philadelphia"
    )


def test_doctor_validates_and_repairs_corrupted_catalog(tmp_path, monkeypatch):
    db_path = tmp_path / ".media_catalog.db"
    monkeypatch.setenv("MEDIAORG_CATALOG", str(db_path))

    # Absent catalog is OK (optional).
    status, msg = doctor.check_catalog()
    assert status == "OK" and "optional" in msg

    # Corrupted catalog is FAIL, and fix_catalog backs it up.
    db_path.write_text("not a sqlite database", encoding="utf-8")
    status, _ = doctor.check_catalog()
    assert status == "FAIL"

    fix_status, _ = doctor.fix_catalog()
    assert fix_status == "OK"
    assert not db_path.exists()
    assert db_path.with_suffix(".db.bak").exists()
