import datetime

import pytest

from mediaorg.parse import ParsedName, parse_name, pre_clean

# (name, kind_hint, expected fields)
CORPUS = [
    ("The.Matrix.1999.1080p.BluRay.x264-RARBG.mkv", None,
     dict(title="The Matrix", year=1999, quality="1080p", kind="movie")),
    ("WALL-E (2008).mkv", None, dict(title="WALL-E", year=2008)),
    ("Se7en.1995.REMASTERED.1080p.mkv", None, dict(title="Se7en", year=1995)),
    ("Blade.Runner.2049.2017.2160p.WEB-DL.mkv", None,
     dict(title="Blade Runner 2049", year=2017, quality="2160p")),
    ("1917 (2019).mkv", None, dict(title="1917", year=2019)),
    ("1917.mkv", None, dict(title="1917", year=None)),
    ("[REC].2007.mkv", None, dict(title="[REC]", source="raw")),
    ("Show.Name.S01E01E02.720p.mkv", "episode",
     dict(title="Show Name", season=1, episodes=[1, 2], kind="episode")),
    ("Show Name 1x02.mkv", "episode",
     dict(title="Show Name", season=1, episodes=[2])),
    ("The.Daily.Show.2024.01.15.Guest.Name.mkv", "episode",
     dict(title="The Daily Show", date=datetime.date(2024, 1, 15), kind="episode")),
    ("www.UIndex.org    -    Breaking.Bad.S01E01.720p.mkv", "episode",
     dict(title="Breaking Bad", season=1, episodes=[1])),
    # The old tool's self-inflicted "Ts.ts" artifact must parse clean.
    ("The Office S09E23 Ts.ts", "episode",
     dict(title="The Office", season=9, episodes=[23])),
    ("Its.Always.Sunny.in.Philadelphia.S14E03.1080p.WEB.x264-TBS.mkv", "episode",
     dict(title="Its Always Sunny in Philadelphia", season=14, episodes=[3])),
    # Round-trip stability: our own output must parse back unchanged.
    ("The Office (US)", "episode", dict(title="The Office (US)")),
    ("The Office (US) S01E01 - Pilot [720p].mkv", "episode",
     dict(title="The Office (US)", season=1, episodes=[1], quality="720p")),
    ("The Matrix (1999) [1080p]", None,
     dict(title="The Matrix", year=1999, quality="1080p")),
]


@pytest.mark.parametrize("name,hint,expected", CORPUS,
                         ids=[c[0] for c in CORPUS])
def test_corpus(name, hint, expected):
    p = parse_name(name, kind_hint=hint)
    for field_name, want in expected.items():
        assert getattr(p, field_name) == want, (
            f"{name}: {field_name}={getattr(p, field_name)!r}, wanted {want!r}")


@pytest.mark.xfail(strict=True, reason="guessit parses absolute anime numbering as SssEee")
def test_anime_absolute_numbering():
    p = parse_name("One Piece - 1071.mkv", kind_hint="episode")
    assert p.episodes == [1071]


def test_never_recases():
    # No .title() anywhere: weird original casing must survive.
    for name, want in [("wall-e (2008).mkv", "wall-e"),
                       ("SE7EN.1995.mkv", "SE7EN")]:
        assert parse_name(name).title == want


def test_pre_clean_html_entities():
    assert pre_clean("Tom &amp; Jerry &#039;79") == "Tom & Jerry '79"


def test_pre_clean_custom_patterns():
    assert pre_clean("MyTag Movie Name", ["MyTag"]) == "Movie Name"
    # A broken user regex must not crash.
    assert pre_clean("Movie", ["[bad"]) == "Movie"


def test_multi_episode_normalized_to_list():
    single = parse_name("Show S01E05.mkv", kind_hint="episode")
    assert single.episodes == [5]
    assert isinstance(single.episodes, list)


def test_defaults():
    p = ParsedName(title="X")
    assert p.episodes == [] and p.year is None and p.source == "guessit"
