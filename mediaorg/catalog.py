"""Offline Movie & TV Show catalog (IMDb non-commercial datasets).

Builds and queries a compact local SQLite database (``.media_catalog.db``)
so filename parsing and title cleaning have an offline reference list when
``guessit`` falls back to a raw stem, leaves release/edition tokens attached
to a title, or encounters a bare numeric movie title like ``1917``.

Uses only Python's standard library (``sqlite3``, ``gzip``, ``urllib``).
ASCII markers only ([OK], [!], ->): no emoji, no cp1252 crashes.
"""

from __future__ import annotations

import gzip
import os
import re
import sqlite3
import ssl
import time
import unicodedata
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable

CATALOG_FILE = ".media_catalog.db"
RATINGS_URL = "https://datasets.imdbws.com/title.ratings.tsv.gz"
BASICS_URL = "https://datasets.imdbws.com/title.basics.tsv.gz"

DEFAULT_MIN_MOVIE_VOTES = 100
DEFAULT_MIN_TV_VOTES = 50
DOWNLOAD_TIMEOUT = 120

_MOVIE_TYPES = frozenset({"movie", "tvMovie"})
_TV_TYPES = frozenset({"tvSeries", "tvMiniSeries"})

_LATIN_FOLD = str.maketrans({
    "ß": "ss",
    "æ": "ae",
    "œ": "oe",
    "ø": "o",
    "ł": "l",
    "đ": "d",
    "ð": "d",
    "þ": "th",
})
_APOSTROPHE_RE = re.compile(r"['\u2018\u2019\u02bc`]")
_NON_WORD_RE = re.compile(r"[^\w]+|_+", re.UNICODE)
_COUNTRY_SUFFIX_RE = re.compile(r"\s*\(([A-Za-z]{2,3})\)\s*$")
_RESOLUTION_OR_CODEC_RE = re.compile(
    r"^(?:\d{3,4}[pi]|[248]k|[xh]26[45]|av1|hevc|divx|xvid|10bit|8bit|"
    r"hdr(?:10(?:\+|plus)?)?|dv|dolby|atmos|aac\d*|ac3|eac3|dts(?:-?hd)?|"
    r"truehd|flac|ddp?\d+|\d\.\d)$",
    re.IGNORECASE,
)

# Sequel boundaries that prefix matching must never strip across.
_SEQUEL_MARKERS = frozenset({
    "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x",
    "part", "pt", "vol", "volume", "chapter",
})

# Bare leading articles that cannot stand alone as a prefix match when stripping
# arbitrary raw tails.
_BARE_ARTICLES = frozenset({"a", "an", "the"})

# Release, edition, technical, and ripper tokens that guessit occasionally
# leaves attached to a title when they appear in non-standard positions.
_JUNK_TAIL_WORDS = frozenset({
    "extended", "unrated", "uncut", "director", "directors", "cut",
    "collector", "collectors", "special", "edition", "ed", "theatrical",
    "remastered", "restored", "imax", "criterion", "anniversary",
    "ultimate", "final", "redux", "complete", "series", "collection",
    "pack", "anthology", "trilogy", "proper", "repack", "rerip",
    "internal", "limited", "festival", "retail", "custom", "dubbed",
    "subbed", "dual", "audio", "multi", "multisub", "open", "matte",
    "3d", "hsbs", "sbs", "hou", "uhd", "sdr", "hdr", "bluray", "bdrip",
    "brrip", "webrip", "webdl", "web", "dvdrip", "hdrip", "hdtv",
    "remux", "rarbg", "yts", "yify", "psa", "qxr", "tigole", "bokutox",
    "lama", "etrg", "evo", "fgt", "ntb", "flux", "galaxyrg", "amzn",
    "nf", "dsnp", "atvp", "hmax", "pcok",
})


@dataclass(frozen=True)
class CatalogMatch:
    title: str
    year: int | None
    kind: str           # "movie" | "tv"
    votes: int
    match_type: str     # "exact" | "year_fuzzy" | "prefix" | "unambiguous"


def catalog_path(folder: Path | str | None = None) -> Path:
    """Where the offline catalog SQLite DB lives, independent of the cwd.

    Order: explicit folder -> ``$MEDIAORG_CATALOG`` -> next to the app ->
    adopt a pre-existing file in the cwd.
    """
    if folder is not None:
        return Path(folder) / CATALOG_FILE
    env = os.environ.get("MEDIAORG_CATALOG")
    if env:
        return Path(env).expanduser()
    app = Path(__file__).resolve().parent.parent / CATALOG_FILE
    if not app.exists():
        legacy = Path.cwd() / CATALOG_FILE
        if legacy.exists() and legacy != app:
            return legacy
    return app


def normalize_title(text: str) -> str:
    """Normalize a title for catalog indexing and lookup.

    Strips accents and apostrophes, maps '&' to 'and', and collapses
    punctuation/whitespace so 'WALL-E', 'WALL·E', and 'wall e' all normalize
    to 'wall e', and "It's Always Sunny" matches 'Its Always Sunny'.
    """
    if not text:
        return ""
    s = unicodedata.normalize("NFKD", str(text))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower().translate(_LATIN_FOLD)
    s = s.replace("&", " and ")
    s = _APOSTROPHE_RE.sub("", s)
    s = _NON_WORD_RE.sub(" ", s)
    return s.strip()


def _is_junk_tail_word(word: str) -> bool:
    w = word.lower()
    return w in _JUNK_TAIL_WORDS or bool(_RESOLUTION_OR_CODEC_RE.match(w))


def _map_kind_hint(kind_hint: str | None) -> str | None:
    if not kind_hint:
        return None
    k = kind_hint.lower()
    if k in ("movie", "tvmovie"):
        return "movie"
    if k in ("episode", "tv", "tvseries", "tvminiseries"):
        return "tv"
    return None


# --- Connection caching ------------------------------------------------------

_cached_conn: sqlite3.Connection | None = None
_cached_sig: tuple[str, int, int] | None = None


def _close_cached() -> None:
    global _cached_conn, _cached_sig
    if _cached_conn is not None:
        try:
            _cached_conn.close()
        except sqlite3.Error:
            pass
    _cached_conn = None
    _cached_sig = None


def _get_conn(db_path: Path | None = None) -> sqlite3.Connection | None:
    """Return a read-only SQLite connection to the catalog, or None if absent."""
    global _cached_conn, _cached_sig
    path = Path(db_path) if db_path is not None else catalog_path()
    try:
        st = path.stat()
    except OSError:
        _close_cached()
        return None
    if st.st_size == 0:
        _close_cached()
        return None
    sig = (str(path.resolve()), st.st_mtime_ns, st.st_size)
    if _cached_conn is not None and _cached_sig == sig:
        return _cached_conn
    _close_cached()
    try:
        uri = f"{path.resolve().as_uri()}?mode=ro"
        conn = sqlite3.connect(uri, uri=True, check_same_thread=False)
        _cached_conn = conn
        _cached_sig = sig
        return conn
    except (sqlite3.Error, ValueError):
        return None


def catalog_stats(db_path: Path | None = None) -> dict | None:
    """Return summary metadata for an existing catalog, or None if missing/corrupt."""
    path = Path(db_path) if db_path is not None else catalog_path()
    conn = _get_conn(path)
    if conn is None:
        return None
    try:
        meta = dict(conn.execute("SELECT key, value FROM meta").fetchall())
        movie_count = int(meta.get("movie_count", 0))
        tv_count = int(meta.get("tv_count", 0))
        if not movie_count and not tv_count:
            total = conn.execute("SELECT COUNT(*) FROM titles").fetchone()[0]
        else:
            total = movie_count + tv_count
        return {
            "path": path,
            "movie_count": movie_count,
            "tv_count": tv_count,
            "total": total,
            "built_at": float(meta.get("built_at", 0.0)),
            "size_mb": round(path.stat().st_size / (1024 * 1024), 2),
        }
    except (sqlite3.Error, ValueError, OSError):
        return None


# --- Lookup ------------------------------------------------------------------

def _query_exact(conn: sqlite3.Connection, norm: str, year: int,
                 target_kind: str | None) -> CatalogMatch | None:
    if target_kind:
        row = conn.execute(
            "SELECT title, year, kind, votes FROM titles "
            "WHERE norm_title = ? AND kind = ? AND year = ? "
            "ORDER BY votes DESC LIMIT 1",
            (norm, target_kind, year),
        ).fetchone()
    else:
        row = conn.execute(
            "SELECT title, year, kind, votes FROM titles "
            "WHERE norm_title = ? AND year = ? "
            "ORDER BY votes DESC LIMIT 1",
            (norm, year),
        ).fetchone()
    if row:
        return CatalogMatch(row[0], row[1] or None, row[2], row[3], "exact")
    return None


def _query_year_fuzzy(conn: sqlite3.Connection, norm: str, year: int,
                      target_kind: str | None) -> CatalogMatch | None:
    if target_kind:
        row = conn.execute(
            "SELECT title, year, kind, votes FROM titles "
            "WHERE norm_title = ? AND kind = ? AND year BETWEEN ? AND ? "
            "ORDER BY ABS(year - ?) ASC, votes DESC LIMIT 1",
            (norm, target_kind, year - 1, year + 1, year),
        ).fetchone()
    else:
        row = conn.execute(
            "SELECT title, year, kind, votes FROM titles "
            "WHERE norm_title = ? AND year BETWEEN ? AND ? "
            "ORDER BY ABS(year - ?) ASC, votes DESC LIMIT 1",
            (norm, year - 1, year + 1, year),
        ).fetchone()
    if row:
        mtype = "exact" if row[1] == year else "year_fuzzy"
        return CatalogMatch(row[0], row[1] or None, row[2], row[3], mtype)
    return None


def _query_unambiguous(conn: sqlite3.Connection, norm: str,
                       target_kind: str | None) -> CatalogMatch | None:
    if target_kind:
        row = conn.execute(
            "SELECT title, year, kind, votes FROM titles "
            "WHERE norm_title = ? AND kind = ? "
            "ORDER BY votes DESC LIMIT 1",
            (norm, target_kind),
        ).fetchone()
    else:
        row = conn.execute(
            "SELECT title, year, kind, votes FROM titles "
            "WHERE norm_title = ? "
            "ORDER BY votes DESC LIMIT 1",
            (norm,),
        ).fetchone()
    if row:
        return CatalogMatch(row[0], row[1] or None, row[2], row[3], "unambiguous")
    return None


def lookup_title(title: str, year: int | None = None,
                 kind_hint: str | None = None,
                 db_path: Path | None = None, *,
                 allow_any_tail: bool = False) -> CatalogMatch | None:
    """Look up `(title, year, kind_hint)` in the local SQLite catalog.

    Returns a :class:`CatalogMatch` or ``None`` if the catalog is not
    installed or no match is found.
    """
    if not title:
        return None
    conn = _get_conn(db_path)
    if conn is None:
        return None

    norm = normalize_title(title)
    if not norm:
        return None
    target_kind = _map_kind_hint(kind_hint)
    country_match = _COUNTRY_SUFFIX_RE.search(title) if not allow_any_tail else None

    try:
        # Tier 1 & 2: Exact normalized title + exact or +-1 year.
        if year is not None:
            match = _query_exact(conn, norm, year, target_kind)
            if match is not None:
                return match
            match = _query_year_fuzzy(conn, norm, year, target_kind)
            if match is not None:
                return match

        # Check exact normalized title before prefix stripping so a full title
        # already in the catalog is never truncated to a shorter prefix.
        match = _query_unambiguous(conn, norm, target_kind)
        if match is not None:
            return match

        # If the title carries a parenthesized country code ("The Office (US)"),
        # match the base title in the catalog while preserving the country tag.
        if country_match:
            base_title = title[:country_match.start()].strip()
            base_m = lookup_title(
                base_title,
                year,
                kind_hint=kind_hint,
                db_path=db_path,
                allow_any_tail=False,
            )
            if base_m is not None:
                code = country_match.group(1)
                return CatalogMatch(
                    f"{base_m.title} ({code})",
                    year,
                    base_m.kind,
                    base_m.votes,
                    base_m.match_type,
                )
            return None

        # Tier 3: Word-boundary prefix match (longest prefix wins).
        words = norm.split()
        if len(words) >= 2:
            for k in range(len(words) - 1, 0, -1):
                tail = words[k:]
                # Never strip across a sequel number or marker ("Part 2", "II")
                # anywhere in the stripped tail unless the token is a known
                # technical/release tag ("1080p").
                if any(
                    not _is_junk_tail_word(w)
                    and (w.isdigit() or w in _SEQUEL_MARKERS)
                    for w in tail
                ):
                    continue
                if not allow_any_tail and not all(_is_junk_tail_word(w) for w in tail):
                    continue
                prefix = " ".join(words[:k])
                if allow_any_tail and prefix in _BARE_ARTICLES:
                    continue
                if year is not None:
                    m = (_query_exact(conn, prefix, year, target_kind)
                         or _query_year_fuzzy(conn, prefix, year, target_kind))
                else:
                    m = _query_unambiguous(conn, prefix, target_kind)
                if m is not None:
                    return CatalogMatch(m.title, m.year, m.kind, m.votes, "prefix")
    except sqlite3.Error:
        return None

    return None


def refine_parsed(parsed, kind_hint: str | None = None,
                  db_path: Path | None = None):
    """Use the local catalog as a non-destructive backup for a `ParsedName`.

    Preserves `guessit`'s original title casing when `guessit` already parsed
    the title cleanly (`test_never_recases`), and steps in only when:
    - `parsed.source == "raw"` (guessit failed to extract a title),
    - `guessit` left trailing release/edition junk tokens attached to the
      title (`match_type == "prefix"`), or
    - a bare numeric movie title (`1917.1080p.BluRay.mkv`) is missing its
      release year.
    """
    if not parsed.title:
        return parsed
    # Preserve country-disambiguated titles ("The Office (US)") untouched.
    if parsed.source != "raw" and _COUNTRY_SUFFIX_RE.search(parsed.title):
        return parsed

    hint = kind_hint or ("episode" if parsed.kind == "episode" else None)
    is_raw = parsed.source == "raw"
    match = lookup_title(
        parsed.title,
        parsed.year,
        kind_hint=hint,
        db_path=db_path,
        allow_any_tail=is_raw,
    )
    if match is None:
        return parsed

    if is_raw:
        parsed.title = match.title
        if parsed.year is None and match.year is not None:
            parsed.year = match.year
        parsed.source = "catalog"
    elif match.match_type == "prefix":
        parsed.title = match.title
        if parsed.year is None and match.year is not None:
            parsed.year = match.year
        parsed.source = "catalog"
    elif (parsed.year is None
          and match.year is not None
          and parsed.title.strip().isdigit()
          and hint != "episode"
          and parsed.season is None
          and not parsed.episodes):
        parsed.year = match.year
        parsed.source = "catalog"

    return parsed


def clean_titles_with_catalog(filenames: Iterable[str],
                              custom_patterns: list[str] = (),
                              db_path: Path | None = None) -> dict[str, dict]:
    """Match a batch of filenames against the local catalog.

    Returns ``{original_filename: {'title': ..., 'year': ..., 'quality': ...}}``
    in the same format as :func:`mediaorg.llm.clean_titles_with_llm` so the
    rename planner can consume either interchangeably.
    """
    from .parse import parse_name

    results: dict[str, dict] = {}
    for name in filenames:
        p = parse_name(name, custom_patterns=custom_patterns, catalog_db=db_path)
        m = lookup_title(
            p.title,
            p.year,
            kind_hint="episode" if p.kind == "episode" else None,
            db_path=db_path,
            allow_any_tail=(p.source == "raw"),
        )
        if m is not None:
            is_ep = p.kind == "episode" or p.season is not None or bool(p.episodes)
            year_val = p.year if is_ep else (p.year or m.year)
            results[str(name)] = {
                "title": m.title,
                "year": str(year_val) if year_val else "",
                "quality": str(p.quality) if p.quality else "",
            }
    return results


# --- Catalog Builder / Updater -----------------------------------------------

def _open_tsv_lines(source: str | Path, timeout: int = DOWNLOAD_TIMEOUT):
    """Yield UTF-8 text lines from a local `.tsv`/`.tsv.gz` path or a URL."""
    src_str = str(source)
    if src_str.startswith(("http://", "https://")):
        ctx = ssl.create_default_context()
        req = urllib.request.Request(
            src_str,
            headers={"User-Agent": "MediaOrganizer/1.0 (+offline-catalog)"},
        )
        resp = urllib.request.urlopen(req, timeout=timeout, context=ctx)
        try:
            with gzip.GzipFile(fileobj=resp) as gz:
                for raw in gz:
                    yield raw.decode("utf-8", errors="replace")
        finally:
            resp.close()
        return

    path = Path(source)
    if path.suffix.lower() == ".gz":
        with gzip.open(path, "rt", encoding="utf-8", errors="replace") as fh:
            yield from fh
    else:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            yield from fh


def build_catalog(
    db_path: Path | str | None = None,
    *,
    basics_source: str | Path = BASICS_URL,
    ratings_source: str | Path = RATINGS_URL,
    min_movie_votes: int = DEFAULT_MIN_MOVIE_VOTES,
    min_tv_votes: int = DEFAULT_MIN_TV_VOTES,
    timeout: int = DOWNLOAD_TIMEOUT,
    progress_cb: Callable[[str], None] | None = None,
) -> dict:
    """Download/read IMDb datasets and build the local SQLite catalog atomically.

    Returns a stats dict ``{'path', 'movie_count', 'tv_count', 'total', 'size_mb'}``.
    """
    target = Path(db_path) if db_path is not None else catalog_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp_prefix = target.name if target.name.startswith(".") else f".{target.name}"
    tmp_path = target.parent / f"{tmp_prefix}.{os.getpid()}.tmp"

    def _log(msg: str) -> None:
        if progress_cb is not None:
            progress_cb(msg)

    floor = min(min_movie_votes, min_tv_votes)
    _log(f"Reading title ratings (min votes: movies={min_movie_votes}, TV={min_tv_votes})...")
    votes_by_id: dict[str, int] = {}
    first = True
    for line in _open_tsv_lines(ratings_source, timeout=timeout):
        if first:
            first = False
            continue
        parts = line.rstrip("\r\n").split("\t")
        if len(parts) < 3:
            continue
        try:
            v = int(parts[2])
        except ValueError:
            continue
        if v >= floor:
            votes_by_id[parts[0]] = v

    _log(f"Loaded {len(votes_by_id):,} rated titles. Building SQLite index...")
    if tmp_path.exists():
        try:
            tmp_path.unlink()
        except OSError:
            pass

    conn = sqlite3.connect(str(tmp_path))
    try:
        conn.execute("PRAGMA journal_mode = OFF")
        conn.execute("PRAGMA synchronous = OFF")
        conn.execute(
            "CREATE TABLE titles ("
            "  norm_title TEXT NOT NULL,"
            "  kind TEXT NOT NULL,"
            "  year INTEGER NOT NULL DEFAULT 0,"
            "  title TEXT NOT NULL,"
            "  votes INTEGER NOT NULL DEFAULT 0,"
            "  PRIMARY KEY (norm_title, kind, year)"
            ") WITHOUT ROWID"
        )
        conn.execute(
            "CREATE TABLE meta ("
            "  key TEXT PRIMARY KEY,"
            "  value TEXT NOT NULL"
            ")"
        )

        upsert_sql = (
            "INSERT INTO titles (norm_title, kind, year, title, votes) "
            "VALUES (?, ?, ?, ?, ?) "
            "ON CONFLICT(norm_title, kind, year) DO UPDATE SET "
            "  title = excluded.title, "
            "  votes = excluded.votes "
            "WHERE excluded.votes > titles.votes"
        )

        batch: list[tuple[str, str, int, str, int]] = []
        movie_count = 0
        tv_count = 0
        first = True

        for line in _open_tsv_lines(basics_source, timeout=timeout):
            if first:
                first = False
                continue
            parts = line.rstrip("\r\n").split("\t")
            if len(parts) < 6:
                continue
            tconst, title_type, primary_title, original_title, is_adult, start_year = (
                parts[0], parts[1], parts[2], parts[3], parts[4], parts[5]
            )
            if is_adult != "0":
                continue
            if title_type in _MOVIE_TYPES:
                kind = "movie"
                min_votes = min_movie_votes
            elif title_type in _TV_TYPES:
                kind = "tv"
                min_votes = min_tv_votes
            else:
                continue

            num_votes = votes_by_id.get(tconst, 0)
            if num_votes < min_votes:
                continue

            year_val = int(start_year) if start_year.isdigit() else 0
            norm_primary = normalize_title(primary_title)
            if not norm_primary:
                continue

            if kind == "movie":
                movie_count += 1
            else:
                tv_count += 1

            batch.append((norm_primary, kind, year_val, primary_title, num_votes))
            if original_title and original_title != r"\N":
                norm_orig = normalize_title(original_title)
                if norm_orig and norm_orig != norm_primary:
                    batch.append((norm_orig, kind, year_val, primary_title, num_votes))

            if len(batch) >= 10000:
                conn.executemany(upsert_sql, batch)
                batch.clear()

        if batch:
            conn.executemany(upsert_sql, batch)
            batch.clear()

        built_at = time.time()
        conn.executemany(
            "INSERT INTO meta (key, value) VALUES (?, ?)",
            [
                ("built_at", str(built_at)),
                ("movie_count", str(movie_count)),
                ("tv_count", str(tv_count)),
                ("min_movie_votes", str(min_movie_votes)),
                ("min_tv_votes", str(min_tv_votes)),
            ],
        )
        conn.commit()
        conn.close()
        conn = None

        _close_cached()
        os.replace(tmp_path, target)
    finally:
        if conn is not None:
            try:
                conn.close()
            except sqlite3.Error:
                pass
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass

    size_mb = round(target.stat().st_size / (1024 * 1024), 2)
    _log(f"Catalog ready: {movie_count:,} movies, {tv_count:,} TV series ({size_mb} MB).")
    return {
        "path": target,
        "movie_count": movie_count,
        "tv_count": tv_count,
        "total": movie_count + tv_count,
        "built_at": built_at,
        "size_mb": size_mb,
    }
