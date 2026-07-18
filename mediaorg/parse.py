"""Filename parsing: guessit wrapper with a small pre-clean and fallback chain.

Rule: this module never transforms letter case. guessit preserves the
original casing of the title substring (WALL-E, Se7en stay intact).
"""

import datetime
import json
import re
from dataclasses import dataclass, field
from pathlib import Path

from guessit import guessit

VIDEO_EXTS = {'.mp4', '.mkv', '.avi', '.ts', '.m4v', '.wmv', '.mov', '.flv',
              '.webm', '.mpg', '.mpeg', '.m2ts'}
COMPANION_EXTS = {'.srt', '.sub', '.idx', '.ass', '.ssa', '.nfo', '.jpg',
                  '.jpeg', '.png', '.txt', '.vtt'}

CUSTOM_PATTERNS_FILE = "custom_strip_patterns.json"

# Only what guessit can't know: site prefixes, HTML entities, user patterns.
_WEBSITE_PREFIX = re.compile(r'^\s*(?:www\.\S+|https?://\S+)\s*[-–—]?\s*', re.IGNORECASE)
_HTML_ENTITY = re.compile(r'&#?\w+;')
_EXT_RE = re.compile(r'\.(' + '|'.join(e.lstrip('.') for e in VIDEO_EXTS) + r')$', re.IGNORECASE)


@dataclass
class ParsedName:
    title: str
    year: int | None = None
    season: int | None = None
    episodes: list[int] = field(default_factory=list)
    quality: str | None = None
    date: datetime.date | None = None
    episode_title: str | None = None
    kind: str = "unknown"      # "movie" | "episode" | "unknown"
    source: str = "guessit"    # "guessit" | "llm" | "raw"


def load_custom_patterns(folder: Path | str = ".") -> list[str]:
    path = Path(folder) / CUSTOM_PATTERNS_FILE
    if path.exists():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            return [p for p in data if isinstance(p, str)]
        except (json.JSONDecodeError, OSError):
            return []
    return []


def save_custom_patterns(patterns: list[str], folder: Path | str = ".") -> None:
    (Path(folder) / CUSTOM_PATTERNS_FILE).write_text(
        json.dumps(patterns, indent=2), encoding="utf-8")


def pre_clean(name: str, custom_patterns: list[str] = ()) -> str:
    """Strip things guessit can't be expected to understand."""
    s = _WEBSITE_PREFIX.sub('', name)
    s = s.replace('&amp;', '&').replace('&quot;', '"').replace('&#039;', "'")
    s = _HTML_ENTITY.sub('', s)
    for pat in custom_patterns:
        try:
            s = re.sub(pat, '', s, flags=re.IGNORECASE)
        except re.error:
            continue
    return re.sub(r'\s{2,}', ' ', s).strip()


def _strip_video_ext(name: str) -> str:
    return _EXT_RE.sub('', name)


def parse_name(name: str, kind_hint: str | None = None,
               custom_patterns: list[str] = ()) -> ParsedName:
    """Parse a file or folder name into its media parts.

    Fallback chain: guessit -> bare-number title ("1917.mkv") -> raw stem
    with separators mapped to spaces ("[REC]"), tagged source="raw" so the
    wizard can offer LLM cleaning for those.
    """
    cleaned = pre_clean(name, custom_patterns)
    options = {'type': kind_hint} if kind_hint else {}
    g = guessit(cleaned, options)

    eps = g.get('episode')
    episodes = [] if eps is None else (list(eps) if isinstance(eps, list) else [int(eps)])

    title = g.get('title')
    year = g.get('year')
    source = "guessit"

    # Fold the country back into the title ("The Office (US)") — otherwise
    # re-parsing our own output would strip it and renames never settle.
    if title and 'country' in g:
        country = g['country']
        code = getattr(country, 'alpha2', str(country))
        title = f"{title} ({code})"

    if not title:
        stem = _strip_video_ext(cleaned).strip()
        if year is not None and stem == str(year):
            # "1917.mkv": the whole name IS the title, not a year tag.
            title, year = str(year), None
        else:
            raw = stem
            if year is not None:
                raw = raw.replace(str(year), '')
            raw = re.sub(r'[._\-]+', ' ', raw)
            title = re.sub(r'\s{2,}', ' ', raw).strip()
            source = "raw"
        if not title:
            title = _strip_video_ext(cleaned).strip()
            source = "raw"

    return ParsedName(
        title=title,
        year=year,
        season=g.get('season') if not isinstance(g.get('season'), list) else g.get('season')[0],
        episodes=episodes,
        quality=str(g['screen_size']) if 'screen_size' in g else None,
        date=g.get('date'),
        episode_title=g.get('episode_title'),
        kind=g.get('type', 'unknown'),
        source=source,
    )
