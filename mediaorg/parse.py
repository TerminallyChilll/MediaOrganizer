"""Filename parsing: guessit wrapper with a small pre-clean and fallback chain.

Rule: this module never transforms letter case. guessit preserves the
original casing of the title substring (WALL-E, Se7en stay intact).
"""

import datetime
import json
import os
import re
from dataclasses import dataclass, field
from pathlib import Path

from guessit import guessit

VIDEO_EXTS = {'.mp4', '.mkv', '.avi', '.ts', '.m4v', '.wmv', '.mov', '.flv',
              '.webm', '.mpg', '.mpeg', '.m2ts'}
COMPANION_EXTS = {'.srt', '.sub', '.idx', '.ass', '.ssa', '.nfo', '.jpg',
                  '.jpeg', '.png', '.txt', '.vtt'}

# OS-generated metadata that carries no user data. These must never be treated
# as media: macOS AppleDouble sidecars ("._Show.S01E01.mkv") otherwise pass
# every video filter, collide with the real episode, and get BOTH skipped.
# They also block rmdir, which is what breaks undo on any volume Finder or
# Explorer has browsed.
JUNK_NAMES = {'.ds_store', 'thumbs.db', 'desktop.ini', '.localized',
              '.apdisk', '.spotlight-v100', '.trashes', '.fseventsd'}
JUNK_PREFIXES = ('._',)
JUNK_DIRS = {'@eadir', '.@__thumb', '.spotlight-v100', '.trashes',
             '.fseventsd', '$recycle.bin', 'system volume information'}


def is_junk_name(name: str) -> bool:
    """Is this an OS metadata file rather than media?"""
    lowered = name.lower()
    return lowered in JUNK_NAMES or lowered.startswith(JUNK_PREFIXES)


def is_junk_dir(name: str) -> bool:
    """Is this an OS/NAS bookkeeping directory that should never be scanned?"""
    return name.lower() in JUNK_DIRS


def is_media_file(name: str, exts=None) -> bool:
    """Does `name` look like a media file we should act on?

    Single gate used by every scanner and planner so junk filtering can never
    drift between them.
    """
    if is_junk_name(name):
        return False
    suffix = Path(name).suffix.lower()
    return suffix in (VIDEO_EXTS if exts is None else exts)


def companion_tail(companion_stem: str, video_stem: str) -> str | None:
    """The tail that makes `companion_stem` a sidecar of `video_stem`, else None.

    Sidecars are matched on a prefix rather than an exact stem so that language
    tags survive: "Episode1.en" against "Episode1" is ".en", and renaming the
    video has to carry that tail or the subtitle is orphaned.

    The boundary check is the whole point. A bare `startswith` also matches
    "Episode10.en" against "Episode1", so renaming Episode1 captures Episode
    *10*'s subtitle and renames it on top of a name it has no business touching.
    A prefix only counts when what follows it is not alphanumeric.

    Shared by the rename planner (`excel._companion_ops`) and the review screen
    (`wizard._edit_destination`), which must agree: the planner emits the
    destinations the review screen then has to recognise.
    """
    if not companion_stem.startswith(video_stem):
        return None
    tail = companion_stem[len(video_stem):]
    if tail and tail[0].isalnum():
        return None
    return tail    # '' for an exact match: the plain "Episode1.srt" case


CUSTOM_PATTERNS_FILE = "custom_strip_patterns.json"

# Suffix of the scratch file a case-only rename passes through. Defined here,
# with the other file-classification constants, so both the executor (which
# creates it) and extfix (which must never treat it as media) share one source
# of truth — and so planners never have to import the executor.
TMP_SUFFIX = ".mediaorg_tmp"

# Only what guessit can't know: site prefixes, HTML entities, user patterns.
# A trailing delimiter is REQUIRED. The old pattern ended in `\S+` with an
# optional dash, so with no whitespace in the name it swallowed the whole
# thing: 'http://x.com/The.Matrix.1999.mkv' became ''.
_WEBSITE_PREFIX = re.compile(
    r'^\s*(?:www\.|https?://)[^\s/\\]+(?:/[^\s]*?)?\s*[-–—_]\s*', re.IGNORECASE)
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
    # guessit parses these and they used to be discarded, which guaranteed a
    # collision between e.g. "Movie.1999.EXTENDED.mkv" and "Movie.1999.mkv":
    # both built the same target name, so check_collisions skipped BOTH and
    # neither ever got renamed.
    edition: str | None = None
    part: int | None = None


def custom_patterns_path(folder: Path | str | None = None) -> Path:
    """Where the custom word list lives, independent of the current directory.

    Same reasoning as the undo journal: a cwd-relative file meant that
    double-clicking the launcher, or running from anywhere else, silently
    started an empty list and the words the user had added stopped being
    stripped. Order: explicit folder -> ``$MEDIAORG_PATTERNS`` -> next to the
    app -> adopt a pre-existing file in the cwd, so upgrading users keep the
    list they already have.
    """
    if folder is not None:
        return Path(folder) / CUSTOM_PATTERNS_FILE
    env = os.environ.get("MEDIAORG_PATTERNS")
    if env:
        return Path(env).expanduser()
    app = Path(__file__).resolve().parent.parent / CUSTOM_PATTERNS_FILE
    if not app.exists():
        legacy = Path.cwd() / CUSTOM_PATTERNS_FILE
        if legacy.exists() and legacy != app:
            return legacy
    return app


def load_custom_patterns(folder: Path | str | None = None) -> list[str]:
    path = custom_patterns_path(folder)
    if path.exists():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            return [p for p in data if isinstance(p, str)]
        except (json.JSONDecodeError, OSError):
            return []
    return []


def save_custom_patterns(patterns: list[str],
                         folder: Path | str | None = None) -> None:
    """Write the word list. Raises OSError if the location is not writable.

    Callers are expected to handle that: the app directory can legitimately
    be read-only (a system-wide install, a read-only container mount), and a
    crash there would take the whole wizard down.
    """
    path = custom_patterns_path(folder)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(patterns, indent=2), encoding="utf-8")


def _strip_keeping_content(pattern, s: str, flags: int = 0) -> str:
    """Apply a substitution, but never let it empty the name.

    Belt and braces for over-greedy strip patterns (ours or a user's): a
    pattern that consumes everything leaves nothing for guessit to parse, and
    an empty title used to propagate all the way to a bare ".mkv" filename.
    """
    try:
        out = (pattern.sub('', s) if hasattr(pattern, 'sub')
               else re.sub(pattern, '', s, flags=flags))
    except re.error:
        return s
    return out if out.strip() else s


def pre_clean(name: str, custom_patterns: list[str] = ()) -> str:
    """Strip things guessit can't be expected to understand."""
    s = _strip_keeping_content(_WEBSITE_PREFIX, name)
    s = s.replace('&amp;', '&').replace('&quot;', '"').replace('&#039;', "'")
    s = s.replace('&#8217;', "'").replace('&rsquo;', "'").replace('&apos;', "'")
    s = _strip_keeping_content(_HTML_ENTITY, s)
    for pat in custom_patterns:
        s = _strip_keeping_content(pat, s, re.IGNORECASE)
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
    # Release group is never used by any name we build, and guessit's
    # heuristic for it eats real titles: in "WALL-E.2008.1080p.mkv" it reads
    # "WALL" as the group and leaves the title as "E". Turning the rule off
    # costs nothing and fixes every hyphenated title in that shape.
    options['excludes'] = ['release_group']
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
        if not title.endswith(f" ({code})"):
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

    def _first(key):
        val = g.get(key)
        return val[0] if isinstance(val, list) and val else val

    edition = _first('edition')
    # Explicit None check: `part or cd` would discard a legitimate part 0.
    part = _first('part')
    if part is None:
        part = _first('cd')
    try:
        part = int(part) if part is not None else None
    except (TypeError, ValueError):
        part = None

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
        edition=str(edition) if edition else None,
        part=part,
    )
