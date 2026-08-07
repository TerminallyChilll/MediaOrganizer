"""Pure planning: inspect disk (read-only), emit operations.

Nothing in this module writes to disk. Every mutation the tool ever makes
is an Op produced here (or in extfix) and executed by mediaorg.execute.
"""

import functools
import os
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

from .parse import (COMPANION_EXTS, VIDEO_EXTS, ParsedName, is_junk_dir,
                    is_junk_name, is_media_file)


# --- Unicode / filesystem helpers -------------------------------------------

def norm(name: str) -> str:
    """Unicode-normalise a name for comparison.

    macOS returns decomposed (NFD) names from ``scandir`` while guessit emits
    composed (NFC) titles, so raw string equality reports "needs renaming" on
    every single run for any accented title and renames never converge. Every
    idempotency comparison must run both sides through this.
    """
    return unicodedata.normalize('NFC', name or '')


@functools.lru_cache(maxsize=256)
def _dir_is_case_sensitive(directory: str) -> bool:
    """Probe (read-only) whether `directory`'s filesystem is case-sensitive.

    ``os.path.normcase`` is the identity function on POSIX, so it cannot tell
    us this: on macOS/APFS and on Linux with an NTFS/exFAT drive — the most
    common media-library setups — case collisions went completely undetected.
    """
    d = Path(directory)
    try:
        for entry in os.scandir(d):
            swapped = entry.name.swapcase()
            if swapped == entry.name:
                continue
            return not os.path.lexists(d / swapped)
    except OSError:
        pass
    # Nothing to probe with: fall back to the platform default.
    return os.name != 'nt' and sys.platform != 'darwin'


def _nearest_existing(path: Path) -> Path:
    p = Path(os.path.abspath(path))
    while not p.exists():
        parent = p.parent
        if parent == p:
            return p
        p = parent
    return p


def _collision_key(p: Path) -> str:
    """Key under which two destinations count as "the same target"."""
    key = norm(str(p))
    if not _dir_is_case_sensitive(str(_nearest_existing(p.parent))):
        key = key.casefold()
    return os.path.normcase(key)


@functools.lru_cache(maxsize=1)
def _windows_long_paths_enabled() -> bool:
    if os.name != 'nt':
        return True
    try:
        import winreg
        key = r"SYSTEM\CurrentControlSet\Control\FileSystem"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key) as handle:
            return bool(winreg.QueryValueEx(handle, "LongPathsEnabled")[0])
    except Exception:
        return False


def max_path_length() -> int:
    """Longest full path the current platform will accept."""
    if os.name == 'nt' and not _windows_long_paths_enabled():
        return 259  # MAX_PATH (260) minus the terminating NUL
    return 4095

# --- Op model ---------------------------------------------------------------

@dataclass(frozen=True)
class Op:
    kind: str          # "move" | "mkdir" | "rmdir" (rmdir = only-if-empty)
    src: Path | None   # None for mkdir/rmdir
    dst: Path


@dataclass
class Plan:
    ops: list[Op] = field(default_factory=list)
    skipped: list[tuple[Op, str]] = field(default_factory=list)

    def merge(self, other: "Plan") -> None:
        self.ops.extend(other.ops)
        self.skipped.extend(other.skipped)


# --- Episode/season detection (name-level, regex — tuned patterns ported) ---

# Season allows 4 digits for date-style seasons ("S2024E05"); the episode's
# trailing (?!\d) is load-bearing — with a bare \d{1,2} the pattern truncated
# "S01E101" to episode 10, silently misfiling every 3-digit episode.
_SE_CORE = r'[Ss](\d{1,4})[Ee](\d{1,3})(?!\d)'

# SxxEyy (group 1=season) or bare NxN like "1x2"/"9x9" (groups 2, 3).
EPISODE_PATTERN = re.compile(
    r'(?:[Ss](\d{1,4})[Ee]\d{1,3}(?!\d)|(?<!\d)(\d{1,2})[xX](\d{1,2})(?!\d))')
# Pure season folders: "Season 1", "S01", "Season 1 (2024)".
SEASON_FOLDER_PATTERN = re.compile(r'^(?:season\s*|s)(\d{1,4})(?:\s*\(\d{4}\))?$', re.IGNORECASE)
# Already-normalized names the organizer must leave untouched (incl. the
# year-tagged form the rename scheme produces). Zero-padding is deliberately
# NOT canonical: accepting "Season 01" here left organize skipping it while
# the renamer looked for the literal "Season 1", so it could never be fixed.
_CANONICAL_SEASON = re.compile(r'^Season (0|[1-9]\d{0,3})( \(\d{4}\))?$')
# Abbreviated "ShowTitle S02" form only — no dashes/pipes/dots, optional year.
SEASON_LIKE_PATTERN = re.compile(r'^([^\-\|.]+?)\s+[Ss](\d{1,4})(?:\s*\(?\d{4}\)?)?\s*$')
# Specials live outside the numbered seasons but still mark their parent as a
# show: a show whose only content is a "Specials" folder is still a show, and
# without this the show search walks straight past it. Defined here, with the
# other name-classification patterns, so the scanner and the organizer cannot
# disagree about what counts as a show.
SPECIALS_FOLDER_PATTERN = re.compile(
    r'^(?:specials?|extras?|featurettes?|bonus)$', re.IGNORECASE)

_SXXEYY = re.compile(_SE_CORE)
_NXN = re.compile(r'(?<!\d)(\d{1,2})[xX](\d{1,2})(?!\d)')


def extract_season_episode(name: str) -> tuple[int | None, int | None]:
    m = _SXXEYY.search(name) or _NXN.search(name)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


def _is_show_marker_dir(name: str) -> bool:
    """Does a subdirectory with this name mark its parent as a show folder?

    Folder names only count with a strict SxxEyy marker — bare NxN like
    "9x9" appears in junk names ("WATCH - Show 9x9 - FREE") and must not
    classify. NxN still counts for files and inside season folders.
    """
    return bool(_SXXEYY.search(name)
                or SEASON_FOLDER_PATTERN.match(name.strip())
                or SPECIALS_FOLDER_PATTERN.match(name.strip())
                or SEASON_LIKE_PATTERN.match(name))


def has_season_structure(path: Path) -> bool:
    """Does this folder hold season/specials/episode *subfolders*?

    The distinction from :func:`folder_has_episodes_or_seasons` matters when
    deciding how deep a show lives: a folder with real season folders is a
    show, whereas one holding only loose episode files might instead be a dump
    folder sitting inside the show.
    """
    try:
        return any(e.is_dir(follow_symlinks=False) and _is_show_marker_dir(e.name)
                   for e in os.scandir(path))
    except OSError:
        return False


def folder_has_episodes_or_seasons(path: Path) -> bool:
    """Is this folder a TV-show root (contains episodes/seasons)?"""
    try:
        items = list(os.scandir(path))
    except OSError:
        return False
    for entry in items:
        name = entry.name
        if entry.is_dir(follow_symlinks=False):
            if _is_show_marker_dir(name):
                return True
        elif entry.is_file(follow_symlinks=False):
            if is_media_file(name) and EPISODE_PATTERN.search(name):
                return True
    return False


# How far below the library root the show search will descend. Wrapper
# layouts in the wild go about as deep as "Genre/SubGenre/Show"; four levels
# covers those without walking an entire NAS looking for TV that isn't there.
SHOW_SEARCH_MAX_DEPTH = 4


def find_show_roots(library_root: Path,
                    max_depth: int = SHOW_SEARCH_MAX_DEPTH) -> list[Path]:
    """Every show folder under `library_root`, however deeply it is nested.

    A show root is the *shallowest* directory that directly contains season
    folders or SxxEyy episode files. The search stops descending as soon as it
    finds one, so a "Season 1" folder — which also looks episode-bearing — can
    never be mistaken for a show of its own.

    This replaces a hard-coded two-level descent that silently did nothing for
    anything deeper (``Genre/SubGenre/Show/Season 1``) and that left the
    scanner and the organizer disagreeing about where the shows were.
    """
    library_root = Path(library_root)
    if folder_has_episodes_or_seasons(library_root):
        return [library_root]

    found: list[Path] = []

    def descend(directory: Path, depth: int) -> None:
        if depth > max_depth:
            return
        try:
            children = sorted(
                (Path(e.path) for e in os.scandir(directory)
                 if e.is_dir(follow_symlinks=False) and not is_junk_dir(e.name)),
                key=lambda p: p.name)
        except OSError:
            return
        for child in children:
            if not folder_has_episodes_or_seasons(child):
                descend(child, depth + 1)
                continue
            # A folder qualifying only because it holds loose episode files is
            # the show itself when it sits at the top of the library, but two
            # levels down it is far more likely a dump folder *inside* the
            # show ("Show/Downloads/Show.S01E01.mkv"). Credit the parent, and
            # let plan_season_structure lift the episodes up into it —
            # otherwise the dump folder's name becomes the show's title.
            if depth > 1 and not has_season_structure(child):
                if directory not in found:
                    found.append(directory)
            else:
                found.append(child)

    descend(library_root, 1)
    return found


def companion_files(video: Path) -> list[Path]:
    """Non-video files next to `video` sharing its stem (subs, nfo, art)."""
    out = []
    try:
        for entry in os.scandir(video.parent):
            if not entry.is_file(follow_symlinks=False):
                continue
            p = Path(entry.path)
            if p == video:
                continue
            if (p.suffix.lower() in COMPANION_EXTS
                    and not is_junk_name(p.name)
                    and norm(p.name).startswith(norm(video.stem))):
                out.append(p)
    except OSError:
        pass
    return sorted(out)


# --- Collision checking -----------------------------------------------------

def _same_file(a: Path, b: Path) -> bool:
    try:
        return a.samefile(b)
    except OSError:
        return False


def _is_case_only_rename(src: Path, dst: Path) -> bool:
    return (src.parent == dst.parent
            and src.name != dst.name
            and src.name.casefold() == dst.name.casefold())


def check_collisions(ops: list[Op]) -> Plan:
    """Split ops into safe vs skipped. Never overwrite, never auto-suffix."""
    plan = Plan()
    limit = max_path_length()
    by_dst: dict[str, list[Op]] = {}
    for op in ops:
        by_dst.setdefault(_collision_key(op.dst), []).append(op)

    for op in ops:
        if op.kind == "move":
            group = by_dst[_collision_key(op.dst)]
            if sum(1 for o in group if o.kind == "move") > 1:
                plan.skipped.append((op, "conflict: multiple sources map to same target"))
                continue
            if len(str(op.dst)) > limit:
                plan.skipped.append((
                    op, f"destination path is {len(str(op.dst))} characters, "
                        f"over this system's {limit} limit"))
                continue
            if op.dst.exists():
                # samefile() is also true for hardlinks, so it must not be
                # used on its own to wave a move through: doing so sent a
                # hardlink pair down the two-step rename path and destroyed
                # one of the two directory entries, unjournaled.
                if not (op.src and _is_case_only_rename(op.src, op.dst)
                        and _same_file(op.src, op.dst)):
                    plan.skipped.append((op, f"conflict: target already exists: {op.dst}"))
                    continue
            if op.src and not os.path.lexists(op.src):
                plan.skipped.append((op, f"source missing: {op.src}"))
                continue
        plan.ops.append(op)
    return plan


# --- TV season structure planner --------------------------------------------

def _lift_media(source_dir: Path, dest_for, *, include_root: bool) -> list[Op]:
    """Move media files nested below `source_dir` to where `dest_for` says.

    ``dest_for(filename)`` returns the directory a file belongs in, or None to
    leave it alone — so a dump folder holding two seasons routes each file to
    its own season folder in a single pass.

    Walked bottom-up, so a directory is only removed after the files inside it
    have been planned out of it, and an emptied child counts as gone when its
    parent is considered. `include_root` also empties (and removes)
    `source_dir` itself rather than only what is nested below it.

    A directory is only rmdir'd when nothing will be left in it: ``rmdir`` is
    empty-only, so emitting one for a directory that still holds, say, a stray
    .exe would just produce a failed op and a confusing report. OS junk
    (.DS_Store, Thumbs.db) does not count as a leftover — the executor
    quarantines that on its own.
    """
    ops: list[Op] = []
    media_exts = VIDEO_EXTS | COMPANION_EXTS
    emptied: set[Path] = set()
    try:
        walked = list(os.walk(source_dir, topdown=False))
    except OSError:
        return ops

    for dirpath, dirnames, filenames in walked:
        current = Path(dirpath)
        if is_junk_dir(current.name):
            continue
        if current == source_dir and not include_root:
            continue

        staying = []
        for name in sorted(filenames):
            dest = dest_for(name) if is_media_file(name, media_exts) else None
            if dest is not None and dest != current:
                ops.append(Op("move", current / name, dest / name))
            elif not is_junk_name(name):
                staying.append(name)

        leftover_dirs = [d for d in dirnames
                         if not is_junk_dir(d) and current / d not in emptied]
        if not staying and not leftover_dirs:
            ops.append(Op("rmdir", None, current))
            emptied.add(current)
    return ops


def plan_season_structure(show_path: Path) -> Plan:
    """Plan grouping loose SxxEyy items into Season N folders, flattening
    episode subfolders, normalizing season folder names, and merging
    duplicate season folders (e.g. both "S02" and "Season 2" exist).

    Op ordering never depends on a prior rename: content settles into the
    season folder under its EXISTING name; the name-normalizing rename
    happens last.
    """
    show_path = Path(show_path)
    ops: list[Op] = []
    try:
        items = sorted(os.scandir(show_path), key=lambda e: e.name)
    except OSError:
        return Plan()

    folders = [e.name for e in items if e.is_dir(follow_symlinks=False)
               and not is_junk_dir(e.name)]
    files = [e.name for e in items if e.is_file(follow_symlinks=False)]

    season_dirs: dict[int, list[str]] = {}   # snum -> existing folder names
    loose_ep_folders: dict[int, list[str]] = {}

    for folder in folders:
        m = SEASON_FOLDER_PATTERN.match(folder.strip())
        if m:
            season_dirs.setdefault(int(m.group(1)), []).append(folder)
            continue
        # Strict SxxEyy only for loose-folder classification: bare NxN
        # matches junk like "WATCH - Show 9x9 - FREE" (old failing test).
        ep = _SXXEYY.search(folder)
        if ep:
            loose_ep_folders.setdefault(int(ep.group(1)), []).append(folder)
            continue
        sl = SEASON_LIKE_PATTERN.match(folder)
        if sl:
            season_dirs.setdefault(int(sl.group(2)), []).append(folder)

    loose_ep_files: dict[int, list[str]] = {}
    for f in files:
        # Companions (subs/nfo) carry the same SxxEyy code but often not the
        # full video stem (no quality tag), so match by episode code.
        if is_media_file(f, VIDEO_EXTS | COMPANION_EXTS):
            s, e = extract_season_episode(f)
            if s is not None:
                loose_ep_files.setdefault(s, []).append(f)

    def canonical(snum: int) -> str:
        """Existing folder that season content should land in (pre-rename).

        A year-tagged "Season N (YYYY)" is a deliberate rename-scheme output
        and counts as already canonical — organize must not strip the year.
        """
        existing = season_dirs.get(snum, [])
        target = f"Season {snum}"
        if target in existing:
            return target
        for name in existing:
            if _CANONICAL_SEASON.match(name):
                return name
        return existing[0] if existing else target

    # 1. Merge duplicate season folders into the canonical one FIRST — so
    #    that the flatten step below sees the merged content.
    for snum, names in sorted(season_dirs.items()):
        canon = canonical(snum)
        for extra in names:
            if extra == canon:
                continue
            extra_path = show_path / extra
            try:
                children = sorted(e.name for e in os.scandir(extra_path))
            except OSError:
                continue
            for child in children:
                ops.append(Op("move", extra_path / child,
                              show_path / canon / child))
            ops.append(Op("rmdir", None, extra_path))

    # 2. Flatten everything nested inside season folders, at any depth. A
    #    season folder should hold episodes, not a directory tree: per-episode
    #    folders, "Disc 1", and a "Subs" folder inside one of those all get
    #    lifted out. Previously only episode-named folders exactly one level
    #    down were flattened, so "Season 1/Disc 1/ep.mkv" stayed buried.
    for snum, names in sorted(season_dirs.items()):
        for season_name in names:
            season_path = show_path / season_name
            ops.extend(_lift_media(season_path, lambda _n, t=season_path: t,
                                   include_root=False))

    # 3. Deal with loose episode folders. When one actually holds the episode,
    #    put the FILES into the season folder rather than nesting the folder
    #    inside it — otherwise the episode ends up at
    #    "Season 1/Show.S01E01.1080p/Show.S01E01.1080p.mkv", which step 2
    #    cannot fix because it planned against the pre-move layout.
    #    A folder with nothing recognisable inside is moved wholesale instead,
    #    so we never discard something we did not understand.
    for snum, ep_folders in sorted(loose_ep_folders.items()):
        target = show_path / canonical(snum)
        for ep_folder in sorted(ep_folders):
            ep_path = show_path / ep_folder
            lifted = _lift_media(ep_path, lambda _n, t=target: t,
                                 include_root=True)
            if any(o.kind == "move" for o in lifted):
                ops.extend(lifted)
            else:
                ops.append(Op("move", ep_path, target / ep_folder))

    # 3b. Any other subfolder that turns out to hold episodes — a "Downloads"
    #     dump, a "Complete Series" wrapper, a disc rip — gets its episodes
    #     routed into the right season folder, per file, since one folder can
    #     hold several seasons. Nothing else sees these: they are not inside a
    #     season folder (so step 2 misses them) and their name carries no
    #     SxxEyy (so step 3 misses them). Files with no episode code are left
    #     exactly where they are rather than guessed at.
    handled = {name for names in season_dirs.values() for name in names}
    handled |= {name for names in loose_ep_folders.values() for name in names}

    def route_by_season(filename: str):
        snum, _ep = extract_season_episode(filename)
        return None if snum is None else show_path / canonical(snum)

    for folder in folders:
        if folder in handled or SPECIALS_FOLDER_PATTERN.match(folder.strip()):
            continue
        lifted = _lift_media(show_path / folder, route_by_season,
                             include_root=True)
        # Only act when there is an episode to lift. Otherwise this is a
        # folder we have no business touching, and the trailing rmdirs would
        # be us tidying away empty directories nobody asked us to remove.
        if any(o.kind == "move" for o in lifted):
            ops.extend(lifted)

    # 4. Move loose episode files (videos + same-code companions) into the
    #    season folder.
    for snum, filenames in sorted(loose_ep_files.items()):
        target = show_path / canonical(snum)
        for filename in sorted(filenames):
            ops.append(Op("move", show_path / filename, target / filename))

    # 5. Rename season folders to "Season N" — last, so nothing above
    #    depends on the new name. "Season N (YYYY)" is left as-is.
    for snum, names in sorted(season_dirs.items()):
        canon = canonical(snum)
        if not _CANONICAL_SEASON.match(canon):
            ops.append(Op("move", show_path / canon,
                          show_path / f"Season {snum}"))

    return check_collisions(ops)


# --- Loose movie files planner ----------------------------------------------

def plan_loose_movies(movies_root: Path) -> Plan:
    """Plan moving loose video files in the movies root into own folders."""
    movies_root = Path(movies_root)
    ops: list[Op] = []
    try:
        entries = sorted(os.scandir(movies_root), key=lambda e: e.name)
    except OSError:
        return Plan()
    for entry in entries:
        if not entry.is_file(follow_symlinks=False):
            continue
        p = Path(entry.path)
        if not is_media_file(p.name):
            continue
        folder = movies_root / p.stem
        if not p.stem or '..' in p.stem or '/' in p.stem or '\\' in p.stem:
            continue  # reject empty or path-traversal stems
        if not folder.exists():
            ops.append(Op("mkdir", None, folder))
        ops.append(Op("move", p, folder / p.name))
        for comp in companion_files(p):
            ops.append(Op("move", comp, folder / comp.name))
    return check_collisions(ops)


# --- Naming scheme & name builders ------------------------------------------

class NamingScheme:
    def __init__(self):
        self.movie_folder_include_year = True
        self.movie_folder_include_quality = True
        self.movie_folder_include_size = False
        self.movie_file_include_year = True
        self.movie_file_include_quality = True
        self.movie_file_include_size = False
        # On by default: without them, two cuts or two discs of the same film
        # build identical names and both get skipped as a conflict.
        self.include_edition = True
        self.include_part = True
        self.tv_parent_include_year = True
        self.tv_parent_include_quality = False
        self.tv_season_include_year = True
        self.tv_episode_include_year = False
        self.tv_episode_include_quality = True
        self.tv_episode_include_size = False

    def to_dict(self):
        return dict(self.__dict__)

    @classmethod
    def from_dict(cls, data):
        scheme = cls()
        for k, v in (data or {}).items():
            if hasattr(scheme, k):
                setattr(scheme, k, v)
        return scheme


_FS_UNSAFE = re.compile(r'[:<>"/\\|?*\x00-\x1f\x7f]')
# Zero-width and bidi-override characters: invisible, and they make two names
# that look identical compare unequal.
_INVISIBLE = re.compile(r'[​-‏‪-‮⁠﻿]')
# Windows refuses these as filenames with or without an extension.
_WIN_RESERVED = re.compile(r'^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$', re.IGNORECASE)
# ext4/APFS cap a path component at 255 *bytes*, so a CJK title can bust the
# limit while Python still sees only 200 characters.
MAX_COMPONENT_BYTES = 255
FALLBACK_NAME = "Untitled"


def _clamp_bytes(name: str, limit: int = MAX_COMPONENT_BYTES) -> str:
    if len(name.encode('utf-8')) <= limit:
        return name
    # Trim whole characters so a multi-byte codepoint is never split.
    while name and len(name.encode('utf-8')) > limit:
        name = name[:-1]
    return name.rstrip(' .')


def sanitize(name: str, *, fallback: str = FALLBACK_NAME) -> str:
    """Make a name safe on all target filesystems (Windows is strictest).

    Applied on every platform so a library organised on Linux stays portable.
    Never returns an empty string: doing so turned a movie file into a bare
    ".mkv" (a hidden file with no stem) whenever the title was stripped away.
    """
    s = _INVISIBLE.sub('', _FS_UNSAFE.sub('', norm(name)))
    # Leading dots hide the file on Unix; trailing dots/spaces are illegal on
    # Windows and get silently dropped by the API otherwise.
    s = s.strip().strip('.').strip()
    head, sep, tail = s.partition('.')
    if _WIN_RESERVED.match(head):
        s = f"{head}_{sep}{tail}"
    s = _clamp_bytes(s)
    return s or fallback


def episode_code(p: ParsedName) -> str:
    """"S01E05", "S01E01-E02", or "2024-01-15" for date-based episodes."""
    if p.date:
        return p.date.isoformat()
    season = p.season if p.season is not None else 1
    if not p.episodes:
        return f"S{season:02d}"
    code = f"S{season:02d}E{p.episodes[0]:02d}"
    for e in p.episodes[1:]:
        code += f"-E{e:02d}"
    return code


def _distinguishers(p: ParsedName, scheme: NamingScheme) -> list[str]:
    """Bits that keep two cuts/discs of the same title from colliding."""
    parts = []
    if scheme.include_edition and p.edition:
        parts.append(f"{{{p.edition}}}")
    if scheme.include_part and p.part:
        parts.append(f"pt{p.part}")
    return parts


def build_movie_folder_name(p: ParsedName, size_gb, scheme: NamingScheme) -> str:
    parts = [p.title]
    if scheme.movie_folder_include_year and p.year:
        parts.append(f"({p.year})")
    parts += _distinguishers(p, scheme)
    if scheme.movie_folder_include_quality and p.quality:
        parts.append(f"[{p.quality}]")
    if scheme.movie_folder_include_size and size_gb:
        parts.append(f"[{size_gb}GB]")
    return sanitize(" ".join(parts))


def build_movie_file_name(p: ParsedName, ext: str, size_gb, scheme: NamingScheme) -> str:
    parts = [p.title]
    if scheme.movie_file_include_year and p.year:
        parts.append(f"({p.year})")
    parts += _distinguishers(p, scheme)
    if scheme.movie_file_include_quality and p.quality:
        parts.append(f"[{p.quality}]")
    if scheme.movie_file_include_size and size_gb:
        parts.append(f"[{size_gb}GB]")
    return sanitize(" ".join(parts)) + ext


def build_tv_show_folder_name(p: ParsedName, scheme: NamingScheme) -> str:
    parts = [p.title]
    if scheme.tv_parent_include_year and p.year:
        parts.append(f"({p.year})")
    if scheme.tv_parent_include_quality and p.quality:
        parts.append(f"[{p.quality}]")
    return sanitize(" ".join(parts))


def build_season_folder_name(season_num: int, season_year, scheme: NamingScheme) -> str:
    name = f"Season {season_num}"
    if scheme.tv_season_include_year and season_year:
        name += f" ({season_year})"
    return sanitize(name, fallback=f"Season {season_num}")


def build_episode_file_name(p: ParsedName, ext: str, size_gb, scheme: NamingScheme) -> str:
    base = f"{p.title} {episode_code(p)}"
    if p.episode_title:
        base += f" - {p.episode_title}"
    parts = [base]
    parts += _distinguishers(p, scheme)
    if scheme.tv_episode_include_year and p.year:
        parts.append(f"({p.year})")
    if scheme.tv_episode_include_quality and p.quality:
        parts.append(f"[{p.quality}]")
    if scheme.tv_episode_include_size and size_gb:
        parts.append(f"[{size_gb}GB]")
    # A fallback of just the episode code keeps a title-less parse from
    # producing a bare extension, and stays unique per episode.
    return sanitize(" ".join(parts), fallback=episode_code(p)) + ext
