"""Pure planning: inspect disk (read-only), emit operations.

Nothing in this module writes to disk. Every mutation the tool ever makes
is an Op produced here (or in extfix) and executed by mediaorg.execute.
"""

import os
import re
from dataclasses import dataclass, field
from pathlib import Path

from .parse import COMPANION_EXTS, VIDEO_EXTS, ParsedName

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

# SxxEyy (group 1=season) or bare NxN like "1x2"/"9x9" (groups 2, 3).
EPISODE_PATTERN = re.compile(
    r'(?:[Ss](\d{1,2})[Ee]\d{1,2}|(?<!\d)(\d{1,2})[xX](\d{1,2})(?!\d))')
# Pure season folders: "Season 1", "S01", "Season 1 (2024)".
SEASON_FOLDER_PATTERN = re.compile(r'^(?:season\s*|s)(\d{1,2})(?:\s*\(\d{4}\))?$', re.IGNORECASE)
# Abbreviated "ShowTitle S02" form only — no dashes/pipes/dots, optional year.
SEASON_LIKE_PATTERN = re.compile(r'^([^\-\|.]+?)\s+[Ss](\d{1,2})(?:\s*\(?\d{4}\)?)?\s*$')

_SXXEYY = re.compile(r'[Ss](\d{1,2})[Ee](\d{1,2})')
_NXN = re.compile(r'(?<!\d)(\d{1,2})[xX](\d{1,2})(?!\d)')


def extract_season_episode(name: str) -> tuple[int | None, int | None]:
    m = _SXXEYY.search(name) or _NXN.search(name)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


def folder_has_episodes_or_seasons(path: Path) -> bool:
    """Is this folder a TV-show root (contains episodes/seasons)?

    Folder names only count with a strict SxxEyy marker — bare NxN like
    "9x9" appears in junk names ("WATCH - Show 9x9 - FREE") and must not
    classify. NxN still counts for files and inside season folders.
    """
    try:
        items = list(os.scandir(path))
    except OSError:
        return False
    for entry in items:
        name = entry.name
        if entry.is_dir(follow_symlinks=False):
            if (_SXXEYY.search(name)
                    or SEASON_FOLDER_PATTERN.match(name.strip())
                    or SEASON_LIKE_PATTERN.match(name)):
                return True
        elif entry.is_file(follow_symlinks=False):
            if (Path(name).suffix.lower() in VIDEO_EXTS
                    and EPISODE_PATTERN.search(name)):
                return True
    return False


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
            if p.suffix.lower() in COMPANION_EXTS and p.name.startswith(video.stem):
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


def check_collisions(ops: list[Op]) -> Plan:
    """Split ops into safe vs skipped. Never overwrite, never auto-suffix."""
    plan = Plan()
    by_dst: dict[str, list[Op]] = {}
    for op in ops:
        by_dst.setdefault(os.path.normcase(str(op.dst)), []).append(op)

    for op in ops:
        if op.kind == "move":
            group = by_dst[os.path.normcase(str(op.dst))]
            if sum(1 for o in group if o.kind == "move") > 1:
                plan.skipped.append((op, "conflict: multiple sources map to same target"))
                continue
            if op.dst.exists() and not (op.src and _same_file(op.src, op.dst)):
                plan.skipped.append((op, f"conflict: target already exists: {op.dst}"))
                continue
            if op.src and not os.path.lexists(op.src):
                plan.skipped.append((op, f"source missing: {op.src}"))
                continue
        plan.ops.append(op)
    return plan


# --- TV season structure planner --------------------------------------------

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

    folders = [e.name for e in items if e.is_dir(follow_symlinks=False)]
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
        if Path(f).suffix.lower() in (VIDEO_EXTS | COMPANION_EXTS):
            s, e = extract_season_episode(f)
            if s is not None:
                loose_ep_files.setdefault(s, []).append(f)

    def canonical(snum: int) -> str:
        """Existing folder that season content should land in (pre-rename)."""
        existing = season_dirs.get(snum, [])
        target = f"Season {snum}"
        if target in existing:
            return target
        return existing[0] if existing else target

    media_exts = VIDEO_EXTS | COMPANION_EXTS

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

    # 2. Flatten episode-named subfolders inside season folders; rmdir after.
    for snum, names in sorted(season_dirs.items()):
        for season_name in names:
            season_path = show_path / season_name
            try:
                subdirs = sorted(e.name for e in os.scandir(season_path)
                                 if e.is_dir(follow_symlinks=False))
            except OSError:
                continue
            for ep_dir in subdirs:
                if not EPISODE_PATTERN.search(ep_dir):
                    continue
                ep_path = season_path / ep_dir
                try:
                    ep_files = sorted(e.name for e in os.scandir(ep_path)
                                      if e.is_file(follow_symlinks=False))
                except OSError:
                    continue
                moved_any = False
                for mf in ep_files:
                    if Path(mf).suffix.lower() in media_exts:
                        ops.append(Op("move", ep_path / mf, season_path / mf))
                        moved_any = True
                if moved_any:
                    ops.append(Op("rmdir", None, ep_path))

    # 3. Move loose episode folders into the season folder (existing name).
    for snum, ep_folders in sorted(loose_ep_folders.items()):
        target = show_path / canonical(snum)
        for ep_folder in sorted(ep_folders):
            ops.append(Op("move", show_path / ep_folder, target / ep_folder))

    # 4. Move loose episode files (videos + same-code companions) into the
    #    season folder.
    for snum, filenames in sorted(loose_ep_files.items()):
        target = show_path / canonical(snum)
        for filename in sorted(filenames):
            ops.append(Op("move", show_path / filename, target / filename))

    # 5. Rename season folders to "Season N" — last, so nothing above
    #    depends on the new name.
    for snum, names in sorted(season_dirs.items()):
        canon = canonical(snum)
        if canon != f"Season {snum}":
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
        if p.suffix.lower() not in VIDEO_EXTS:
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


_FS_UNSAFE = re.compile(r'[:<>"/\\|?*\x00-\x1f]')


def sanitize(name: str) -> str:
    """Make a name safe on all target filesystems (Windows is strictest)."""
    return _FS_UNSAFE.sub('', name).rstrip(' .').strip()


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


def build_movie_folder_name(p: ParsedName, size_gb, scheme: NamingScheme) -> str:
    parts = [p.title]
    if scheme.movie_folder_include_year and p.year:
        parts.append(f"({p.year})")
    if scheme.movie_folder_include_quality and p.quality:
        parts.append(f"[{p.quality}]")
    if scheme.movie_folder_include_size and size_gb:
        parts.append(f"[{size_gb}GB]")
    return sanitize(" ".join(parts))


def build_movie_file_name(p: ParsedName, ext: str, size_gb, scheme: NamingScheme) -> str:
    parts = [p.title]
    if scheme.movie_file_include_year and p.year:
        parts.append(f"({p.year})")
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
    return name


def build_episode_file_name(p: ParsedName, ext: str, size_gb, scheme: NamingScheme) -> str:
    base = f"{p.title} {episode_code(p)}"
    if p.episode_title:
        base += f" - {p.episode_title}"
    parts = [base]
    if scheme.tv_episode_include_year and p.year:
        parts.append(f"({p.year})")
    if scheme.tv_episode_include_quality and p.quality:
        parts.append(f"[{p.quality}]")
    if scheme.tv_episode_include_size and size_gb:
        parts.append(f"[{size_gb}GB]")
    return sanitize(" ".join(parts)) + ext
