"""Strictly read-only library scanning. Returns Excel-schema row dicts.

The old scanner silently moved loose files during "scan" — this one never
touches disk. Loose-file organizing is a planned, previewed operation in
the wizard (plan.plan_loose_movies).
"""

import os
import re
from pathlib import Path

from tqdm import tqdm

from .parse import VIDEO_EXTS, parse_name
from .plan import extract_season_episode

_SEASONISH_DIR = re.compile(r'season|s\d{1,2}', re.IGNORECASE)


def folder_size_gb(path: Path) -> float:
    total = 0
    for dirpath, _dirnames, filenames in os.walk(path):
        for f in filenames:
            try:
                total += os.path.getsize(os.path.join(dirpath, f))
            except OSError:
                pass
    return round(total / (1024 ** 3), 2)


def scan_movies(movies_root: Path, custom_patterns: list[str] = ()) -> list[dict]:
    movies_root = Path(movies_root)
    if not movies_root.is_dir():
        return []
    folders = sorted(e.name for e in os.scandir(movies_root)
                     if e.is_dir(follow_symlinks=False))
    rows = []
    for folder_name in tqdm(folders, desc="Scanning Movies", unit="folder"):
        folder_path = movies_root / folder_name
        p = parse_name(folder_name, custom_patterns=custom_patterns)
        try:
            video_files = sorted(
                e.name for e in os.scandir(folder_path)
                if e.is_file(follow_symlinks=False)
                and Path(e.name).suffix.lower() in VIDEO_EXTS)
        except OSError:
            video_files = []
        # Fill year/quality from the video files when the folder name lacks them.
        year, quality = p.year, p.quality
        for vf in video_files:
            if year and quality:
                break
            pf = parse_name(vf, custom_patterns=custom_patterns)
            year = year or pf.year
            quality = quality or pf.quality
        rows.append({
            'Folder Name': folder_name, 'Folder Fixed': '',
            'Title': p.title, 'Title Fixed': '',
            'Year': year or '', 'Year Fixed': '',
            'Quality': quality or '', 'Quality Fixed': '',
            'Size (GB)': folder_size_gb(folder_path),
            'Video Files': ' | '.join(video_files), 'Files Fixed': '',
        })
    return rows


def _scan_show_episodes(show_path: Path) -> list[dict]:
    """Episodes of one show: loose files in the root plus everything under
    season-ish folders (any depth). rel_path is relative to show_path."""
    episodes = []
    try:
        entries = sorted(os.scandir(show_path), key=lambda e: e.name)
    except OSError:
        return episodes

    season_dirs = [e.name for e in entries
                   if e.is_dir(follow_symlinks=False) and _SEASONISH_DIR.search(e.name)]

    for e in entries:
        if e.is_file(follow_symlinks=False) and Path(e.name).suffix.lower() in VIDEO_EXTS:
            s, ep = extract_season_episode(e.name)
            if s is not None:
                episodes.append({'season': s, 'episode': ep, 'rel_path': e.name,
                                 'size': round(e.stat().st_size / 1024 ** 3, 2)})

    for season_dir in season_dirs:
        for dirpath, _dirnames, filenames in os.walk(show_path / season_dir):
            for f in sorted(filenames):
                if Path(f).suffix.lower() not in VIDEO_EXTS:
                    continue
                s, ep = extract_season_episode(f)
                if s is None:
                    continue
                full = Path(dirpath) / f
                episodes.append({
                    'season': s, 'episode': ep,
                    'rel_path': str(full.relative_to(show_path)),
                    'size': round(full.stat().st_size / 1024 ** 3, 2)})
    return episodes


def scan_tv(tv_root: Path, custom_patterns: list[str] = ()) -> list[dict]:
    tv_root = Path(tv_root)
    if not tv_root.is_dir():
        return []
    folders = sorted(e.name for e in os.scandir(tv_root)
                     if e.is_dir(follow_symlinks=False))
    rows = []
    for folder_name in tqdm(folders, desc="Scanning TV Shows", unit="show"):
        show_path = tv_root / folder_name
        p = parse_name(folder_name, custom_patterns=custom_patterns)
        episodes = _scan_show_episodes(show_path)
        if not episodes:
            rows.append({
                'Show Folder': folder_name, 'Folder Fixed': '',
                'Title': p.title, 'Title Fixed': '',
                'Season': '', 'Season Year': '', 'Episode': '',
                'Episode File': '', 'File Fixed': '',
                'Quality': '', 'Quality Fixed': '',
                'Size (GB)': folder_size_gb(show_path),
            })
            continue
        for ep in episodes:
            pf = parse_name(Path(ep['rel_path']).name, kind_hint="episode",
                            custom_patterns=custom_patterns)
            rows.append({
                'Show Folder': folder_name, 'Folder Fixed': '',
                'Title': p.title, 'Title Fixed': '',
                'Season': ep['season'],
                'Season Year': pf.year or p.year or '',
                'Episode': ep['episode'],
                'Episode File': ep['rel_path'], 'File Fixed': '',
                'Quality': pf.quality or '', 'Quality Fixed': '',
                'Size (GB)': ep['size'],
            })
    return rows


def scan_recursive(root: Path, custom_patterns: list[str] = ()) -> list[dict]:
    """Recursively walk *every* directory under *root*, find all video files,
    and return one row per containing folder (Movies-sheet format).

    Unlike ``scan_movies`` / ``scan_tv`` this makes no assumption about the
    directory layout — it simply walks the whole tree with :func:`os.walk`,
    which makes it safe for deeply nested or irregular structures and for
    SMB/GVFS mounts where only the top-level path is navigable.
    """
    root = Path(root)
    if not root.is_dir():
        return []

    # ── collect video files keyed by their parent folder ──────────────────
    folder_files: dict[Path, list[Path]] = {}
    walk_errors: list[str] = []
    try:
        for dirpath, dirnames, filenames in os.walk(
            root, onerror=lambda err: walk_errors.append(str(err))
        ):
            dirnames.sort()
            vids = sorted(
                f for f in filenames
                if Path(f).suffix.lower() in VIDEO_EXTS)
            if vids:
                folder_files[Path(dirpath)] = [Path(dirpath) / f for f in vids]
    except OSError as exc:
        walk_errors.append(str(exc))

    if not folder_files:
        if walk_errors:
            print(f"   [!] Could not read directory: {'; '.join(walk_errors)}")
        return []

    rows: list[dict] = []
    for folder_path, video_paths in tqdm(
        sorted(folder_files.items()), desc="Scanning recursively", unit="folder"
    ):
        # Human-friendly label: path relative to root (or "." for root itself)
        try:
            folder_name = str(folder_path.relative_to(root))
        except ValueError:
            folder_name = str(folder_path)
        if folder_name == ".":
            folder_name = "."

        p = parse_name(folder_path.name, custom_patterns=custom_patterns)
        video_names = [vp.name for vp in video_paths]

        # Fill year/quality from video files when the folder name lacks them
        year, quality = p.year, p.quality
        for vp in video_paths:
            if year and quality:
                break
            pf = parse_name(vp.name, custom_patterns=custom_patterns)
            year = year or pf.year
            quality = quality or pf.quality

        # Total size of all video files in this folder
        total_gb = 0.0
        for vp in video_paths:
            try:
                total_gb += vp.stat().st_size / (1024 ** 3)
            except OSError:
                pass

        rows.append({
            'Folder Name': folder_name, 'Folder Fixed': '',
            'Title': p.title, 'Title Fixed': '',
            'Year': year or '', 'Year Fixed': '',
            'Quality': quality or '', 'Quality Fixed': '',
            'Size (GB)': round(total_gb, 2),
            'Video Files': ' | '.join(video_names), 'Files Fixed': '',
        })

    return rows


def scan_recursive_tv(root: Path, custom_patterns: list[str] = ()) -> list[dict]:
    """Like :func:`scan_recursive` but returns TV-schema rows (one per
    episode file) suitable for writing to the ``TV Shows`` sheet.

    Every video file found under *root* becomes one row; its parent directory
    becomes the ``Show Folder``.  Season / episode numbers are extracted from
    the filename when possible, otherwise default to season 0 / episode 0 so
    the row is still visible in the spreadsheet.
    """
    root = Path(root)
    if not root.is_dir():
        return []

    # ── walk the whole tree and collect every video file ──────────────────
    all_files: list[tuple[Path, Path]] = []  # (show_path, video_path)
    walk_errors: list[str] = []
    try:
        for dirpath, dirnames, filenames in os.walk(
            root, onerror=lambda err: walk_errors.append(str(err))
        ):
            dirnames.sort()
            for f in sorted(filenames):
                if Path(f).suffix.lower() in VIDEO_EXTS:
                    all_files.append((Path(dirpath), Path(dirpath) / f))
    except OSError as exc:
        walk_errors.append(str(exc))

    if not all_files:
        if walk_errors:
            print(f"   [!] Could not read directory: {'; '.join(walk_errors)}")
        return []

    rows: list[dict] = []
    for show_path, video_path in tqdm(
        sorted(all_files, key=lambda x: (str(x[0]), x[1].name)),
        desc="Scanning TV recursively", unit="file"
    ):
        try:
            show_folder = str(show_path.relative_to(root))
        except ValueError:
            show_folder = str(show_path)
        if show_folder == ".":
            show_folder = "."

        # Parse show title from the folder name
        p_show = parse_name(show_path.name, custom_patterns=custom_patterns)
        # Parse season/episode/quality from the filename
        pf = parse_name(video_path.name, kind_hint="episode",
                        custom_patterns=custom_patterns)
        s, ep = extract_season_episode(video_path.name)

        try:
            size_gb = round(video_path.stat().st_size / (1024 ** 3), 2)
        except OSError:
            size_gb = 0.0

        rows.append({
            'Show Folder': show_folder, 'Folder Fixed': '',
            'Title': p_show.title, 'Title Fixed': '',
            'Season': s if s is not None else '',
            'Season Year': pf.year or p_show.year or '',
            'Episode': ep if ep is not None else '',
            'Episode File': str(video_path.relative_to(show_path)),
            'File Fixed': '',
            'Quality': pf.quality or '', 'Quality Fixed': '',
            'Size (GB)': size_gb,
        })

    return rows
