"""xlsx journal: write scans, read back with human-edited Fixed overrides,
plan renames from the sheet, log executed changes.

Sheet/column schema is unchanged from v1 so existing spreadsheets keep
working: Movies, TV Shows, Metadata, Changes; "… Fixed" override columns.
"""

import os
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

from .parse import COMPANION_EXTS, ParsedName, parse_name
from .plan import (EPISODE_PATTERN, NamingScheme, Op, Plan,
                   build_episode_file_name, build_movie_file_name,
                   build_movie_folder_name, build_season_folder_name,
                   build_tv_show_folder_name, check_collisions,
                   extract_season_episode)

_SEASON_FOLDER_RE = re.compile(r'^Season\s+\d+$', re.IGNORECASE)


def _clean(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)) or pd.isna(val):
        return ''
    s = str(val).strip()
    return s[:-2] if s.endswith('.0') else s


def get_val(row, fixed_col: str, auto_col: str) -> str:
    fixed = _clean(row.get(fixed_col))
    return fixed if fixed else _clean(row.get(auto_col))


def read_library(path: Path):
    """Returns (df_movies | None, df_tv | None, metadata dict)."""
    excel = pd.ExcelFile(path)
    df_movies = pd.read_excel(path, sheet_name='Movies') if 'Movies' in excel.sheet_names else None
    df_tv = pd.read_excel(path, sheet_name='TV Shows') if 'TV Shows' in excel.sheet_names else None
    meta = {}
    if 'Metadata' in excel.sheet_names:
        df_meta = pd.read_excel(path, sheet_name='Metadata')
        meta = dict(zip(df_meta['Key'].astype(str), df_meta['Value'].astype(str)))
    return df_movies, df_tv, meta


def write_library(path: Path, movies_rows: list[dict], tv_rows: list[dict],
                  movies_path=None, tv_path=None, append: bool = False) -> None:
    path = Path(path)
    movies_df = pd.DataFrame(movies_rows)
    tv_df = pd.DataFrame(tv_rows)
    meta = {}

    if append and path.exists():
        old_movies, old_tv, meta = read_library(path)
        if old_movies is not None and not movies_df.empty:
            movies_df = pd.concat([old_movies, movies_df]).drop_duplicates(
                subset=['Folder Name'], keep='last')
        elif old_movies is not None:
            movies_df = old_movies
        if old_tv is not None and not tv_df.empty:
            tv_df = pd.concat([old_tv, tv_df]).drop_duplicates(
                subset=['Show Folder', 'Season', 'Episode'], keep='last')
        elif old_tv is not None:
            tv_df = old_tv

    if movies_path:
        meta['Movies Path'] = str(movies_path)
    if tv_path:
        meta['TV Shows Path'] = str(tv_path)
    meta['Last Scan'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    meta_df = pd.DataFrame(list(meta.items()), columns=['Key', 'Value'])

    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        if not movies_df.empty:
            movies_df.to_excel(writer, sheet_name='Movies', index=False)
        if not tv_df.empty:
            tv_df.to_excel(writer, sheet_name='TV Shows', index=False)
        meta_df.to_excel(writer, sheet_name='Metadata', index=False)
        _autosize_columns(writer)


def _autosize_columns(writer) -> None:
    for sheetname in writer.sheets:
        worksheet = writer.sheets[sheetname]
        for col in worksheet.columns:
            width = max((len(str(c.value)) for c in col if c.value is not None),
                        default=0)
            worksheet.column_dimensions[col[0].column_letter].width = min(width + 2, 60)


def append_changes(path: Path, entries: list[dict]) -> None:
    """Log executed journal entries to the Changes sheet (actual paths)."""
    if not entries:
        return
    df_new = pd.DataFrame([{
        'Timestamp': datetime.fromtimestamp(e['ts']).strftime("%Y-%m-%d %H:%M:%S"),
        'Operation': e['op'],
        'From': e.get('src') or '',
        'To': e['dst'],
    } for e in entries])
    try:
        old = pd.read_excel(path, sheet_name='Changes')
        df_new = pd.concat([old, df_new])
    except (FileNotFoundError, ValueError):
        pass
    with pd.ExcelWriter(path, engine='openpyxl', mode='a',
                        if_sheet_exists='replace') as writer:
        df_new.to_excel(writer, sheet_name='Changes', index=False)


def _companion_ops(video_path: Path, old_base: str, new_base: str) -> list[Op]:
    """Rename subs/nfo alongside their video so players keep matching them.

    A companion belongs to the video when its stem starts with the video's
    old stem, or (for episodes) it carries the same SxxEyy code. Extra stem
    tail like a ".en" language tag is preserved.
    """
    old_stem, new_stem = Path(old_base).stem, Path(new_base).stem
    code = extract_season_episode(old_base)
    ops: list[Op] = []
    try:
        entries = list(os.scandir(video_path.parent))
    except OSError:
        return ops
    for e in entries:
        if not e.is_file(follow_symlinks=False):
            continue
        p = Path(e.path)
        if p.suffix.lower() not in COMPANION_EXTS:
            continue
        if p.stem.startswith(old_stem):
            tail = p.stem[len(old_stem):]
        elif code != (None, None) and extract_season_episode(p.name) == code:
            # Keep whatever follows the episode code (e.g. ".en" language tag).
            m = EPISODE_PATTERN.search(p.stem)
            tail = p.stem[m.end():] if m else ''
        else:
            continue
        new_name = new_stem + tail + p.suffix
        if new_name != p.name:
            ops.append(Op("move", p, p.with_name(new_name)))
    return ops


# --- Rename planning from the sheet -----------------------------------------

def _safe_int_year(val) -> int | None:
    """Convert a year value to int, returning None on failure."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    try:
        return int(float(str(val).strip()))
    except (ValueError, TypeError):
        return None


def _parsed_from_llm(llm_r: dict, fallback: ParsedName) -> ParsedName:
    p = ParsedName(**{**fallback.__dict__})
    p.title = llm_r.get('title') or p.title
    if llm_r.get('year'):
        p.year = _safe_int_year(llm_r['year'])
    p.source = "llm"
    return p


def plan_renames(df_movies, movies_path, df_tv, tv_path, scheme: NamingScheme,
                 llm_results: dict | None = None,
                 custom_patterns: list[str] = ()) -> Plan:
    """Build rename ops from the spreadsheet (with Fixed-column overrides).

    Emits pure renames only (dst.parent == src.parent), children before
    parents, so no op ever depends on an earlier rename. Cross-directory
    moves are the organizer's job.
    """
    llm_results = llm_results or {}
    ops: list[Op] = []

    if df_movies is not None and movies_path:
        movies_path = Path(movies_path)
        for _, row in df_movies.iterrows():
            old_folder = _clean(row.get('Folder Name'))
            if not old_folder:
                continue
            folder_path = movies_path / old_folder

            p = parse_name(old_folder, custom_patterns=custom_patterns)
            if old_folder in llm_results:
                p = _parsed_from_llm(llm_results[old_folder], p)
            if get_val(row, 'Title Fixed', ''):
                p.title = get_val(row, 'Title Fixed', '')
            if get_val(row, 'Year Fixed', ''):
                p.year = _safe_int_year(get_val(row, 'Year Fixed', ''))
            elif not p.year and _clean(row.get('Year')):
                p.year = _safe_int_year(_clean(row.get('Year')))
            quality = get_val(row, 'Quality Fixed', 'Quality')
            if quality:
                p.quality = quality

            # Files first (children before parent).
            for vf in str(row.get('Video Files') or '').split('|'):
                vf = vf.strip()
                if not vf:
                    continue
                ext = Path(vf).suffix
                pf = parse_name(vf, custom_patterns=custom_patterns)
                if vf in llm_results:
                    pf = _parsed_from_llm(llm_results[vf], pf)
                if get_val(row, 'Title Fixed', ''):
                    pf.title = get_val(row, 'Title Fixed', '')
                pf.year = pf.year or p.year
                pf.quality = pf.quality or p.quality
                new_file = build_movie_file_name(pf, ext, _clean(row.get('Size (GB)')) if scheme.movie_file_include_size else None, scheme)
                if new_file != vf:
                    ops.extend(_companion_ops(folder_path / vf, vf, new_file))
                    ops.append(Op("move", folder_path / vf, folder_path / new_file))

            new_folder = _clean(row.get('Folder Fixed')) or build_movie_folder_name(
                p, _clean(row.get('Size (GB)')) if scheme.movie_folder_include_size else None, scheme)
            if new_folder and new_folder != old_folder:
                ops.append(Op("move", folder_path, movies_path / new_folder))

    if df_tv is not None and tv_path:
        tv_path = Path(tv_path)
        for show_folder, show_eps in df_tv.groupby('Show Folder'):
            show_folder = str(show_folder)
            first = show_eps.iloc[0]
            # Scanner run on a single show root: "show folders" are seasons.
            show_is_season = bool(_SEASON_FOLDER_RE.match(show_folder))
            show_path = tv_path / show_folder

            p_show = parse_name(show_folder, custom_patterns=custom_patterns)
            if show_folder in llm_results:
                p_show = _parsed_from_llm(llm_results[show_folder], p_show)
            if get_val(first, 'Title Fixed', ''):
                p_show.title = get_val(first, 'Title Fixed', '')
            if not p_show.title:
                continue
            # Show year: season 1's year, like v1.
            s1 = show_eps[show_eps['Season'].astype(str).str.strip().isin(['1', '1.0'])]
            s1_year = _clean(s1.iloc[0].get('Season Year')) if len(s1) else ''
            if s1_year:
                p_show.year = _safe_int_year(s1_year)

            season_ops: list[Op] = []
            for season_num, season_eps in show_eps.groupby('Season'):
                if _clean(season_num) == '':
                    continue
                season_num = int(float(season_num))
                season_year = _clean(season_eps.iloc[0].get('Season Year'))

                for _, episode in season_eps.iterrows():
                    rel = _clean(episode.get('Episode File'))
                    if not rel:
                        continue
                    old_path = show_path / rel
                    base = Path(rel).name
                    ext = Path(base).suffix

                    pe = parse_name(base, kind_hint="episode",
                                    custom_patterns=custom_patterns)
                    if base in llm_results:
                        pe = _parsed_from_llm(llm_results[base], pe)
                    pe.title = p_show.title if not show_is_season else (pe.title or Path(tv_path).name)
                    s, e = extract_season_episode(base)
                    pe.season = s if s is not None else season_num
                    if not pe.episodes and e is not None:
                        pe.episodes = [e]
                    if get_val(episode, 'Quality Fixed', ''):
                        pe.quality = get_val(episode, 'Quality Fixed', '')
                    pe.year = _safe_int_year(season_year) if season_year else pe.year

                    new_base = _clean(episode.get('File Fixed')) or build_episode_file_name(
                        pe, ext,
                        _clean(episode.get('Size (GB)')) if scheme.tv_episode_include_size else None,
                        scheme)
                    if not new_base.lower().endswith(ext.lower()):
                        new_base += ext
                    if new_base != base:
                        ops.extend(_companion_ops(old_path, base, new_base))
                        ops.append(Op("move", old_path, old_path.parent / new_base))

                if not show_is_season:
                    old_season = f"Season {season_num}"
                    new_season = build_season_folder_name(season_num, season_year, scheme)
                    if new_season != old_season and (show_path / old_season).is_dir():
                        season_ops.append(Op("move", show_path / old_season,
                                             show_path / new_season))

            ops.extend(season_ops)

            if not show_is_season:
                new_show = _clean(first.get('Folder Fixed')) or \
                    build_tv_show_folder_name(p_show, scheme)
                if new_show and new_show != show_folder:
                    ops.append(Op("move", show_path, tv_path / new_show))

    return check_collisions(ops)
