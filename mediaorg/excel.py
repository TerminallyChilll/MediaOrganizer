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

from .parse import (COMPANION_EXTS, VIDEO_EXTS, ParsedName, companion_tail,
                    is_junk_name, parse_name)
from .plan import (EPISODE_PATTERN, SEASON_FOLDER_PATTERN, NamingScheme, Op,
                   Plan, build_episode_file_name, build_movie_file_name,
                   build_movie_folder_name, build_season_folder_name,
                   build_tv_show_folder_name, check_collisions,
                   extract_season_episode, norm, sanitize)

_SEASON_FOLDER_RE = re.compile(r'^Season\s+\d+$', re.IGNORECASE)


def _differs(new: str, old: str) -> bool:
    """Does `new` actually differ from `old`?

    Unicode-normalising both sides is what makes renames converge: macOS hands
    back NFD from the filesystem while our builders emit NFC, so a raw compare
    flagged every accented title as needing a rename on every single run.
    """
    return norm(new) != norm(old)


def _existing_season_dir(show_path: Path, season_num: int) -> str | None:
    """The season folder on disk for this season, whatever it is called.

    Looking for the literal "Season 1" meant a zero-padded "Season 01" could
    never be normalised — organize skipped it and the renamer never found it.
    """
    target = f"Season {season_num}"
    try:
        names = sorted(e.name for e in os.scandir(show_path)
                       if e.is_dir(follow_symlinks=False))
    except OSError:
        return None
    if target in names:
        return target
    for name in names:
        m = SEASON_FOLDER_PATTERN.match(name.strip())
        if m and int(m.group(1)) == season_num:
            return name
    return None


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

    The prefix test is `parse.companion_tail`, shared with the review screen
    that has to recognise these destinations later. Its boundary check matters
    here too: a bare `startswith` matches "Episode10.en" against "Episode1", so
    renaming Episode1 would rename Episode 10's subtitle on top of it.
    """
    old_stem, new_stem = Path(old_base).stem, Path(new_base).stem
    code = extract_season_episode(old_base)
    ops: list[Op] = []
    try:
        entries = list(os.scandir(video_path.parent))
    except OSError:
        return ops
    for e in entries:
        if not e.is_file(follow_symlinks=False) or is_junk_name(e.name):
            continue
        p = Path(e.path)
        if p.suffix.lower() not in COMPANION_EXTS:
            continue
        # Normalise before the prefix test, or a subtitle stored NFD next to an
        # NFC video is orphaned instead of being renamed alongside it.
        normed_stem, normed_old = norm(p.stem), norm(old_stem)
        prefix_tail = companion_tail(normed_stem, normed_old)
        if prefix_tail is not None:
            tail = prefix_tail
        elif code != (None, None) and extract_season_episode(p.name) == code:
            # Keep whatever follows the episode code (e.g. ".en" language tag).
            m = EPISODE_PATTERN.search(p.stem)
            tail = p.stem[m.end():] if m else ''
        else:
            continue
        new_name = new_stem + tail + p.suffix
        if _differs(new_name, p.name):
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
    # The prompt asks for quality and _parse_llm_response returns it, but it
    # used to be dropped here — so on a name guessit could not read at all,
    # the LLM's resolution was thrown away and the new name had no [1080p].
    if llm_r.get('quality'):
        p.quality = str(llm_r['quality'])
    p.source = "llm"
    return p


def _video_files(row) -> list[str]:
    """The names listed in one row's "Video Files" cell.

    The scanner joins with " | ". The legacy scanner used a comma, which is
    still accepted — but only when every part looks like a media filename.
    Splitting unconditionally on a comma tore a lone
    "Crouching Tiger, Hidden Dragon.mkv" into two phantom sources; both were
    then skipped as missing and the file was never renamed at all.
    """
    raw = str(row.get('Video Files') or '').strip()
    if not raw:
        return []
    if '|' in raw:
        return [v.strip() for v in raw.split('|') if v.strip()]
    parts = [v.strip() for v in raw.split(',') if v.strip()]
    if len(parts) > 1 and all(Path(p).suffix.lower() in VIDEO_EXTS
                              for p in parts):
        return parts
    return [raw]


def _plan_movie_files(row, folder_path: Path, scheme: NamingScheme,
                      llm_results: dict, custom_patterns, ops: list[Op],
                      folder_parse: ParsedName | None = None) -> None:
    """Plan the renames for the video files listed on one Movies row."""
    for vf in _video_files(row):
        ext = Path(vf).suffix
        pf = parse_name(vf, kind_hint="movie", custom_patterns=custom_patterns)
        if vf in llm_results:
            pf = _parsed_from_llm(llm_results[vf], pf)
        if get_val(row, 'Title Fixed', ''):
            pf.title = get_val(row, 'Title Fixed', '')
        # Honor Fixed-column overrides on file-level parses too, so Year
        # Fixed / Quality Fixed apply uniformly to folder and file renames.
        if get_val(row, 'Year Fixed', ''):
            pf.year = _safe_int_year(get_val(row, 'Year Fixed', ''))
        if get_val(row, 'Quality Fixed', 'Quality'):
            pf.quality = get_val(row, 'Quality Fixed', 'Quality')
        if folder_parse is not None:
            pf.year = pf.year or folder_parse.year
            pf.quality = pf.quality or folder_parse.quality
        new_file = build_movie_file_name(
            pf, ext,
            _clean(row.get('Size (GB)')) if scheme.movie_file_include_size else None,
            scheme)
        # Guard against a stem-less name (".mkv"): sanitize now always
        # returns a fallback, but the check keeps the invariant local.
        if Path(new_file).stem and _differs(new_file, vf):
            ops.extend(_companion_ops(folder_path / vf, vf, new_file))
            ops.append(Op("move", folder_path / vf, folder_path / new_file))


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
            if old_folder == '.':
                # A recursive scan records files sitting directly in the
                # movies root under '.'. Their own renames are handled from
                # the Video Files column below, but treating '.' as a folder
                # to rename means planning a move of the library root into a
                # subdirectory of itself — it fails with EINVAL, and until it
                # does it sits in the confirmation preview looking terrifying.
                _plan_movie_files(row, movies_path, scheme, llm_results,
                                  custom_patterns, ops)
                continue
            folder_path = movies_path / old_folder

            # These rows are known to be movies, so say so. Left to guess,
            # guessit reads the trailing number of a title like "Blade Runner
            # 2049 (2017)" — which is our own output — as an episode number,
            # and the second run renames the file to "Blade Runner (2017)".
            # "300 (2006) [1080p].mkv" grew a new "() [1080p]" every run.
            p = parse_name(old_folder, kind_hint="movie",
                           custom_patterns=custom_patterns)
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
            _plan_movie_files(row, folder_path, scheme, llm_results,
                              custom_patterns, ops, folder_parse=p)

            folder_fixed = _clean(row.get('Folder Fixed'))
            # Spreadsheet overrides were previously taken verbatim, so a user
            # typing "Movie: Part 2" produced a destination with a raw colon.
            new_folder = sanitize(folder_fixed) if folder_fixed else \
                build_movie_folder_name(
                    p, _clean(row.get('Size (GB)')) if scheme.movie_folder_include_size else None,
                    scheme)
            if new_folder and _differs(new_folder, Path(old_folder).name):
                # Preserve the parent directory when Folder Name contains a
                # nested path (e.g. "Collection/Movie.2020" from a recursive
                # scan) so the rename stays in-place.
                parent_dir = Path(old_folder).parent
                if parent_dir != Path('.') and parent_dir != Path(''):
                    dest = movies_path / parent_dir / new_folder
                else:
                    dest = movies_path / new_folder
                ops.append(Op("move", folder_path, dest))

    if df_tv is not None and tv_path:
        tv_path = Path(tv_path)
        for show_folder, show_eps in df_tv.groupby('Show Folder'):
            show_folder = str(show_folder)
            first = show_eps.iloc[0]
            # A nested show is recorded as a relative path ("Genre/Show"), so
            # every name-level decision below has to look at the last
            # component. Matching against the whole path meant a nested show
            # was never recognised as a season-only scan, and its title was
            # parsed out of the wrapper directories leading to it.
            show_name = Path(show_folder).name
            # Scanner run on a single show root: "show folders" are seasons.
            show_is_season = bool(_SEASON_FOLDER_RE.match(show_name))
            # Recursive scans record root-level episodes as '.': the show is
            # the selected tv_path itself.
            show_is_root = show_folder == '.'
            show_path = tv_path if show_is_root else tv_path / show_folder

            p_show = parse_name(Path(tv_path).name if show_is_root else show_name,
                                custom_patterns=custom_patterns)
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
            for season_num, season_eps in show_eps.groupby('Season', dropna=False):
                if _clean(season_num) == '':
                    # No detectable season: an extra, a trailer, or a movie
                    # filed under TV. Leave it alone rather than inventing a
                    # name for it.
                    continue
                try:
                    season_num = int(float(season_num))
                except (TypeError, ValueError):
                    continue
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
                    if not pe.episodes and pe.date is None:
                        # No episode number and no air date (a special, an
                        # extra). Naming it "Show S00" would be invented, and
                        # two such files would collide and both be skipped —
                        # leave it as the user has it.
                        continue
                    if get_val(episode, 'Quality Fixed', ''):
                        pe.quality = get_val(episode, 'Quality Fixed', '')
                    pe.year = _safe_int_year(season_year) if season_year else pe.year

                    file_fixed = _clean(episode.get('File Fixed'))
                    if file_fixed:
                        # Route the override through sanitize too. Only strip a
                        # RECOGNISED media extension: Path.suffix on
                        # "Show S01E01 - Mr. Robot" is ". Robot", so a blanket
                        # strip truncated dotted titles to "Show S01E01 - Mr".
                        suffix = Path(file_fixed).suffix.lower()
                        known_ext = suffix in VIDEO_EXTS or suffix in COMPANION_EXTS
                        stem = Path(file_fixed).stem if known_ext else file_fixed
                        new_base = sanitize(stem)
                    else:
                        new_base = build_episode_file_name(
                            pe, ext,
                            _clean(episode.get('Size (GB)')) if scheme.tv_episode_include_size else None,
                            scheme)
                    if not new_base.lower().endswith(ext.lower()):
                        new_base += ext
                    if Path(new_base).stem and _differs(new_base, base):
                        ops.extend(_companion_ops(old_path, base, new_base))
                        ops.append(Op("move", old_path, old_path.parent / new_base))

                if not show_is_season:
                    old_season = _existing_season_dir(show_path, season_num)
                    new_season = build_season_folder_name(season_num, season_year, scheme)
                    if old_season and _differs(new_season, old_season):
                        season_ops.append(Op("move", show_path / old_season,
                                             show_path / new_season))

            ops.extend(season_ops)

            if not show_is_season and not show_is_root:
                show_fixed = _clean(first.get('Folder Fixed'))
                new_show = sanitize(show_fixed) if show_fixed else \
                    build_tv_show_folder_name(p_show, scheme)
                if new_show and _differs(new_show, Path(show_folder).name):
                    # Preserve parent dir for nested show folders from
                    # recursive TV scans (e.g. "Parent/ShowName").
                    show_parent = Path(show_folder).parent
                    if show_parent != Path('.') and show_parent != Path(''):
                        dest = tv_path / show_parent / new_show
                    else:
                        dest = tv_path / new_show
                    ops.append(Op("move", show_path, dest))

    return check_collisions(ops)
