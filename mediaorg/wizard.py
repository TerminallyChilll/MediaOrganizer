"""Interactive wizard and CLI entry point. All console UI lives here.

ASCII markers only ([OK], [!], ->): no emoji, no cp1252 crashes.
"""

import argparse
import getpass
import json
import os
import sys
import uuid
from datetime import datetime
from pathlib import Path

from . import excel, extfix, llm, scan
from .execute import (execute, journal_path, last_run_ops, list_runs,
                      pending_runs, recover, undo_last, undo_last_run,
                      undo_run, undo_session)
from .parse import load_custom_patterns, parse_name
from .plan import (NamingScheme, Op, Plan, find_show_roots,
                   plan_loose_movies, plan_season_structure)

CONFIG_FILE = ".media_renamer_config.json"


class BackNavigation(Exception):
    pass


def prompt_input(message: str, default: str = '') -> str:
    val = input(message).strip()
    if val.lower() in ('back', 'b'):
        raise BackNavigation()
    return val if val else default


def ask_yes_no(message: str, default: bool = True) -> bool:
    val = prompt_input(f"{message} (y/n) [{'y' if default else 'n'}]: ")
    if not val:
        return default
    return val.upper() in ('Y', 'YES')


def _load_config() -> dict:
    try:
        with open(CONFIG_FILE, encoding='utf-8') as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def _save_config(config: dict) -> None:
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2)
    except OSError:
        pass


# --- Folder selection --------------------------------------------------------

def _validate_path(raw: str) -> str | None:
    raw = (raw or '').strip().strip('"').strip("'")
    if raw:
        p = Path(raw)
        if p.is_dir():
            return str(p)
    return None


def browse_for_folder(prompt: str, allow_skip: bool = True) -> str | None:
    while True:
        print(f"\n{prompt}")
        print("-" * 50)
        print("  [1] Open folder picker (GUI dialog)")
        print("  [2] Browse folders in terminal")
        print("  [3] Paste a path manually")
        print("  [4] Skip" if allow_skip else "  [4] Cancel / go back")
        choice = input("\nSelect (1-4): ").strip()

        if choice == '1':
            try:
                import tkinter as tk
                from tkinter import filedialog
                root = tk.Tk()
                root.withdraw()
                root.attributes('-topmost', True)
                folder = filedialog.askdirectory(title=prompt)
                root.destroy()
            except Exception as e:
                print(f"   [!] Could not open folder picker: {e}")
                print("   Falling back to terminal browser...")
                folder = _cli_folder_browser(allow_skip)
            if folder:
                # GUI pickers on Linux can return virtual paths (GVFS, SMB
                # "//server/share" URIs) that don't exist on the filesystem.
                valid = _validate_path(folder)
                if valid:
                    folder = valid
                    print(f"   [OK] Selected: {folder}")
                    return folder
                print(f"   [!] Path does not exist or is not a directory: {folder}")
                print("   This can happen with network shares selected via GUI.")
                print("   Try pasting the mounted path manually (option 3), or")
                print("   use the terminal browser (option 2) to navigate to it.")
                # loop back so user can try another method
        elif choice == '2':
            folder = _cli_folder_browser(allow_skip)
            if folder:
                return folder
        elif choice == '3':
            result = _validate_path(input("\nPaste folder path: "))
            if result:
                print(f"   [OK] Valid: {result}")
                return result
            print("   [!] Invalid path. Try again.")
        elif choice == '4':
            return None


def _cli_folder_browser(allow_skip: bool = True) -> str | None:
    current = Path.home()
    while True:
        print(f"\nCurrent: {current}")
        print("-" * 50)
        try:
            subfolders = sorted(
                (d for d in current.iterdir()
                 if d.is_dir() and not d.name.startswith('.')),
                key=lambda x: x.name.lower())
        except PermissionError:
            print("   [!] Permission denied. Going back up...")
            current = current.parent
            continue
        print("  [Enter] Select THIS folder")
        if current.parent != current:
            print("  [..] Go up one level")
        for i, folder in enumerate(subfolders[:20], 1):
            print(f"  [{i}] {folder.name}")
        if len(subfolders) > 20:
            print(f"  ... and {len(subfolders) - 20} more folders")
        print("  [q] Cancel" + ("/skip" if allow_skip else ""))

        nav = input("\nNavigate to: ").strip()
        if nav.lower() == 'q':
            return None
        if nav == '' or nav.lower() in ('v', 'ok'):
            print(f"   [OK] Selected: {current}")
            return str(current)
        if nav == '..':
            current = current.parent
        elif nav.isdigit() and 1 <= int(nav) <= len(subfolders):
            current = subfolders[int(nav) - 1]
        else:
            typed = _validate_path(nav)
            if typed:
                return typed
            print("   Invalid input.")


# --- Preview -----------------------------------------------------------------

def paginated_preview(lines: list[str], page_size: int = 20) -> bool:
    """Page through lines; returns True to proceed, False to abort."""
    total = len(lines)
    if total == 0:
        return True
    page, total_pages = 0, max(1, (total + page_size - 1) // page_size)
    while True:
        start, end = page * page_size, min((page + 1) * page_size, total)
        print(f"\n--- Page {page + 1}/{total_pages}  (items {start + 1}-{end} of {total}) ---")
        for line in lines[start:end]:
            print(line)
        nav = []
        if page < total_pages - 1:
            nav += ["[N]ext", "[A]ll at once"]
        if page > 0:
            nav.append("[P]rev")
        if total_pages > 2:
            nav.append("[G]o to page")
        nav += ["[Y]es proceed", "[Q]uit / abort"]
        print("  " + "  ".join(nav))
        choice = input("Choice: ").strip().upper()
        if choice in ('N', '') and page < total_pages - 1:
            page += 1
        elif choice == 'P' and page > 0:
            page -= 1
        elif choice == 'G' and total_pages > 2:
            try:
                pg = int(input(f"Go to page (1-{total_pages}): ")) - 1
                if 0 <= pg < total_pages:
                    page = pg
            except ValueError:
                pass
        elif choice == 'A':
            for line in lines:
                print(line)
            sub = input("\n[Y]es proceed  [Q]uit / abort: ").strip().upper()
            if sub in ('Y', 'Q'):
                return sub == 'Y'
        elif choice == 'Y':
            return True
        elif choice == 'Q':
            return False


def _crosses_devices(plan: Plan) -> list[Op]:
    """Moves that land on a different filesystem than their source.

    Those are byte copies, not renames — they can take minutes for a media
    file, so the preview should not present them as instantaneous.
    """
    crossing = []
    for op in plan.ops:
        if op.kind != "move" or not op.src:
            continue
        try:
            src_dev = os.stat(op.src, follow_symlinks=False).st_dev
            probe = op.dst.parent
            while not probe.exists() and probe.parent != probe:
                probe = probe.parent
            if probe.exists() and os.stat(probe, follow_symlinks=False).st_dev != src_dev:
                crossing.append(op)
        except OSError:
            continue
    return crossing


def confirm_and_execute(plan: Plan, journal: Path, dry_run: bool = False,
                        label: str = "changes", *, roots=None,
                        session: str | None = None) -> list:
    """Preview a plan, ask, execute. Returns executed journal-style entries."""
    lines = [f"  {op.kind.upper():6} {op.src if op.src else ''}  ->  {op.dst}"
             for op in plan.ops]
    lines += [f"  SKIPPED ({reason}): {op.src or op.dst}"
              for op, reason in plan.skipped]
    if not plan.ops:
        print(f"\n[OK] Nothing to do ({label}).")
        for op, reason in plan.skipped:
            print(f"  SKIPPED ({reason}): {op.src or op.dst}")
        return []
    print(f"\nPlanned {label}: {len(plan.ops)} operation(s), "
          f"{len(plan.skipped)} skipped due to conflicts.")
    crossing = _crosses_devices(plan)
    if crossing:
        print(f"  [!] {len(crossing)} of these cross a filesystem boundary and "
              f"will be COPIED, not renamed.")
        print(f"      That can take a while for large files. Each copy is "
              f"size-verified before the original is removed.")
    if dry_run:
        for line in lines:
            print(line)
        print("\n[dry-run] No changes made.")
        return []
    if not paginated_preview(lines):
        print("Aborted. No changes made.")
        return []
    result = execute(plan, journal, roots=roots, label=label, session=session)
    print(f"\n[OK] {len(result.done)} operation(s) applied.")
    for op, err in result.failed:
        print(f"  [!] FAILED: {op.src or op.dst}: {err}")
    print(f"Undo any time: menu option [7] or 'python run.py --undo' "
          f"(journal: {journal})")
    return [{'op': op.kind, 'src': str(op.src) if op.src else None,
             'dst': str(op.dst), 'ts': 0.0} for op in result.done]


# --- Flows -------------------------------------------------------------------

def _journal_path() -> Path:
    """The journal, anchored to the app rather than the current directory."""
    return journal_path()


def run_organize(tv_path: str, dry_run: bool = False, *,
                 session: str | None = None) -> None:
    root = Path(tv_path)
    show_roots = find_show_roots(root)
    if not show_roots:
        print("\n   [!] No TV shows found under this folder.")
        print("       A show folder is one that directly contains season "
              "folders\n       (\"Season 1\", \"S01\") or SxxEyy episode files.")
        return

    print(f"\n   [OK] Found {len(show_roots)} show folder(s):")
    for show in show_roots[:10]:
        print(f"       {'.' if show == root else show.relative_to(root)}")
    if len(show_roots) > 10:
        print(f"       ... and {len(show_roots) - 10} more")

    plan = Plan()
    for show in show_roots:
        plan.merge(plan_season_structure(show))
    confirm_and_execute(plan, _journal_path(), dry_run, "TV structure changes",
                        roots=[root], session=session)


def _report_misfiled(movies_rows, tv_rows, patterns) -> None:
    """Point out media that looks like it is in the wrong library.

    guessit already works out movie-vs-episode for every name; that verdict was
    computed and thrown away. Reporting it is deliberate: renames stay
    in-place, so moving something between the Movies and TV trees is a
    different, explicit operation - this just stops it being invisible.
    """
    misfiled_tv, misfiled_movies = [], []
    for row in movies_rows:
        for vf in str(row.get('Video Files') or '').split('|'):
            vf = vf.strip()
            if not vf:
                continue
            if parse_name(vf, custom_patterns=patterns).kind == 'episode':
                misfiled_tv.append(f"{row.get('Folder Name', '?')}/{vf}")
    for row in tv_rows:
        rel = str(row.get('Episode File') or '')
        if not rel or _clean_season(row.get('Season')):
            continue
        if parse_name(Path(rel).name, custom_patterns=patterns).kind == 'movie':
            misfiled_movies.append(f"{row['Show Folder']}/{rel}")

    if misfiled_tv:
        print(f"\n   [!] {len(misfiled_tv)} file(s) in the Movies folder look "
              f"like TV episodes:")
        for name in misfiled_tv[:5]:
            print(f"       {name}")
        if len(misfiled_tv) > 5:
            print(f"       ... and {len(misfiled_tv) - 5} more")
    if misfiled_movies:
        print(f"\n   [!] {len(misfiled_movies)} file(s) in the TV folder look "
              f"like movies:")
        for name in misfiled_movies[:5]:
            print(f"       {name}")
        if len(misfiled_movies) > 5:
            print(f"       ... and {len(misfiled_movies) - 5} more")
    if misfiled_tv or misfiled_movies:
        print("       Renaming leaves these where they are. Move them to the "
              "right folder\n       and re-scan if you want them organized.")


def _clean_season(val) -> str:
    return '' if val is None else str(val).strip()


def run_scan(movies_path, tv_path, excel_path: Path, dry_run: bool = False) -> None:
    patterns = load_custom_patterns()
    # Loose movie moves are part of organize, not scan. Scan is read-only.
    # The --action full flow handles loose movies in run_organize/run_rename.

    movies_rows = scan.scan_movies(Path(movies_path), patterns) if movies_path else []
    tv_rows = scan.scan_tv(Path(tv_path), patterns) if tv_path else []

    # ── recursive supplement: the structured scanners only look one level
    # deep, so deeply nested media (e.g. "Collection/Movie/file.mkv") is
    # missed.  Run the recursive walk whenever the structured scan produced
    # any placeholder rows (no actual media) or found nothing at all, and
    # merge the results — structured rows with media take priority.
    if movies_path:
        movies_has_gaps = (not movies_rows
                           or any(not r.get('Video Files') for r in movies_rows))
        if movies_has_gaps:
            print("   [!] Structured scan has gaps — running recursive walk for Movies...")
            rec = scan.scan_recursive(Path(movies_path), patterns)
            if rec:
                structured = {r['Folder Name']: r for r in movies_rows
                              if r.get('Video Files')}
                added_rows = [rr for rr in rec
                              if rr['Folder Name'] not in structured]
                movies_rows.extend(added_rows)
                # Placeholder rows for containers that now have recursive
                # descendants must go — otherwise plan_renames renames the
                # parent before the child ops and strands their paths.
                # Root-level rows ('.') have no parts and cover no placeholder.
                covered = {Path(rr['Folder Name']).parts[0] for rr in added_rows
                           if rr['Folder Name'] != '.'}
                movies_rows[:] = [r for r in movies_rows
                                  if r.get('Video Files')
                                  or r['Folder Name'] not in covered]
                print(f"   [OK] Recursive scan added {len(added_rows)} folder(s) "
                      f"(total {len(movies_rows)}).")
            elif not movies_rows:
                print("   [!] Recursive scan also found nothing.")
    if tv_path:
        # Only descend recursively for the shows that actually came up empty.
        # Triggering on "any row anywhere lacks an Episode File" meant a single
        # movie folder in the TV root dragged the entire tree through the
        # recursive path.
        gap_shows = {r['Show Folder'] for r in tv_rows
                     if not r.get('Episode File')}
        tv_has_gaps = bool(not tv_rows or gap_shows)
        if gap_shows and tv_rows:
            print(f"   [!] {len(gap_shows)} show folder(s) had no detectable "
                  f"episodes - running a recursive walk for those...")
        if tv_has_gaps:
            if not tv_rows:
                print("   [!] Structured scan found nothing — "
                      "running recursive walk for TV...")
                rec = scan.scan_recursive_tv(Path(tv_path), patterns)
            else:
                rec = []
                for show in sorted(gap_shows):
                    sub = Path(tv_path) / show if show != '.' else Path(tv_path)
                    rec.extend(scan.scan_recursive_tv(sub, patterns,
                                                      base=Path(tv_path)))
            if rec:
                # Key by (Show Folder, Episode File) so episodes from
                # different sources for the same show don't clobber each
                # other — a structured scan for "Show/S01/E01" shouldn't
                # block a recursive find of "Show/Extras/E02".
                structured = {(r['Show Folder'], r.get('Episode File', ''))
                              for r in tv_rows if r.get('Episode File')}
                added_rows = [rr for rr in rec
                              if (rr['Show Folder'], rr.get('Episode File', ''))
                              not in structured]
                tv_rows.extend(added_rows)
                # Drop placeholder show rows now covered by recursive finds
                # (same stale-parent-rename hazard as the movies side).
                # A structured Show Folder can now be a nested path
                # ("Genre/Show"), so this needs a real ancestor test — the old
                # first-component comparison left the placeholder in place, and
                # plan_renames then renamed that parent ahead of the child ops
                # it had just stranded.
                covered = {Path(rr['Show Folder']) for rr in added_rows
                           if rr['Show Folder'] != '.'}
                def _is_covered(folder: str) -> bool:
                    here = Path(folder)
                    return any(c == here or here in c.parents for c in covered)
                tv_rows[:] = [r for r in tv_rows
                              if r.get('Episode File')
                              or not _is_covered(r['Show Folder'])]
                print(f"   [OK] Recursive scan added {len(added_rows)} episode(s) "
                      f"(total {len(tv_rows)}).")
            elif not tv_rows:
                print("   [!] Recursive scan also found nothing.")
    if not movies_rows and not tv_rows:
        print("[!] Nothing found to scan — not even with recursive walk.")
        print("    Check that the path contains video files and is accessible.")
        return

    _report_misfiled(movies_rows, tv_rows, patterns)

    if dry_run:
        print(f"\n[dry-run] Would save {len(movies_rows)} movie row(s), {len(tv_rows)} TV row(s) "
              f"to {excel_path}")
        return

    append = False
    if excel_path.exists():
        append = ask_yes_no(f"'{excel_path.name}' exists. Append to it? "
                            f"(no = overwrite)", default=True)
    try:
        excel.write_library(excel_path, movies_rows, tv_rows,
                            movies_path, tv_path, append=append)
    except PermissionError:
        print(f"[!] Cannot write '{excel_path.name}' - close it in other programs first.")
        return
    print(f"[OK] Saved {len(movies_rows)} movie row(s), {len(tv_rows)} TV row(s) "
          f"to {excel_path}")


def _ask_scheme(config: dict) -> NamingScheme:
    scheme = NamingScheme.from_dict(config.get('scheme', {}))
    if config.get('scheme') and not ask_yes_no(
            "Use your saved naming scheme?", default=True):
        config.pop('scheme', None)
    if not config.get('scheme'):
        if ask_yes_no("Customize the naming scheme? (no = sensible defaults)",
                      default=False):
            for attr in vars(scheme):
                if not isinstance(getattr(scheme, attr), bool):
                    continue
                label = attr.replace('_', ' ')
                setattr(scheme, attr, ask_yes_no(f"  {label}?",
                                                 default=getattr(scheme, attr)))
        config['scheme'] = scheme.to_dict()
        _save_config(config)
    return scheme


def _llm_candidates(df_movies, df_tv, patterns) -> list[str]:
    """Names guessit couldn't confidently parse (source == 'raw')."""
    names = []
    if df_movies is not None:
        names += [str(v) for v in df_movies['Folder Name'].dropna()]
        for vfs in df_movies.get('Video Files', []):
            if vfs and str(vfs) != 'nan':
                names += [v.strip() for v in str(vfs).split('|') if v.strip()]
    if df_tv is not None:
        names += [str(v) for v in df_tv['Show Folder'].dropna().unique()]
        names += [Path(str(v)).name for v in df_tv['Episode File'].dropna()]
    return [n for n in dict.fromkeys(names)
            if parse_name(n, custom_patterns=patterns).source == "raw"]


def _ask_llm_results(df_movies, df_tv, patterns) -> dict:
    print("\nRenaming engine:")
    print("  [1] guessit (fast, offline - recommended)")
    print("  [2] guessit + local LLM (Ollama) for unparseable names")
    print("  [3] guessit + cloud LLM (OpenAI / Gemini) for unparseable names")
    choice = prompt_input("Select (1-3) [1]: ", default='1')
    if choice not in ('2', '3'):
        return {}

    candidates = _llm_candidates(df_movies, df_tv, patterns)
    if ask_yes_no(f"{len(candidates)} name(s) look hard to parse. Clean ALL "
                  f"names with the LLM instead of just those?", default=False):
        candidates = None  # signal: everything

    if candidates is not None and not candidates:
        print("[OK] Nothing needs LLM cleaning.")
        return {}
    if candidates is None:
        names = []
        if df_movies is not None:
            names += [str(v) for v in df_movies['Folder Name'].dropna()]
        if df_tv is not None:
            names += [str(v) for v in df_tv['Show Folder'].dropna().unique()]
        candidates = list(dict.fromkeys(names))

    cfg = llm.load_llm_config()
    if choice == '2':
        url = prompt_input(f"Ollama URL [{cfg.get('ollama_url', 'http://localhost:11434')}]: ",
                           default=cfg.get('ollama_url', 'http://localhost:11434'))
        models = llm.list_ollama_models(url)
        if not models:
            print(f"[!] No Ollama models found at {url}.")
            return {}
        for i, m in enumerate(models, 1):
            print(f"  [{i}] {m}")
        idx = prompt_input(f"Model (1-{len(models)}) [1]: ", default='1')
        model = models[int(idx) - 1] if idx.isdigit() and 1 <= int(idx) <= len(models) else models[0]
        cfg.update({'ollama_url': url, 'ollama_model': model})
        llm.save_llm_config(cfg)
        print(f"Cleaning {len(candidates)} name(s) with {model}...")
        return llm.clean_titles_with_llm(candidates, 'ollama', model=model,
                                         ollama_url=url)
    provider = prompt_input("Provider - [1] Gemini  [2] OpenAI [1]: ", default='1')
    provider = 'gemini' if provider != '2' else 'openai'
    key = cfg.get(f'{provider}_key') or getpass.getpass(f"{provider} API key (input hidden): ")
    if not key:
        return {}
    cfg[f'{provider}_key'] = key
    llm.save_llm_config(cfg)
    print(f"[!] API key saved to {llm.LLM_CONFIG_FILE} in plaintext. Restrict file permissions.")
    print(f"Cleaning {len(candidates)} name(s) with {provider}...")
    return llm.clean_titles_with_llm(candidates, provider, api_key=key)


def run_rename(movies_path, tv_path, excel_path: Path, dry_run: bool = False,
               *, session: str | None = None) -> None:
    if not excel_path.exists():
        print(f"[!] Spreadsheet not found: {excel_path}. Run a scan first.")
        return
    df_movies, df_tv, meta = excel.read_library(excel_path)
    movies_path = movies_path or meta.get('Movies Path')
    tv_path = tv_path or meta.get('TV Shows Path')
    patterns = load_custom_patterns()

    config = _load_config()
    scheme = _ask_scheme(config) if not dry_run else \
        NamingScheme.from_dict(config.get('scheme', {}))
    llm_results = _ask_llm_results(df_movies, df_tv, patterns) if not dry_run else {}

    plan = excel.plan_renames(df_movies, movies_path, df_tv, tv_path,
                              scheme, llm_results, patterns)
    rename_roots = [Path(p) for p in (movies_path, tv_path) if p]
    entries = confirm_and_execute(plan, _journal_path(), dry_run, "renames",
                                  roots=rename_roots, session=session)
    if entries:
        import time as _time
        for e in entries:
            e['ts'] = _time.time()
        try:
            excel.append_changes(excel_path, entries)
            print(f"[OK] Logged changes to the 'Changes' sheet of {excel_path.name}")
        except (OSError, PermissionError) as e:
            print(f"[!] Could not log to Changes sheet: {e}")


def run_extension_fixer(dry_run: bool = False) -> None:
    folder = browse_for_folder("Folder to scan for extension problems",
                               allow_skip=False)
    if not folder:
        return
    print("\n  [1] Restore missing video extensions (magic-byte detection)")
    print("  [2] Bulk-rename one extension to another (e.g. .ts -> .mp4)")
    choice = prompt_input("Select (1-2) [1]: ", default='1')
    if choice == '2':
        from_ext = prompt_input("From extension (e.g. ts): ")
        to_ext = prompt_input("To extension (e.g. mp4): ")
        if not from_ext or not to_ext:
            return
        plan = extfix.plan_extension_convert(Path(folder), from_ext, to_ext)
    else:
        plan = extfix.plan_extension_restore(Path(folder))
    confirm_and_execute(plan, _journal_path(), dry_run, "extension fixes",
                        roots=[Path(folder)])


def _fmt_ts(ts) -> str:
    try:
        return datetime.fromtimestamp(float(ts)).strftime("%Y-%m-%d %H:%M:%S")
    except (TypeError, ValueError):
        return "?"


def run_list_runs() -> None:
    """Show the journal's history so a specific run can be picked."""
    journal = _journal_path()
    runs = list_runs(journal)
    print(f"\nJournal: {journal}")
    if not runs:
        print("  (no runs recorded yet)")
        return
    print(f"\n  {'RUN':14} {'WHEN':20} {'OPS':>5}  {'STATE':9} LABEL")
    for run in runs:
        state = "undone" if run["undone"] else ("open" if run["open"] else "undoable")
        print(f"  {(run['id'] or '?'):14} {_fmt_ts(run['ts']):20} "
              f"{len(run['ops']):>5}  {state:9} {run['label'] or ''}")
    sessions = {r["session"] for r in runs if not r["undone"]}
    if len(sessions) < len([r for r in runs if not r["undone"]]):
        print("\n  Runs sharing a session were one action "
              "- 'python run.py --undo-session' reverses the whole thing.")
    print("\n  Reverse a specific run:  python run.py --undo-run <RUN>")


def _report_undo(result) -> None:
    print(f"[OK] Reverted {len(result.done)} operation(s).")
    for op, err in result.failed:
        print(f"  [!] FAILED: {op.src or op.dst}: {err}")
    if not result.ok:
        print("  [!] Some reversals failed - fix the conflicts and run undo again.")


def _run_recovery(dry_run: bool = False) -> None:
    """Clean up mutations that were interrupted mid-flight."""
    notes = recover(_journal_path(), dry_run=dry_run)
    if not notes:
        return
    print(f"\n{'[dry-run] Would recover' if dry_run else '[OK] Recovered'} "
          f"{len(notes)} interrupted change(s):")
    for note in notes:
        print(f"  {note}")


def run_undo(dry_run: bool = False, *, run_id: str | None = None,
             session: bool = False, count: int = 1,
             force: bool = False) -> None:
    journal = _journal_path()
    _run_recovery(dry_run)

    if run_id:
        result, err = undo_run(journal, run_id, dry_run=dry_run, force=force)
        if err:
            print(f"[!] {err}")
            return
        if dry_run:
            for op in result.done:
                print(f"  {op.kind.upper():6} {op.src or ''}  ->  {op.dst}")
            return
        _report_undo(result)
        return

    pending = pending_runs(journal)
    if not pending:
        print(f"[OK] Nothing to undo (no journaled runs). Journal: {journal}")
        return

    if session:
        target = pending[-1]["session"]
        group = [r for r in pending if r["session"] == target]
        total = sum(len(r["ops"]) for r in group)
        print(f"\nLast action spans {len(group)} run(s), "
              f"{total} operation(s) to reverse.")
        if dry_run:
            result = undo_session(journal, dry_run=True, force=force)
            for op in result.done:
                print(f"  {op.kind.upper():6} {op.src or ''}  ->  {op.dst}")
            return
        if not ask_yes_no("Undo the whole action now?", default=True):
            return
        _report_undo(undo_session(journal, force=force))
        return

    ops = last_run_ops(journal)
    print(f"\nLast run has {len(ops)} operation(s) to reverse.")
    if dry_run:
        result = undo_last(journal, count, dry_run=True, force=force)
        for op in result.done:
            print(f"  {op.kind.upper():6} {op.src or ''}  ->  {op.dst}")
        return
    if not ask_yes_no("Undo it now?", default=True):
        return
    _report_undo(undo_last(journal, count, force=force))


def run_text_export() -> None:
    folder = browse_for_folder("Folder to export", allow_skip=False)
    if not folder:
        return
    out = Path(prompt_input("Output file [media_library.txt]: ",
                            default="media_library.txt"))
    lines = []
    walk_errs: list[str] = []
    try:
        for dirpath, dirnames, filenames in os.walk(
            folder, onerror=lambda e: walk_errs.append(str(e))
        ):
            dirnames.sort()
            try:
                depth = Path(dirpath).relative_to(folder).parts
            except ValueError:
                depth = ()
            indent = "  " * len(depth)
            lines.append(f"{indent}{Path(dirpath).name}/")
            for f in sorted(filenames):
                lines.append(f"{indent}  {f}")
    except OSError as exc:
        print(f"   [!] Could not walk directory: {exc}")
    if walk_errs:
        print(f"   [!] {len(walk_errs)} subdirectory error(s) during walk "
              f"— export may be incomplete.")
    if not lines:
        print("[!] No files or folders found. The path may be inaccessible "
              "or empty.")
        print("    If this is a network share, ensure it is mounted and "
              "accessible from the terminal.")
        return
    try:
        out.write_text("\n".join(lines), encoding="utf-8")
    except OSError as exc:
        print(f"   [!] Could not write output file: {exc}")
        return
    print(f"[OK] Exported {len(lines)} line(s) to {out}")


# --- Menu / entry point ------------------------------------------------------

def _ask_paths(tv_only: bool = False):
    if tv_only:
        tv = browse_for_folder("Select your TV Shows folder", allow_skip=False)
        return None, tv
    movies = browse_for_folder("Select Movies folder (skip if none)")
    tv = browse_for_folder("Select TV Shows folder (skip if none)")
    return movies, tv


def _ask_excel_path() -> Path:
    name = prompt_input("Excel file name [media_library.xlsx]: ",
                        default="media_library.xlsx")
    if not name.endswith('.xlsx'):
        name += '.xlsx'
    return Path(name).resolve()


def run_wizard() -> None:
    while True:
        try:
            print("\n" + "=" * 70)
            print("MEDIA ORGANIZER")
            print("=" * 70)
            print("""
  [1] Clean file names        (scan -> preview -> rename)
  [2] Organize TV structure   (loose episodes -> Season folders)
  [3] Do both                 (organize -> scan -> rename)
  [4] Fix file extensions     (restore missing / bulk convert)
  [5] Scan library only       (create/update the Excel spreadsheet)
  [6] Export library to text file
  [7] Undo last run
  [8] Exit

Tip: type 'back' or 'b' at any prompt to return here.""")
            choice = prompt_input("\nSelect an option (1-8): ")

            if choice == '8':
                break
            elif choice == '7':
                run_list_runs()
                if pending_runs(_journal_path()):
                    run_undo(session=ask_yes_no(
                        "Undo the entire last action (all its runs)?",
                        default=False))
            elif choice == '6':
                run_text_export()
            elif choice == '4':
                run_extension_fixer()
            elif choice == '2':
                _, tv = _ask_paths(tv_only=True)
                if tv:
                    run_organize(tv)
            elif choice in ('1', '3', '5'):
                movies, tv = _ask_paths()
                if not movies and not tv:
                    print("No folders selected.")
                    continue
                xlsx = _ask_excel_path()
                if choice == '3' and tv:
                    print("\n[1/3] Organizing TV structure...")
                    run_organize(tv)
                if choice == '3':
                    print("\n[2/3] Scanning library...")
                elif choice == '1':
                    print("\n[1/2] Scanning library...")
                run_scan(movies, tv, xlsx)
                if choice in ('1', '3'):
                    print(f"\n[{'3/3' if choice == '3' else '2/2'}] Renaming...")
                    print("(You can edit the '... Fixed' columns in the spreadsheet "
                          "first - reopen this step afterwards.)")
                    if ask_yes_no("Proceed to renaming now?", default=True):
                        run_rename(movies, tv, xlsx)
            else:
                print("Invalid choice.")
        except BackNavigation:
            pass
        except KeyboardInterrupt:
            print("\nInterrupted.")
            break


def main() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            try:
                stream.reconfigure(encoding='utf-8', errors='replace')
            except Exception:
                pass

    parser = argparse.ArgumentParser(description="Media Organizer")
    parser.add_argument('--action', choices=['scan', 'organize', 'rename', 'full'])
    parser.add_argument('--movies', help="Path to movies folder")
    parser.add_argument('--tv', help="Path to TV shows folder")
    parser.add_argument('--output', help="Excel file name")
    parser.add_argument('--dry-run', action='store_true',
                        help="Show planned changes without touching disk")
    parser.add_argument('--undo', action='store_true',
                        help="Undo the last run from the journal")
    parser.add_argument('--list-runs', action='store_true',
                        help="List journaled runs and whether they can be undone")
    parser.add_argument('--undo-run', metavar='RUN',
                        help="Undo one specific run by id (see --list-runs)")
    parser.add_argument('--undo-last', type=int, metavar='N', default=None,
                        help="Undo the newest N runs")
    parser.add_argument('--undo-session', action='store_true',
                        help="Undo every run of the last action "
                             "(--action full makes several)")
    parser.add_argument('--force', action='store_true',
                        help="Undo even if a file was modified since, or out of order")
    args = parser.parse_args()

    if args.list_runs:
        run_list_runs()
        return
    if args.undo or args.undo_run or args.undo_session or args.undo_last:
        run_undo(dry_run=args.dry_run, run_id=args.undo_run,
                 session=args.undo_session,
                 count=1 if args.undo_last is None else args.undo_last,
                 force=args.force)
        return
    if not args.action:
        run_wizard()
        return

    movies = str(Path(args.movies).resolve()) if args.movies else None
    tv = str(Path(args.tv).resolve()) if args.tv else None
    xlsx = Path(args.output or "media_library.xlsx").resolve()
    if args.action == 'organize':
        if not tv:
            sys.exit("[!] --tv is required for 'organize'.")
        run_organize(tv, dry_run=args.dry_run)
    elif args.action == 'scan':
        run_scan(movies, tv, xlsx, dry_run=args.dry_run)
    elif args.action == 'rename':
        run_rename(movies, tv, xlsx, dry_run=args.dry_run)
    elif args.action == 'full':
        # One session id across all three phases so a single
        # `--undo-session` reverses the whole action.
        session = uuid.uuid4().hex[:12]
        if tv:
            run_organize(tv, dry_run=args.dry_run, session=session)
        if movies:
            loose = plan_loose_movies(Path(movies))
            if loose.ops:
                confirm_and_execute(loose, _journal_path(), args.dry_run,
                                    "loose-file moves", roots=[Path(movies)],
                                    session=session)
        run_scan(movies, tv, xlsx, dry_run=args.dry_run)
        run_rename(movies, tv, xlsx, dry_run=args.dry_run, session=session)
        if not args.dry_run:
            print("\nUndo this entire action: python run.py --undo-session")


if __name__ == "__main__":
    main()
