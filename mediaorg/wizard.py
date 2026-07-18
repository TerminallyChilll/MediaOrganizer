"""Interactive wizard and CLI entry point. All console UI lives here.

ASCII markers only ([OK], [!], ->): no emoji, no cp1252 crashes.
"""

import argparse
import getpass
import json
import os
import sys
from pathlib import Path

from . import excel, extfix, llm, scan
from .execute import JOURNAL_FILE, execute, last_run_ops, undo_last_run
from .parse import load_custom_patterns, parse_name
from .plan import (NamingScheme, Plan, folder_has_episodes_or_seasons,
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
                print(f"   [OK] Selected: {folder}")
                return folder
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


def confirm_and_execute(plan: Plan, journal: Path, dry_run: bool = False,
                        label: str = "changes") -> list:
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
    if dry_run:
        for line in lines:
            print(line)
        print("\n[dry-run] No changes made.")
        return []
    if not paginated_preview(lines):
        print("Aborted. No changes made.")
        return []
    result = execute(plan, journal)
    print(f"\n[OK] {len(result.done)} operation(s) applied.")
    for op, err in result.failed:
        print(f"  [!] FAILED: {op.src or op.dst}: {err}")
    print(f"Undo any time: menu option [7] or 'python run.py --undo' "
          f"(journal: {journal})")
    return [{'op': op.kind, 'src': str(op.src) if op.src else None,
             'dst': str(op.dst), 'ts': 0.0} for op in result.done]


# --- Flows -------------------------------------------------------------------

def _journal_path() -> Path:
    return Path.cwd() / JOURNAL_FILE


def run_organize(tv_path: str, dry_run: bool = False) -> None:
    root = Path(tv_path)
    plan = Plan()
    if folder_has_episodes_or_seasons(root):
        plan.merge(plan_season_structure(root))
    else:
        for entry in sorted(os.scandir(root), key=lambda e: e.name):
            if entry.is_dir(follow_symlinks=False):
                child = Path(entry.path)
                if folder_has_episodes_or_seasons(child):
                    plan.merge(plan_season_structure(child))
                else:
                    # Descend one more level for nested Show/Show/Season layouts.
                    try:
                        for sub in sorted(os.scandir(child), key=lambda e: e.name):
                            if sub.is_dir(follow_symlinks=False) and \
                                    folder_has_episodes_or_seasons(Path(sub.path)):
                                plan.merge(plan_season_structure(Path(sub.path)))
                    except OSError:
                        pass
    confirm_and_execute(plan, _journal_path(), dry_run, "TV structure changes")


def run_scan(movies_path, tv_path, excel_path: Path, dry_run: bool = False) -> None:
    patterns = load_custom_patterns()
    # Loose movie moves are part of organize, not scan. Scan is read-only.
    # The --action full flow handles loose movies in run_organize/run_rename.

    movies_rows = scan.scan_movies(Path(movies_path), patterns) if movies_path else []
    tv_rows = scan.scan_tv(Path(tv_path), patterns) if tv_path else []
    if not movies_rows and not tv_rows:
        print("[!] Nothing found to scan.")
        return

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


def run_rename(movies_path, tv_path, excel_path: Path, dry_run: bool = False) -> None:
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
    entries = confirm_and_execute(plan, _journal_path(), dry_run, "renames")
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
    confirm_and_execute(plan, _journal_path(), dry_run, "extension fixes")


def run_undo(dry_run: bool = False) -> None:
    journal = _journal_path()
    ops = last_run_ops(journal)
    if not ops:
        print("[OK] Nothing to undo (no journaled runs).")
        return
    print(f"\nLast run has {len(ops)} operation(s) to reverse.")
    if dry_run:
        result = undo_last_run(journal, dry_run=True)
        for op in result.done:
            print(f"  {op.kind.upper():6} {op.src or ''}  ->  {op.dst}")
        return
    if not ask_yes_no("Undo it now?", default=True):
        return
    result = undo_last_run(journal)
    print(f"[OK] Reverted {len(result.done)} operation(s).")
    for op, err in result.failed:
        print(f"  [!] FAILED: {op.src or op.dst}: {err}")
    if not result.ok:
        print("  [!] Some reversals failed - fix the conflicts and run undo again.")


def run_text_export() -> None:
    folder = browse_for_folder("Folder to export", allow_skip=False)
    if not folder:
        return
    out = Path(prompt_input("Output file [media_library.txt]: ",
                            default="media_library.txt"))
    lines = []
    for dirpath, dirnames, filenames in os.walk(folder):
        dirnames.sort()
        depth = Path(dirpath).relative_to(folder).parts
        indent = "  " * len(depth)
        lines.append(f"{indent}{Path(dirpath).name}/")
        for f in sorted(filenames):
            lines.append(f"{indent}  {f}")
    out.write_text("\n".join(lines), encoding="utf-8")
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
                run_undo()
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
    args = parser.parse_args()

    if args.undo:
        run_undo(dry_run=args.dry_run)
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
        if tv:
            run_organize(tv, dry_run=args.dry_run)
        if movies:
            loose = plan_loose_movies(Path(movies))
            if loose.ops:
                confirm_and_execute(loose, _journal_path(), args.dry_run, "loose-file moves")
        run_scan(movies, tv, xlsx, dry_run=args.dry_run)
        run_rename(movies, tv, xlsx, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
