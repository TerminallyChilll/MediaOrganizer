"""Interactive wizard and CLI entry point. All console UI lives here.

ASCII markers only ([OK], [!], ->): no emoji, no cp1252 crashes.
"""

import argparse
import dataclasses
import enum
import getpass
import json
import os
import re
import sys
import time
import uuid
from datetime import datetime
from pathlib import Path

from . import excel, extfix, llm, scan, update
from .execute import (execute, journal_path, last_run_ops, list_runs,
                      pending_runs, recover, undo_last, undo_last_run,
                      undo_run, undo_session)
from .parse import (COMPANION_EXTS, VIDEO_EXTS, companion_tail,
                    custom_patterns_path, is_media_file, load_custom_patterns,
                    parse_name, pre_clean, save_custom_patterns)
from .plan import (NamingScheme, Op, Plan, check_collisions, find_show_roots,
                   plan_loose_movies, plan_season_structure, sanitize)

CONFIG_FILE = ".media_renamer_config.json"
#: Seconds a launch will ever spend waiting for the update check.
LAUNCH_CHECK_BUDGET = 2.0


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


def config_path() -> Path:
    """Where the remembered folders live, independent of the current directory.

    Same reasoning as the undo journal and the word list: a cwd-relative file
    meant that launching from anywhere but the app folder — `python
    /opt/MediaOrganizer/run.py` from your home directory — silently forgot the
    folders you picked last time and asked for them again, which looks like
    the feature is broken. This was the last of the four state files still
    resolved against the cwd; making it match the others is also what lets the
    re-clone advice name a folder to copy it *from*.

    Order: ``$MEDIAORG_CONFIG`` -> next to the app -> adopt a pre-existing file
    in the cwd, so upgrading users keep the folders they already have.
    """
    env = os.environ.get("MEDIAORG_CONFIG")
    if env:
        return Path(env).expanduser()
    app = Path(__file__).resolve().parent.parent / CONFIG_FILE
    if not app.exists():
        legacy = Path.cwd() / CONFIG_FILE
        if legacy.exists() and legacy != app:
            return legacy
    return app


def _load_config() -> dict:
    try:
        with open(config_path(), encoding='utf-8') as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def _save_config(config: dict) -> None:
    try:
        with open(config_path(), 'w', encoding='utf-8') as f:
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


# --- Preview and review ------------------------------------------------------

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


def _stdin_is_interactive() -> bool:
    """Is there a human who could answer a prompt?

    A cron job, a Docker service and a piped run all reach the same prompts an
    interactive session does. Asking them a question produces an EOFError over
    an already-mutated library, which is the worst of both worlds.
    """
    try:
        return bool(sys.stdin) and sys.stdin.isatty()
    except (AttributeError, ValueError):
        return False    # detached or closed stdin


def _parse_item_numbers(arg: str, total: int) -> list[int] | None:
    """"3", "3,7", "3-5,9" -> zero-based indices. None if it doesn't parse.

    Reviewing a four-hundred-item rename one number at a time is not review,
    it is data entry, so ranges are worth the twenty lines.
    """
    picked: list[int] = []
    for chunk in arg.replace(' ', ',').split(','):
        if not chunk:
            continue
        lo, sep, hi = chunk.partition('-')
        # isdigit() is not a promise that int() will succeed: it is True for
        # '²' and '①', and for decimal strings past CPython's conversion limit.
        # An uncaught ValueError here would unwind past the accept gate.
        try:
            start, end = int(lo), int(hi) if sep else int(lo)
        except ValueError:
            return None
        if start < 1 or end > total or start > end:
            return None
        picked.extend(range(start - 1, end))
    return picked or None


def _item_lines(item: "_ReviewItem", number: int, width: int,
                crossing: set) -> list[str]:
    """The one or two printed lines for a single change."""
    num = f"[{number:>{width}}]"
    tags = []
    if not item.keep:
        tags.append("EXCLUDED")
    if item.op.dst != item.original.dst:
        tags.append("edited")
    # Keyed on the *original* op: Op is frozen and compared by value, so an
    # edited item is a different value and would silently lose the warning —
    # on exactly the item the user stopped to look at. Retyping a leaf name
    # cannot change which filesystem it lands on.
    if item.original in crossing:
        tags.append("COPIED across drives")
    tag = f"   ({', '.join(tags)})" if tags else ""
    if item.op.kind == "move" and item.op.src:
        return [f"  {num} BEFORE  {item.op.src}{tag}",
                f"  {' ' * len(num)} AFTER   {item.op.dst}"]
    verb = "NEW FOLDER" if item.op.kind == "mkdir" else "REMOVE EMPTY FOLDER"
    return [f"  {num} {verb}  {item.op.dst}{tag}"]


def _review_lines(items: list["_ReviewItem"], crossing: set,
                  start: int = 0, end: int | None = None) -> list[str]:
    """Render items[start:end]. Paging slices *items*, never rendered lines.

    A move prints two lines and a folder op prints one, so a fixed line stride
    puts the page break between a BEFORE and its AFTER as soon as an odd number
    of folder ops precedes it — and plan_loose_movies interleaves exactly that.
    """
    width = len(str(len(items)))
    lines = []
    for offset, item in enumerate(items[start:end], start + 1):
        lines.extend(_item_lines(item, offset, width, crossing))
    return lines


@dataclasses.dataclass
class _ReviewItem:
    op: Op
    original: Op          # what the planner proposed, so edits can be shown
    keep: bool = True


def _ask_new_name(current: str, what: str) -> str | None:
    """Prompt for a replacement leaf name. None means "leave it alone".

    Read with a plain input(): this is a value, not a menu step, so "b" has to
    be a name you can give a file rather than a command that unwinds the whole
    review.
    """
    typed = input(f"   New {what} name (blank to keep): ").strip()
    if not typed:
        return None
    if any(sep in typed for sep in ('/', '\\')) or typed in ('.', '..'):
        print("   [!] This changes the name only, not where the item lands. "
              "No slashes.")
        return None
    # sanitize() is what every generated name goes through, so a typed one
    # has to clear the same bar - but silently rewriting what someone typed
    # is worse than telling them, so this asks.
    safe = sanitize(typed)
    if safe != typed:
        print(f"   [!] '{typed}' is not usable on every filesystem.")
        # 'b' here means "not this rename", not "throw away the review" —
        # prompt_input raises BackNavigation, which would otherwise unwind all
        # the way to the menu and discard every edit made so far, silently.
        try:
            use_it = ask_yes_no(f"   Use '{safe}' instead?", default=True)
        except (EOFError, BackNavigation):
            use_it = False
        if not use_it:
            return None
    return None if safe == current else safe


def _edit_destination(items: list[_ReviewItem], index: int) -> None:
    """Retype the final name of one move, carrying its companions along."""
    item = items[index]
    if item.op.kind != "move":
        print("   [!] Only a rename/move has a name to change. Use 'x' to "
              "exclude a folder operation instead.")
        return
    old = item.op.dst
    # A folder has no extension to protect and no companions to drag along.
    is_file = item.op.src is not None and item.op.src.is_file()
    print(f"\n   now      : {item.op.src}")
    print(f"   proposed : {old.name}")
    safe = _ask_new_name(old.name, "file" if is_file else "folder")
    if safe is None:
        return

    if is_file:
        # Compared against the *planned* suffix, not the source's. The planner
        # often changes the suffix on purpose: plan_extension_convert turns
        # ".ts" into ".mp4", and plan_extension_restore appends ".mkv" to a
        # name whose "suffix" is really ".1080p". Judging by the source there
        # made the default answer silently undo the very conversion that was
        # asked for.
        want_ext = old.suffix
        if want_ext and Path(safe).suffix.lower() != want_ext.lower():
            got = Path(safe).suffix or "no extension"
            try:
                change_ext = ask_yes_no(
                    f"   That gives {got}, but this was going to be "
                    f"{want_ext}. Really change the extension?", default=False)
            except (EOFError, BackNavigation):
                change_ext = False      # same reasoning as the prompt above
            if not change_ext:
                # Append rather than restem. Media names are dot-delimited and
                # Path.stem eats only the last segment, so restemming turned
                # "Show.S01E01.1080p" into "Show.S01E01.mkv" — dropping the
                # quality tag the user had just typed out.
                safe = safe + want_ext
                print(f"   Using: {safe}")
        if safe == old.name:
            return

    # Subtitles and .nfo sidecars are separate ops named after the video.
    # Renaming the video alone would leave them pointing at a name that no
    # longer exists — exactly the breakage the rename is meant to fix.
    #
    # Matched with parse.companion_tail, the same helper excel._companion_ops
    # uses to produce these destinations: it keeps ".en"/".fr" language tags, so
    # "My Show S01E01.en.srt" never has the video's stem, and its boundary check
    # is what stops "Episode1" from capturing "Episode10.en.srt". One helper
    # rather than two copies of the rule, because a planner and the screen that
    # reviews its output cannot be allowed to disagree about what a sidecar is.
    #
    # Carrying the tail is not optional: matching loosely *without* it would be
    # worse than the bug - ".en.srt" and ".fr.srt" would both target one name
    # and check_collisions would drop both as a duplicate target.
    #
    # It flows one way only, video -> sidecar. Symmetric matching meant that
    # renaming "Episode.srt" to "English.srt" dragged "Episode.mkv" with it.
    leads = is_file and old.suffix.lower() in VIDEO_EXTS
    followers = []
    if leads:
        # Which video owns a sidecar is a question about the whole folder, not
        # about two names, so it is settled here rather than in companion_tail:
        # "New.2010.1080p.en.srt" is a valid tail of both "New" and
        # "New.2010.1080p", and only the longer one should claim it. Same rule
        # as excel._companion_ops, over planned destinations instead of the
        # directory listing.
        video_dsts = [o.op.dst for o in items
                      if o.op.kind == "move" and o.op.dst.parent == old.parent
                      and o.op.dst.suffix.lower() in VIDEO_EXTS]
        for o in items:
            if (o is item or o.op.kind != "move" or o.op.src is None
                    or not o.op.src.is_file()
                    or o.op.dst.parent != old.parent
                    or o.op.dst.suffix.lower() not in COMPANION_EXTS):
                continue
            tail = companion_tail(o.op.dst.stem, old.stem)
            if tail is None:
                continue
            if any(len(v.stem) > len(old.stem)
                   and companion_tail(o.op.dst.stem, v.stem) is not None
                   for v in video_dsts):
                continue    # a more specific video in this plan owns it
            followers.append((o, tail))
    item.op = dataclasses.replace(item.op, dst=old.with_name(safe))
    print(f"   [OK] {safe}")
    new_stem = Path(safe).stem
    for follower, tail in followers:
        renamed = follower.op.dst.with_name(
            new_stem + tail + follower.op.dst.suffix)
        follower.op = dataclasses.replace(follower.op, dst=renamed)
        print(f"   [OK] companion follows it: {renamed.name}")


def review_changes(plan: Plan, label: str = "changes") -> Plan | None:
    """Show the before/after list; let it be trimmed or corrected first.

    Returns the plan to run, or None if the user backed out. The result is
    always re-checked by :func:`check_collisions`, because a hand-typed
    destination can collide exactly as readily as a planned one.
    """
    crossing = set(_crosses_devices(plan))
    items = [_ReviewItem(op, op) for op in plan.ops]
    # Backstop for anything the individual prompts did not already handle.
    # Ctrl-D means no further answers are coming, so abort; treating it as an
    # invalid choice would spin the loop forever on a closed stdin.
    # BackNavigation should never arrive here — the sub-prompts inside
    # _edit_destination catch their own, so 'b' cancels one rename instead of
    # the review — but if a future prompt forgets, this makes the failure a
    # message rather than every edit vanishing on the way to the menu.
    #
    # The wording is deliberately about *these* changes: review_changes also
    # runs in the rename phase of [4], where the organize phase is already on
    # disk, so "nothing has been changed" would be false there.
    try:
        if _review_items(items, crossing, label, plan.skipped) is False:
            return None
    except (EOFError, BackNavigation):
        print("\n   [!] End of input. None of these changes were applied.")
        return None

    keep = [i.op for i in items if i.keep]
    dropped = [i.op for i in items if not i.keep]
    if not keep:
        print("Every change was excluded. Nothing to do.")
        return None
    if not dropped and all(i.op is i.original for i in items):
        return plan   # untouched: keep the planner's own skipped list

    revalidated = check_collisions(keep, dropped=dropped)
    # The planner's own skips are still true and still worth showing.
    revalidated.skipped = list(plan.skipped) + revalidated.skipped
    newly = len(revalidated.skipped) - len(plan.skipped)
    if newly:
        print(f"\n[!] {newly} of your changes cannot be applied as edited:")
        for op, reason in revalidated.skipped[len(plan.skipped):]:
            print(f"  SKIPPED ({reason}): {op.src or op.dst}")
        if not revalidated.ops:
            print("Nothing left to do.")
            return None
        try:
            go_on = ask_yes_no(f"Continue with the remaining "
                               f"{len(revalidated.ops)}?", default=False)
        except (EOFError, BackNavigation):
            go_on = False   # both mean "no" here; still nothing applied
        if not go_on:
            return None
    return revalidated


def _review_items(items: list[_ReviewItem], crossing: set, label: str,
                  skipped=(), page_size: int = 10) -> bool:
    """The review screen: a paged, numbered before/after list.

    Every change is on screen before anything is asked, and Enter pages forward
    rather than applying — a bare Enter must never be the thing that commits a
    library-wide rename it has not shown you.

    `R` turns on the editing commands so paging and editing are one screen.
    Returns True to apply, False to abort the whole run.
    """
    page, editing = 0, False
    while True:
        total_pages = max(1, (len(items) + page_size - 1) // page_size)
        page = min(page, total_pages - 1)
        start, end = page * page_size, min((page + 1) * page_size, len(items))
        kept = sum(1 for i in items if i.keep)
        print(f"\n--- {label}: items {start + 1}-{end} of {len(items)}, "
              f"page {page + 1}/{total_pages} "
              f"({kept} will be applied) ---")
        for line in _review_lines(items, crossing, start, end):
            print(line)
        # The planner's own skips, with the per-file reasons. A bare count told
        # the user something was dropped without ever saying what or why.
        if skipped and page == total_pages - 1:
            print(f"\n  {len(skipped)} change(s) the planner had to skip:")
            for op, reason in list(skipped)[:10]:
                print(f"    SKIPPED ({reason}): {op.src or op.dst}")
            if len(skipped) > 10:
                print(f"    ... and {len(skipped) - 10} more")

        nav = []
        if total_pages > 1:
            nav += ["[Enter/N]ext", "[P]rev", "[G] page", "[A]ll"]
        if editing:
            nav += ["[x N] exclude", "[k N] keep", "[e N] rename"]
        else:
            nav.append("[R] review & edit")
        nav += ["[Y]es apply", "[Q]uit"]
        print("  " + "  ".join(nav))
        if editing:
            print("  (x and k take ranges too: 'x 3,7-9')")
        # Plain input(): 'b'/'back' must be an invalid choice here, not an
        # exception that unwinds to the menu and silently discards every
        # exclusion and rename made so far.
        raw = input("Choice: ").strip()
        cmd, _, arg = raw.partition(' ')
        cmd, arg = cmd.strip().upper(), arg.strip()

        if cmd in ('N', '') and page < total_pages - 1:
            page += 1
        elif cmd in ('N', '') and page == total_pages - 1:
            print("   [!] That was the last page. [Y] applies, [Q] cancels.")
        elif cmd == 'P' and page > 0:
            page -= 1
        elif cmd == 'A':
            for line in _review_lines(items, crossing):
                print(line)
        elif cmd == 'G':
            target = arg or input(f"Go to page (1-{total_pages}): ").strip()
            try:
                pg = int(target)
            except ValueError:
                pg = -1
            if 1 <= pg <= total_pages:
                page = pg - 1
            else:
                print("   [!] No such page.")
        elif cmd == 'R':
            editing = True
            print("   [OK] Editing on: exclude with 'x N', rename with 'e N'.")
        elif cmd in ('X', 'K') and editing:
            picked = _parse_item_numbers(arg, len(items))
            if picked is None:
                print(f"   [!] Give an item number, 1-{len(items)} "
                      f"(e.g. '{cmd.lower()} 3' or '{cmd.lower()} 3,7-9').")
                continue
            for i in picked:
                items[i].keep = (cmd == 'K')
            verb = "excluded" if cmd == 'X' else "kept"
            print(f"   [OK] {len(picked)} item(s) {verb}.")
        elif cmd == 'E' and editing:
            picked = _parse_item_numbers(arg, len(items))
            if picked is None or len(picked) != 1:
                print("   [!] Rename one item at a time, e.g. 'e 3'.")
                continue
            _edit_destination(items, picked[0])
        elif cmd in ('X', 'K', 'E'):
            print("   [!] Press [R] first to turn on editing.")
        elif cmd == 'Y':
            return True
        elif cmd == 'Q':
            return False
        else:
            print("   [!] Not a command here.")


def confirm_and_execute(plan: Plan, journal: Path, dry_run: bool = False,
                        label: str = "changes", *, roots=None,
                        session: str | None = None,
                        accept_gate: bool = True) -> list:
    """Preview a plan, review it, execute, then ask to keep or put it back.

    `accept_gate` False skips only the last question, never the review: the
    multi-phase flows ask it once for the whole session instead of once per
    phase, so answering "no" puts back everything rather than the last step.

    When stdin is not a terminal — a cron job, a Docker service, a piped run —
    there is nobody to answer either question, so the whole plan is printed and
    applied and the undo command is given. Interactively, nothing is ever
    applied that has not been listed on screen first.
    """
    interactive = _stdin_is_interactive()
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
        for op in plan.ops:
            print(f"  {op.kind.upper():6} {op.src if op.src else ''}  ->  "
                  f"{op.dst}")
        for op, reason in plan.skipped:
            print(f"  SKIPPED ({reason}): {op.src or op.dst}")
        print("\n[dry-run] No changes made.")
        return []
    if interactive:
        reviewed = review_changes(plan, label)
        if reviewed is None:
            print("Aborted. No changes made.")
            return []
        plan = reviewed
    else:
        # Unattended: still show every change, since the log is the only record
        # anyone will read afterwards.
        print("(stdin is not a terminal - applying without prompting.)")
        for op in plan.ops:
            print(f"  {op.kind.upper():6} {op.src if op.src else ''}  ->  "
                  f"{op.dst}")
        for op, reason in plan.skipped:
            print(f"  SKIPPED ({reason}): {op.src or op.dst}")
    result = execute(plan, journal, roots=roots, label=label, session=session)
    print(f"\n[OK] {len(result.done)} operation(s) applied.")
    for op, err in result.failed:
        print(f"  [!] FAILED: {op.src or op.dst}: {err}")
    if accept_gate and interactive:
        outcome = accept_or_revert(result, journal, label=label)
        # REVERT_FAILED deliberately falls through to the return below: the
        # library is still renamed, so the caller must log it. Reporting
        # nothing there would leave the audit trail blank for the one state
        # that most needs one.
        if outcome in (GateOutcome.REVERTED, GateOutcome.NOTHING):
            return []
    else:
        undo = (f"python run.py --undo-run {result.run_id}" if result.run_id
                else "python run.py --undo")
        print(f"Undo any time: menu option [9] or '{undo}' "
              f"(journal: {journal})")
    return [{'op': op.kind, 'src': str(op.src) if op.src else None,
             'dst': str(op.dst), 'ts': 0.0} for op in result.done]


def _print_capped(lines: list[str], cap: int = 30) -> None:
    """Print a list, offering the rest rather than flooding the terminal."""
    for line in lines[:cap]:
        print(line)
    if len(lines) > cap:
        print(f"  ... and {len(lines) - cap} more")
        try:
            more = ask_yes_no("Show the full list?", default=False)
        except (BackNavigation, EOFError):
            # This is a viewer, not a step. Neither 'back' nor a closed stdin
            # may unwind out of the accept gate that is about to be asked —
            # that would leave the library changed with the question never put.
            more = False
        if more:
            for line in lines[cap:]:
                print(line)


def _session_ops(journal: Path, session: str) -> list[dict]:
    """Journal entries for every run of one action that is still standing."""
    ops = []
    for run in pending_runs(journal):
        if run["session"] == session:
            ops.extend(run["ops"])
    return ops


def _undo_commands(journal: Path, session: str | None, run_id: str | None
                   ) -> list[str]:
    """The exact commands that reverse what just happened.

    Not `--undo-session`: it takes no id and resolves to the *newest pending*
    session, so advice meant for later can reverse a different batch entirely —
    and a session that gets skipped over is then unreachable, since session ids
    are never displayed anywhere. Run ids are addressable and are printed by
    `--list-runs`, so name them.
    """
    if session:
        ids = [r["id"] for r in pending_runs(journal)
               if r["session"] == session and r["id"]]
        if ids:
            return [f"python run.py --undo-run {i}" for i in reversed(ids)]
    if run_id:
        return [f"python run.py --undo-run {run_id}"]
    return ["python run.py --undo"]


def _report_unaccepted(journal: Path, session: str) -> None:
    """Say what a session left on disk when its gate was never reached.

    Only ever called on the way out of an interrupt or a crash, so it prints
    and returns rather than asking anything.
    """
    landed = _session_ops(journal, session)
    if not landed:
        return
    print(f"\n[!] {len(landed)} change(s) were already applied and you never "
          f"got the chance to accept them.")
    print("    They are still on disk. Put them back with:")
    for cmd in _undo_commands(journal, session, None):
        print(f"        {cmd}")


class GateOutcome(str, enum.Enum):
    """What the acceptance gate actually did.

    A bool could not tell "reverted" from "the revert failed", and the caller
    needs to: after a failed revert the library is still renamed, so the
    changes must still be recorded in the spreadsheet.
    """
    KEPT = "kept"
    REVERTED = "reverted"
    REVERT_FAILED = "revert-failed"
    NOTHING = "nothing"


def accept_or_revert(result, journal: Path, *, label: str = "changes",
                     session: str | None = None) -> GateOutcome:
    """The last gate: keep what just happened, or put every file back.

    Anything but a positive yes — "n", a bare Enter, 'back', or a closed
    stdin — reverses the change through the journal, using the same code path
    as `--undo`. `session` reverses every run of a multi-phase action rather
    than only its last phase.
    """
    if session:
        entries = _session_ops(journal, session)
        pairs = [(e["src"], e["dst"]) for e in entries if e["op"] == "move"]
        total = len(entries)
    else:
        if result is None or not result.done:
            return GateOutcome.NOTHING   # nothing landed, nothing to accept
        pairs = [(str(op.src), str(op.dst))
                 for op in result.done if op.kind == "move"]
        total = len(result.done)
    if not total:
        return GateOutcome.NOTHING

    undo_cmds = _undo_commands(journal, session,
                               result.run_id if result is not None else None)
    hint = "\n".join(f"        {c}" for c in undo_cmds)
    folders = total - len(pairs)
    # Everything from the banner onwards is inside the try: _print_capped asks
    # its own question, and an interrupt there used to escape with the library
    # changed and nothing said about it.
    try:
        print("\n" + "!" * 70)
        print(f"  CHECK YOUR FILES NOW - {total} change(s) are already on disk.")
        print("!" * 70)
        print(f"  Open the folder in another window and make sure the {label} are")
        print("  what you wanted. This is your last chance to have them undone")
        print("  automatically.")
        print("\n  Answering NO puts every one of these files back exactly where")
        print("  it was. Answering YES keeps them (you can still undo later, but")
        print("  you will have to ask for it).")
        if pairs:
            print(f"\n  What changed ({len(pairs)} file/folder rename(s)):")
            _print_capped([f"  BEFORE  {src}\n  AFTER   {dst}\n"
                           for src, dst in pairs])
        if folders:
            print(f"  ...plus {folders} folder(s) created or removed.")
        keep = ask_yes_no("\nKeep these changes?", default=False)
    except BackNavigation:
        # 'back' must not escape to the menu leaving the library changed and
        # the question unanswered. It is not a yes, so it is a no.
        print("   Taking that as 'no'.")
        keep = False
    except EOFError:
        # stdin closed mid-gate (a pipe that ran dry). Same reasoning as
        # 'back': not a yes, and a traceback over a mutated library is the
        # worst possible answer.
        print("\n   [!] No answer available - treating that as 'no'.")
        keep = False
    except KeyboardInterrupt:
        # Deliberately NOT treated as a rejection: reverting would start a
        # large unasked-for mutation at the exact moment the user is trying to
        # stop the program. Say plainly what is on disk and how to reverse it.
        print(f"\n[!] Interrupted before you answered - the {total} change(s) "
              f"are still on disk.")
        print("    Put them back with:")
        print(hint)
        raise

    if keep:
        print("\n[OK] Kept. Undo later with:")
        print(hint)
        print(f"     or menu option [9]. (journal: {journal})")
        return GateOutcome.KEPT

    print(f"\nPutting {total} change(s) back...")
    if session:
        undone = undo_session(journal, session)
    elif result.run_id:
        undone, err = undo_run(journal, result.run_id)
        if err:
            # The likeliest error is the out-of-order refusal, whose remedy is
            # --force. Printing the message without the remedy left the user
            # with a mutated library and no next step.
            print(f"[!] Could not reverse the run: {err}")
            print(f"    The {total} change(s) are still in their new locations. "
                  f"Force it with:")
            print("\n".join(f"{c} --force" for c in hint.splitlines()))
            return GateOutcome.REVERT_FAILED
    else:
        undone = undo_last_run(journal)
    _report_undo(undone)
    if not undone.ok:
        print("    The rest are still in their new locations. Force it with:")
        print("\n".join(f"{c} --force" for c in hint.splitlines()))
        return GateOutcome.REVERT_FAILED
    return GateOutcome.REVERTED


# --- Flows -------------------------------------------------------------------

def _journal_path() -> Path:
    """The journal, anchored to the app rather than the current directory."""
    return journal_path()


def run_organize(tv_path: str, dry_run: bool = False, *,
                 session: str | None = None,
                 accept_gate: bool = True) -> None:
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
    print("   (This step moves episodes into Season folders. It is decided by "
          "the\n    episode codes on disk, so neither the word list nor an LLM "
          "is used.)")
    confirm_and_execute(plan, _journal_path(), dry_run, "TV structure changes",
                        roots=[root], session=session,
                        accept_gate=accept_gate)


def run_organize_movies(movies_path: str, dry_run: bool = False, *,
                        session: str | None = None,
                        accept_gate: bool = True) -> None:
    """Give each loose movie file in the movies root its own folder.

    Only reachable from `--action full` before now, so anyone using the menu
    never saw it — and loose files in the root are invisible to `scan_movies`
    (it walks directories only), so those movies were silently left out of
    both the spreadsheet and the rename.
    """
    root = Path(movies_path)
    plan = plan_loose_movies(root)
    if not plan.ops:
        print("\n   [OK] No loose movie files in the top of that folder.")
        for op, reason in plan.skipped:
            print(f"   SKIPPED ({reason}): {op.src or op.dst}")
        return
    movies = len({op.dst.parent for op in plan.ops if op.kind == "move"})
    print(f"\n   [OK] {movies} loose movie file(s) to give a folder of their own.")
    confirm_and_execute(plan, _journal_path(), dry_run, "loose-file moves",
                        roots=[root], session=session,
                        accept_gate=accept_gate)


def _warn_loose_movies(movies_path) -> None:
    """Point out loose files the media scan cannot see."""
    if not movies_path:
        return
    try:
        loose = [e.name for e in os.scandir(movies_path)
                 if e.is_file(follow_symlinks=False) and is_media_file(e.name)]
    except OSError:
        return
    if not loose:
        return
    print(f"\n   [!] {len(loose)} movie file(s) sit loose in the top of the "
          f"Movies folder:")
    for name in sorted(loose)[:5]:
        print(f"       {name}")
    if len(loose) > 5:
        print(f"       ... and {len(loose) - 5} more")
    print("       A movie scan only looks inside folders, so these are not in "
          "the\n       spreadsheet and will not be renamed. Menu option "
          "[3] gives each\n       one its own folder first.")


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
    env_key = llm.env_value(f'{provider}_key')
    key = env_key or cfg.get(f'{provider}_key')
    if key:
        source = (f"the {llm.ENV_KEYS[f'{provider}_key']} environment variable"
                  if env_key else llm.llm_config_path())
        print(f"[OK] Using the saved {provider} key from {source}.")
    else:
        key = getpass.getpass(f"{provider} API key (input hidden): ")
        if not key:
            return {}
        # Only a key typed here is written to disk. One supplied through the
        # environment stays in the environment — copying it into a plaintext
        # file the user never asked for would be a nasty surprise.
        cfg[f'{provider}_key'] = key
        llm.save_llm_config(cfg)
        print(f"[!] API key saved to {llm.llm_config_path()} in plaintext "
              f"(permissions 0600). Set "
              f"{llm.ENV_KEYS[f'{provider}_key']} instead to avoid storing it.")
    print(f"Cleaning {len(candidates)} name(s) with {provider}...")
    return llm.clean_titles_with_llm(candidates, provider, api_key=key)


def _warn_stale_spreadsheet(excel_path: Path) -> None:
    """The scan ran between the organize and the revert, so the sheet lies.

    A multi-phase action goes organize -> scan -> rename -> gate. Rejecting at
    the gate reverses the organize as well, but the spreadsheet was written in
    between and records the post-organize paths. Those paths no longer exist,
    so a later rename would plan from sources that are gone.
    """
    if not excel_path.exists():
        return
    print(f"\n[!] {excel_path.name} was written before those changes were put "
          f"back,\n    so it now describes a layout that no longer exists.")
    print("    Run a scan again before renaming from it.")


def log_changes(excel_path: Path, entries: list) -> None:
    """Record accepted renames in the spreadsheet's 'Changes' sheet.

    Only ever called once the user has accepted them: a sheet that lists
    renames which were then put back is worse than no sheet at all.
    """
    if not entries:
        return
    for e in entries:
        e['ts'] = time.time()
    try:
        excel.append_changes(excel_path, entries)
        print(f"[OK] Logged changes to the 'Changes' sheet of {excel_path.name}")
    except (OSError, PermissionError) as e:
        print(f"[!] Could not log to Changes sheet: {e}")


def run_rename(movies_path, tv_path, excel_path: Path, dry_run: bool = False,
               *, session: str | None = None,
               accept_gate: bool = True, log: bool = True) -> list:
    """Plan and apply renames. Returns the entries that were applied.

    The caller gets the entries back so a multi-phase flow can hold off on
    logging them until its own accept gate has been answered.
    """
    if not excel_path.exists():
        print(f"[!] Spreadsheet not found: {excel_path}. Run a scan first.")
        return []
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
                                  roots=rename_roots, session=session,
                                  accept_gate=accept_gate)
    # `log` is False only for the multi-phase flows, which log after their own
    # session gate. It is NOT tied to accept_gate: doing that stopped
    # `--action rename` writing the sheet at all, since it passes
    # accept_gate=False. confirm_and_execute already returns [] when the user
    # rejects, and returns the entries when a revert failed and the library is
    # still renamed — which is exactly when the audit trail matters most.
    if log:
        log_changes(excel_path, entries)
    return entries


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


# --- Inventory ---------------------------------------------------------------

def collect_inventory(root: Path) -> tuple[list[dict], list[str]]:
    """Every file under `root`, media or not. Returns (rows, walk errors).

    Strictly read-only, and it classifies rather than filters: nothing is
    left out for not looking like media.
    """
    rows: list[dict] = []
    errors: list[str] = []
    for dirpath, dirnames, filenames in os.walk(
            root, onerror=lambda e: errors.append(str(e))):
        dirnames[:] = sorted(dirnames)
        for fname in sorted(filenames):
            full = Path(dirpath) / fname
            try:
                st = full.stat()
                size, mtime = st.st_size, st.st_mtime
            except OSError:
                size, mtime = 0, 0
            try:
                rel = full.relative_to(root).as_posix()
            except ValueError:
                rel = str(full)
            suffix = full.suffix.lower()
            rows.append({
                'Path': rel,
                'Folder': Path(rel).parent.as_posix(),
                'File Name': fname,
                'Extension': suffix,
                'Type': ('video' if is_media_file(fname)
                         else 'companion' if suffix in COMPANION_EXTS
                         else 'other'),
                'Size (bytes)': size,
                'Size (MB)': round(size / (1024 ** 2), 2),
                'Modified': _fmt_ts(mtime) if mtime else '',
            })
    return rows, errors


def _inventory_tree(root: Path) -> list[str]:
    lines: list[str] = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(dirnames)
        try:
            depth = len(Path(dirpath).relative_to(root).parts)
        except ValueError:
            depth = 0
        lines.append("  " * depth + Path(dirpath).name + "/")
        for f in sorted(filenames):
            lines.append("  " * (depth + 1) + f)
    return lines


def write_inventory(root: Path, out: Path, dry_run: bool = False) -> bool:
    """Write an inventory of `root` to `out`. Format follows out's suffix."""
    rows, errors = collect_inventory(root)
    if errors:
        print(f"   [!] {len(errors)} directory error(s) during the walk — "
              f"the inventory may be incomplete.")
    if not rows:
        print("[!] No files found. The path may be empty or inaccessible.")
        print("    If this is a network share, check it is still mounted.")
        return False

    total_gb = sum(r['Size (bytes)'] for r in rows) / (1024 ** 3)
    video = sum(1 for r in rows if r['Type'] == 'video')
    if dry_run:
        print(f"\n[dry-run] Would write {len(rows)} file(s) "
              f"({total_gb:.2f} GB) to {out}")
        return True

    suffix = out.suffix.lower()
    try:
        if suffix == '.csv':
            import csv
            with open(out, 'w', newline='', encoding='utf-8-sig') as fh:
                writer = csv.DictWriter(fh, fieldnames=list(rows[0]))
                writer.writeheader()
                writer.writerows(rows)
        elif suffix == '.txt':
            out.write_text("\n".join(_inventory_tree(root)), encoding="utf-8")
        else:
            import pandas as pd
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                pd.DataFrame(rows).to_excel(writer, sheet_name='Inventory',
                                            index=False)
                excel._autosize_columns(writer)
    except PermissionError:
        print(f"[!] Cannot write '{out.name}' — close it in other programs first.")
        return False
    except OSError as exc:
        print(f"[!] Could not write '{out.name}': {exc}")
        return False

    print(f"\n[OK] Inventoried {len(rows)} file(s) ({total_gb:.2f} GB) — "
          f"{video} video, {len(rows) - video} other.")
    print(f"     Written to {out}")
    return True


def run_inventory() -> None:
    """List EVERY file under a folder — media or not — and change nothing.

    Distinct from [Scan library only], which understands media and writes the
    Movies/TV sheets the renamer reads. This one makes no judgement about
    what a file is: it is a plain record of what is on disk, which is what
    you want for an audit, a backup list, or a before/after comparison.
    """
    folder = browse_for_folder("Folder to inventory", allow_skip=False)
    if not folder:
        return
    print("\n  [1] Excel spreadsheet (.xlsx)  - one row per file, with sizes")
    print("  [2] Comma-separated values (.csv)")
    print("  [3] Plain text tree (.txt)")
    choice = prompt_input("Select (1-3) [1]: ", default='1')
    ext = {'2': '.csv', '3': '.txt'}.get(choice, '.xlsx')
    name = prompt_input(f"Output file [inventory{ext}]: ",
                        default=f"inventory{ext}")
    if not name.lower().endswith(ext):
        name += ext
    write_inventory(Path(folder), Path(name).resolve())


# --- Custom word list --------------------------------------------------------

def run_custom_words() -> None:
    """View, add and remove the words stripped from names before parsing.

    These were previously only reachable by hand-editing the JSON file:
    ``save_custom_patterns`` existed but nothing ever called it, so there was
    no way to add a word from the app and no way at all to take one back out.
    """
    def store(updated: list[str]) -> bool:
        """Persist the list, reporting rather than crashing the wizard.

        The app directory can legitimately be read-only — a system-wide
        install, a read-only container mount — and an unhandled OSError here
        would unwind all the way out of the menu.
        """
        try:
            save_custom_patterns(updated)
            return True
        except OSError as exc:
            print(f"   [!] Could not write {custom_patterns_path()}: {exc}")
            print("   Set MEDIAORG_PATTERNS to a writable location and "
                  "try again.")
            return False

    while True:
        patterns = load_custom_patterns()
        print("\n" + "-" * 60)
        print("CUSTOM WORD LIST")
        print("-" * 60)
        print(f"File: {custom_patterns_path()}")
        print("Anything matching one of these is removed from a name before\n"
              "it is parsed - release-group tags, tracker names, and so on.\n"
              "Each entry is a regular expression, matched case-insensitively.")
        if patterns:
            print()
            for i, pat in enumerate(patterns, 1):
                print(f"  [{i}] {pat}")
        else:
            print("\n  (the list is empty)")
        print("\n  [a] Add a word        [r] Remove one")
        print("  [c] Clear all         [t] Test against a filename")
        print("  [q] Back to the menu")

        action = prompt_input("\nSelect: ").lower()
        if action in ('q', ''):
            return

        if action == 'a':
            word = prompt_input("Word or pattern to strip: ")
            if not word:
                continue
            try:
                re.compile(word)
            except re.error as exc:
                print(f"   [!] Not a valid pattern: {exc}")
                print(f"   If you meant it literally, use: {re.escape(word)}")
                if not ask_yes_no("Add the escaped version instead?",
                                  default=True):
                    continue
                word = re.escape(word)
            if word in patterns:
                print("   [!] Already in the list.")
                continue
            # A pattern that eats a whole name is useless: pre_clean detects
            # that at runtime and declines to apply it, so the entry would sit
            # in the list doing nothing. Test the raw regex rather than
            # pre_clean, which would hide exactly the case being looked for.
            if not any(re.sub(word, '', probe, flags=re.IGNORECASE).strip()
                       for probe in ("The Matrix 1999 1080p",
                                     "Show S01E01 720p")):
                print("   [!] That pattern matches whole names, so it would "
                      "never be applied. Not added.")
                continue
            if store(patterns + [word]):
                print(f"   [OK] Added: {word}")

        elif action == 'r':
            if not patterns:
                print("   [!] Nothing to remove.")
                continue
            raw = prompt_input(f"Remove which (1-{len(patterns)}, or the word): ")
            if not raw:
                continue
            if raw.isdigit() and 1 <= int(raw) <= len(patterns):
                gone = patterns.pop(int(raw) - 1)
            elif raw in patterns:
                gone = raw
                patterns.remove(raw)
            else:
                print("   [!] No such entry.")
                continue
            if store(patterns):
                print(f"   [OK] Removed: {gone}")

        elif action == 'c':
            if not patterns:
                print("   [!] Already empty.")
                continue
            if ask_yes_no(f"Remove all {len(patterns)} entries?", default=False):
                if store([]):
                    print("   [OK] Cleared.")

        elif action == 't':
            sample = prompt_input("Filename to test: ")
            if not sample:
                continue
            parsed = parse_name(sample, custom_patterns=patterns)
            print(f"   after stripping : {pre_clean(sample, patterns)}")
            print(f"   parsed title    : {parsed.title}")
            print(f"   year / quality  : {parsed.year or '-'} / "
                  f"{parsed.quality or '-'}")


def confirm_word_list() -> list[str]:
    """Show the words that will be stripped, and offer to change them first.

    The list is what turns "The.Matrix.1999.1080p.YIFY" into a title, so the
    moment to look at it is on the way into a rename — not on a separate trip
    through menu [8] that you have to know to take.
    """
    patterns = load_custom_patterns()
    print("\n" + "-" * 60)
    print("WORDS STRIPPED FROM NAMES")
    print("-" * 60)
    if patterns:
        print("  " + ", ".join(patterns[:12])
              + (f"  ... and {len(patterns) - 12} more" if len(patterns) > 12
                 else ""))
    else:
        print("  (none - names are parsed exactly as they are)")
    if ask_yes_no("Edit this list before continuing?", default=False):
        run_custom_words()
        patterns = load_custom_patterns()
    return patterns


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


def run_update() -> bool:
    """Menu entry [U]: update the clone in place.

    Returns True once the files on disk have moved on from the modules this
    process already imported, so the caller can send the user back to a
    fresh launch rather than run half-old code.

    The answer comes from HEAD either side of the update, not from a status
    read beforehand: an offline status can be measuring refs that were
    already stale (a clone that has never fetched reads as up to date), and
    would then report "nothing changed" for an update that changed
    everything.
    """
    before = update.head_revision()
    update.run_update()
    moved = update.head_revision() != before
    if moved:
        update.begin_background_check(force=True)
    # Deliberately not keyed on the exit code: a pull that succeeded and a
    # dependency install that then failed is a non-zero exit *and* new files
    # on disk. Staying in the menu there is the worst of both.
    return moved


def _start_update_check() -> None:
    """Kick off the update check. Never lets it stop the app from starting."""
    deadline = time.monotonic() + LAUNCH_CHECK_BUDGET
    try:
        update.begin_background_check()
        if update.wait_for_cache(LAUNCH_CHECK_BUDGET) is None:
            # Nothing remembered from a previous launch (a fresh install, or
            # the first run after the cache was cleared): wait a moment so the
            # very first launch is the one that tells you an update exists.
            # One shared deadline, because a None above is ambiguous — it can
            # also mean the local phase simply has not finished yet, and
            # paying the budget twice is how "never slower" becomes four
            # seconds on a cold start behind an on-access virus scanner.
            update.wait_for_check(max(0.0, deadline - time.monotonic()))
    except Exception:
        pass                    # an update check is never worth a failed launch


def _update_notice(shown_already: bool) -> str:
    """The full notice the first time, a one-liner on every redraw after.

    Fourteen lines between the header and the menu, every single time you
    come back from a task, stops being information and becomes wallpaper.
    """
    try:
        status = update.latest_status()
        if status is None or not status.update_available:
            return ''
        if not shown_already:
            return update.banner(status)
        plural = '' if status.behind == 1 else 's'
        return (f"\n  [!] Update available ({status.behind} commit{plural} "
                f"behind) - press [U] to install it.")
    except Exception:
        return ''


def run_wizard() -> None:
    from . import __version__
    _start_update_check()
    notice_shown = False
    while True:
        try:
            print("\n" + "=" * 70)
            print(f"MEDIA ORGANIZER v{__version__}")
            print("=" * 70)
            notice = _update_notice(notice_shown)
            if notice:
                print(notice)
                notice_shown = True
            print("""
  [1] Clean file names        (scan -> preview -> rename)
  [2] Organize TV structure   (loose episodes -> Season folders)
  [3] Organize movie files    (loose files -> one folder per movie)
  [4] Do it all               (organize -> scan -> rename)
  [5] Fix file extensions     (restore missing / bulk convert)
  [6] Scan library only       (media only -> Excel, no changes made)
  [7] Inventory every file    (every file, media or not -> xlsx/csv/txt)
  [8] Custom word list        (add / remove words stripped from names)
  [9] Undo last run
  [U] Update Media Organizer  (git pull + dependencies)
  [0] Exit

Tip: type 'back' or 'b' at any prompt to return here.""")
            choice = prompt_input("\nSelect an option (0-9, U): ")

            if choice == '0':
                break
            elif choice.lower() == 'u':
                if run_update():
                    print("\nExiting so the new version is loaded on the "
                          "next launch.")
                    break
            elif choice == '9':
                run_list_runs()
                if pending_runs(_journal_path()):
                    run_undo(session=ask_yes_no(
                        "Undo the entire last action (all its runs)?",
                        default=False))
            elif choice == '8':
                run_custom_words()
            elif choice == '7':
                run_inventory()
            elif choice == '5':
                run_extension_fixer()
            elif choice == '3':
                movies = browse_for_folder("Select your Movies folder",
                                           allow_skip=False)
                if movies:
                    run_organize_movies(movies)
            elif choice == '2':
                _, tv = _ask_paths(tv_only=True)
                if tv:
                    run_organize(tv)
            elif choice in ('1', '4', '6'):
                movies, tv = _ask_paths()
                if not movies and not tv:
                    print("No folders selected.")
                    continue
                xlsx = _ask_excel_path()
                # One step per thing the user is asked about, so the word list
                # and the renaming engine are visibly part of the sequence
                # rather than surprises in the middle of it.
                # [4] organize, words, scan, rename; [1] drops the organize;
                # [6] stops after the scan.
                steps = {'4': 4, '1': 3, '6': 2}[choice]
                step = 1
                session = uuid.uuid4().hex[:12] if choice == '4' else None
                entries = []
                try:
                    if choice == '4':
                        if tv:
                            print(f"\n[{step}/{steps}] Organizing TV "
                                  f"structure...")
                            run_organize(tv, session=session, accept_gate=False)
                        if movies:
                            run_organize_movies(movies, session=session,
                                                accept_gate=False)
                        step += 1
                    else:
                        # A media scan only looks inside folders, so loose files
                        # in the movies root would go unmentioned otherwise.
                        _warn_loose_movies(movies)
                    print(f"\n[{step}/{steps}] Words to strip from names...")
                    confirm_word_list()
                    step += 1
                    print(f"\n[{step}/{steps}] Scanning library...")
                    run_scan(movies, tv, xlsx)
                    step += 1
                    if choice in ('1', '4'):
                        print(f"\n[{step}/{steps}] Renaming "
                              f"(you will be asked whether to use AI next)...")
                        print("(You can edit the '... Fixed' columns in the "
                              "spreadsheet first - reopen this step afterwards.)")
                        if ask_yes_no("Proceed to renaming now?", default=True):
                            # For [4] the keep-or-put-back question is asked
                            # once, below, for the whole action - so answering
                            # "no" reverses the organizing too, not just the
                            # renaming.
                            entries = run_rename(
                                movies, tv, xlsx, session=session,
                                accept_gate=(choice != '4'),
                                log=(choice != '4'))
                except BackNavigation:
                    # 'back' anywhere after the organize phase would otherwise
                    # skip the gate below, leaving those moves applied and the
                    # question never put. Fall through to it instead.
                    if session is None:
                        raise
                    print("\n(Going back - but the organizing already ran.)")
                except KeyboardInterrupt:
                    # An interrupt must not be answered with a prompt: the user
                    # is trying to stop, and reverting unasked is the one thing
                    # worse than not reverting. Say what is on disk and how to
                    # reverse it, then let the outer handler take it - silence
                    # was the actual defect, not the absence of a rollback.
                    if session is not None:
                        _report_unaccepted(_journal_path(), session)
                    raise
                except Exception:
                    # Same guarantee, for everything else. run_scan and
                    # excel.read_library can raise past their own handling, and
                    # without this the process died with the organize phase
                    # applied, no gate, and nothing said - the exact failure
                    # this whole flow exists to prevent.
                    if session is not None:
                        _report_unaccepted(_journal_path(), session)
                    raise
                if session is not None:
                    outcome = accept_or_revert(None, _journal_path(),
                                               label="changes", session=session)
                    if outcome is GateOutcome.KEPT:
                        log_changes(xlsx, entries)
                    elif outcome is GateOutcome.REVERTED:
                        _warn_stale_spreadsheet(xlsx)
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

    # No abbreviations: run.py dispatches --update/--version before the
    # dependency install, and it cannot resolve prefixes the way argparse
    # does. Rejecting "--upd" in both places beats one honouring it and
    # the other silently ignoring it.
    parser = argparse.ArgumentParser(description="Media Organizer",
                                     allow_abbrev=False)
    parser.add_argument('--action', choices=['scan', 'organize',
                                             'organize-movies', 'rename',
                                             'full', 'inventory'])
    parser.add_argument('--movies', help="Path to movies folder")
    parser.add_argument('--tv', help="Path to TV shows folder")
    parser.add_argument('--output', help="Excel file name")
    parser.add_argument('--path', help="Folder to inventory (--action inventory)")
    parser.add_argument('--dry-run', action='store_true',
                        help="Show planned changes without touching disk")
    parser.add_argument('--review', action='store_true',
                        help="After applying, show what changed and ask "
                             "whether to keep it; answering no puts every "
                             "file back. Without this, --action runs keep "
                             "their changes and print the undo command.")
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
    parser.add_argument('--update', action='store_true',
                        help="Update to the latest version (git pull + dependencies)")
    parser.add_argument('--check-update', action='store_true',
                        help="Report how many commits behind this copy is, "
                             "without changing anything")
    parser.add_argument('--yes', '-y', action='store_true',
                        help="Answer yes to the --update confirmation")
    parser.add_argument('--version', '-V', action='store_true',
                        help="Print the version and installed commit")
    args = parser.parse_args()

    if args.version:
        update.print_version()
        return
    if args.check_update:
        print(update.describe(update.check_and_cache(fetch=True)))
        return
    if args.update:
        sys.exit(update.run_update(assume_yes=args.yes, dry_run=args.dry_run))

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
    if args.action == 'inventory':
        target = args.path or args.movies or args.tv
        if not target:
            sys.exit("[!] --path is required for 'inventory'.")
        out = Path(args.output or "inventory.xlsx").resolve()
        write_inventory(Path(target).resolve(), out, dry_run=args.dry_run)
        return
    # Scripted runs apply what they were asked to apply and print the undo
    # command; only --review adds the keep-or-put-back question, which nothing
    # unattended could answer. The pre-apply review runs either way.
    gate = args.review
    if args.action == 'organize':
        if not tv:
            sys.exit("[!] --tv is required for 'organize'.")
        run_organize(tv, dry_run=args.dry_run, accept_gate=gate)
    elif args.action == 'organize-movies':
        if not movies:
            sys.exit("[!] --movies is required for 'organize-movies'.")
        run_organize_movies(movies, dry_run=args.dry_run, accept_gate=gate)
    elif args.action == 'scan':
        run_scan(movies, tv, xlsx, dry_run=args.dry_run)
    elif args.action == 'rename':
        run_rename(movies, tv, xlsx, dry_run=args.dry_run, accept_gate=gate)
    elif args.action == 'full':
        # One session id across all three phases so a single
        # `--undo-session` reverses the whole action.
        session = uuid.uuid4().hex[:12]
        if tv:
            run_organize(tv, dry_run=args.dry_run, session=session,
                         accept_gate=False)
        if movies:
            loose = plan_loose_movies(Path(movies))
            if loose.ops:
                confirm_and_execute(loose, _journal_path(), args.dry_run,
                                    "loose-file moves", roots=[Path(movies)],
                                    session=session, accept_gate=False)
        run_scan(movies, tv, xlsx, dry_run=args.dry_run)
        # With --review, one question covers every phase: answering "no" puts
        # the organize back too, not just the rename.
        entries = run_rename(movies, tv, xlsx, dry_run=args.dry_run,
                             session=session, accept_gate=False, log=not gate)
        if args.dry_run:
            return
        if gate:
            outcome = accept_or_revert(None, _journal_path(), label="changes",
                                       session=session)
            if outcome is GateOutcome.KEPT:
                log_changes(xlsx, entries)
            elif outcome is GateOutcome.REVERTED:
                _warn_stale_spreadsheet(xlsx)
        else:
            # run_rename already logged (log=not gate), so just say how to undo.
            print("\nUndo this entire action:")
            for cmd in _undo_commands(_journal_path(), session, None):
                print(f"    {cmd}")


if __name__ == "__main__":
    main()
