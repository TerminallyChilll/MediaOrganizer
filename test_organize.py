#!/usr/bin/env python3
"""
Unit tests for the organize_season_structure and _folder_has_episodes_or_seasons functions.
Run with: python test_organize.py
"""
import os
import sys
import re
import tempfile
import shutil

# ---------------------------------------------------------------------------
# Inline the two functions under test so this file is self-contained
# ---------------------------------------------------------------------------

def _folder_has_episodes_or_seasons(path):
    """Check if a folder directly contains episode folders (SxxExx) or season folders."""
    try:
        folders = [f for f in os.listdir(path) if os.path.isdir(os.path.join(path, f))]
    except Exception:
        return False
    ep_pat = re.compile(r'S\d{1,2}E\d{1,2}', re.IGNORECASE)
    season_pat = re.compile(r'^(?:season\s*|s)\d{1,2}(?:\s*\(?\d{4}\)?)?\s*$', re.IGNORECASE)
    season_like_pat = re.compile(
        r'^([^\-\|.]+?)\s+(?:season\s*|s)\d{1,2}(?:\s*\(?\d{4}\)?)?\s*$', re.IGNORECASE
    )
    def _is_season_like(f):
        m = season_like_pat.match(f)
        return bool(m and len(m.group(1).strip().split()) <= 3)
    return any(ep_pat.search(f) or season_pat.search(f) or _is_season_like(f) for f in folders)


def organize_season_structure(show_path):
    """Organize a TV show folder by grouping loose episode folders and normalizing season names."""
    changes = []
    try:
        items = os.listdir(show_path)
    except Exception:
        return changes

    folders = [f for f in items if os.path.isdir(os.path.join(show_path, f))]

    episode_folder_pattern = re.compile(r'S(\d{1,2})E\d{1,2}', re.IGNORECASE)
    season_folder_pattern = re.compile(r'^(?:season\s*|s)(\d{1,2})$', re.IGNORECASE)
    season_like_pattern = re.compile(
        r'^([^\-\|.]+?)\s+(?:season\s*|s)(\d{1,2})(?:\s*\(?\d{4}\)?)?\s*$', re.IGNORECASE
    )

    existing_seasons = {}
    loose_episodes = {}

    for folder in folders:
        season_match = season_folder_pattern.match(folder.strip())
        if season_match:
            existing_seasons[int(season_match.group(1))] = folder
            continue
        ep_match = episode_folder_pattern.search(folder)
        if ep_match:
            snum = int(ep_match.group(1))
            loose_episodes.setdefault(snum, []).append(folder)
            continue
        s_match = season_like_pattern.match(folder)
        if s_match and not episode_folder_pattern.search(folder):
            prefix = s_match.group(1).strip()
            if len(prefix.split()) <= 3:
                existing_seasons[int(s_match.group(2))] = folder

    for snum, ep_folders in sorted(loose_episodes.items()):
        target_season = f"Season {snum}"
        target_path = os.path.join(show_path, target_season)
        for ep_folder in sorted(ep_folders):
            changes.append({
                'type': 'move_to_season',
                'old_path': os.path.join(show_path, ep_folder),
                'new_path': os.path.join(target_path, ep_folder),
                'description': f"  📦 {ep_folder}  →  {target_season}/{ep_folder}"
            })

    for snum, folder_name in sorted(existing_seasons.items()):
        target_name = f"Season {snum}"
        if folder_name != target_name:
            changes.append({
                'type': 'rename_season',
                'old_path': os.path.join(show_path, folder_name),
                'new_path': os.path.join(show_path, target_name),
                'description': f"  ✏️  {folder_name}  →  {target_name}"
            })

    return changes


# ---------------------------------------------------------------------------
# Test helpers
# ---------------------------------------------------------------------------

PASS = "[PASS]"
FAIL = "[FAIL]"
results = []

def make_dirs(base, *names):
    """Create subdirectories inside base."""
    for name in names:
        os.makedirs(os.path.join(base, name), exist_ok=True)

def run_test(name, actual, expected):
    ok = actual == expected
    tag = PASS if ok else FAIL
    print(f"{tag}  {name}")
    if not ok:
        print(f"       Expected: {expected}")
        print(f"       Got:      {actual}")
    results.append(ok)


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

def test_watch_folder_not_treated_as_season():
    """WATCH - The Office Extended 9x9 - FREE should NOT be identified as Season 9."""
    with tempfile.TemporaryDirectory() as tmp:
        # Simulate TV root containing show-style folders that contain digits after 'x'
        make_dirs(tmp,
            "WATCH - The Office Extended 9x9 - FREE",
            "WATCH - The Office Extended 9x8 - FREE",
            "Deleted Scenes Season 9",
            "Bloopers Season 8",
        )
        changes = organize_season_structure(tmp)
        # None of these folders should trigger any moves or renames
        run_test(
            "WATCH-style folders are NOT treated as seasons",
            len(changes), 0
        )


def test_pure_season_folders_get_normalized():
    """S01, S02 style folders get renamed to Season 1, Season 2."""
    with tempfile.TemporaryDirectory() as tmp:
        make_dirs(tmp, "S01", "S02", "S03")
        changes = organize_season_structure(tmp)
        rename_changes = [c for c in changes if c['type'] == 'rename_season']
        run_test("S01/S02/S03 produce 3 rename changes", len(rename_changes), 3)
        targets = sorted(os.path.basename(c['new_path']) for c in rename_changes)
        run_test("Target names are Season 1/2/3", targets, ["Season 1", "Season 2", "Season 3"])


def test_show_name_season_folders_get_normalized():
    """Snowfall S02 style folders get renamed to Season 2."""
    with tempfile.TemporaryDirectory() as tmp:
        make_dirs(tmp, "Snowfall S01", "Snowfall S02", "Snowfall S03")
        changes = organize_season_structure(tmp)
        rename_changes = [c for c in changes if c['type'] == 'rename_season']
        run_test("Snowfall S01/S02/S03 produce 3 rename changes", len(rename_changes), 3)


def test_loose_episode_folders_get_grouped():
    """Loose episode folders like S01E01_FolderName get moved into Season 1."""
    with tempfile.TemporaryDirectory() as tmp:
        make_dirs(tmp,
            "S01E01 Pilot",
            "S01E02 Dundies",
            "S02E01 New Season",
        )
        changes = organize_season_structure(tmp)
        move_changes = [c for c in changes if c['type'] == 'move_to_season']
        run_test("3 loose episode folders produce 3 moves", len(move_changes), 3)
        # S01 episodes should go into Season 1
        s1_moves = [c for c in move_changes if "Season 1" in c['new_path']]
        run_test("S01 episodes moved to Season 1", len(s1_moves), 2)
        s2_moves = [c for c in move_changes if "Season 2" in c['new_path']]
        run_test("S02 episodes moved to Season 2", len(s2_moves), 1)


def test_already_organized_produces_no_changes():
    """A correctly organized show (Season 1, Season 2 folders) produces no changes."""
    with tempfile.TemporaryDirectory() as tmp:
        make_dirs(tmp, "Season 1", "Season 2", "Season 3")
        changes = organize_season_structure(tmp)
        run_test("Already-organized show produces 0 changes", len(changes), 0)


def test_folder_has_episodes_detection():
    """_folder_has_episodes_or_seasons correctly detects season-containing folders."""
    with tempfile.TemporaryDirectory() as tmp:
        # Case 1: contains episode folders — should return True
        with tempfile.TemporaryDirectory() as show_dir:
            make_dirs(show_dir, "S01E01 Pilot", "S01E02 Second Episode")
            run_test("Folder with SxxExx episode folders detected", _folder_has_episodes_or_seasons(show_dir), True)

        # Case 2: contains Season folders — should return True
        with tempfile.TemporaryDirectory() as show_dir:
            make_dirs(show_dir, "Season 1", "Season 2")
            run_test("Folder with Season N folders detected", _folder_has_episodes_or_seasons(show_dir), True)

        # Case 3: WATCH-style show-root folder — should return False
        with tempfile.TemporaryDirectory() as root_dir:
            make_dirs(root_dir,
                "WATCH - The Office Extended 9x9 - FREE",
                "WATCH - Bloopers - Season 8 - FREE",
                "Deleted Scenes Season 9"
            )
            run_test(
                "WATCH-style root folder NOT detected as having seasons",
                _folder_has_episodes_or_seasons(root_dir), False
            )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("  MEDIA ORGANIZER — organize_season_structure() Tests")
    print("=" * 60 + "\n")

    test_watch_folder_not_treated_as_season()
    test_pure_season_folders_get_normalized()
    test_show_name_season_folders_get_normalized()
    test_loose_episode_folders_get_grouped()
    test_already_organized_produces_no_changes()
    test_folder_has_episodes_detection()

    print("\n" + "=" * 60)
    passed = sum(results)
    total = len(results)
    print(f"  Results: {passed}/{total} passed")
    print("=" * 60 + "\n")

    sys.exit(0 if passed == total else 1)
