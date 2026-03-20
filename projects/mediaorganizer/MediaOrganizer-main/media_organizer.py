#!/usr/bin/env python3
"""
UNIFIED MEDIA ORGANIZER
=======================
A cohesive tool to scan media libraries and rename files beautifully.
Combines both the Scanner and Renamer functionalities.

Features:
- Scans library to Excel
- Interactive naming scheme builder (persists via config)
- Auto-generates undo scripts
- Executes renaming safely
"""

import os
import re
import json
import sys
import shutil
from pathlib import Path

try:
    from send2trash import send2trash as _send2trash
    def safe_delete(path: str) -> None:
        """Send a file or folder to the system trash (Windows Recycle Bin / macOS Trash / Linux Trash)."""
        _send2trash(path)
except ImportError:
    # Fallback: permanent delete if send2trash is somehow not installed
    def safe_delete(path: str) -> None:  # type: ignore[misc]
        if os.path.isdir(path):
            os.rmdir(path)
        else:
            os.remove(path)
from datetime import datetime
from collections import defaultdict
import builtins
import uuid

# Global Emoji State and Wrappers
USE_EMOJIS = True

def _strip_emojis(text: str) -> str:
    if not isinstance(text, str):
        return text
    # Strip emojis and pictographs using a regex covering the main unicode ranges:
    import re
    text = re.sub(
        r'[\U0001f300-\U0001faff\U0001f600-\U0001f64f\U0001f680-\U0001f6ff\u2600-\u27bf\u2300-\u23ff\u25a0-\u25ff\u2b50\u203c-\u3299\xae\xa9\u2122\u2139\U0001f004-\U0001f251\ufe0f]+', 
        '', text
    )
    # Also clean up double spaces that might be left behind but preserve leading/trailing whitespace
    return re.sub(r'(?<=\S)  +(?=\S)', ' ', text)

_builtin_print = builtins.print
_builtin_input = builtins.input

def print(*args, **kwargs):
    if USE_EMOJIS:
        _builtin_print(*args, **kwargs)
    else:
        new_args = [_strip_emojis(a) if isinstance(a, str) else a for a in args]
        _builtin_print(*new_args, **kwargs)

def input(prompt=''):
    if USE_EMOJIS:
        return _builtin_input(prompt)
    else:
        return _builtin_input(_strip_emojis(prompt))

try:
    import pandas as pd # type: ignore
    from tqdm import tqdm # type: ignore
except ImportError:
    print("Missing required libraries. Please run: pip install pandas openpyxl tqdm")
    sys.exit(1)

try:
    from llm_cleaner import clean_titles_with_llm, load_llm_config, save_llm_config, list_ollama_models # type: ignore
    LLM_AVAILABLE = True
except ImportError:
    LLM_AVAILABLE = False

# =========================
# CONFIGURATION
# =========================
DEFAULT_EXCEL_NAME = "media_library_scan.xlsx"
CUSTOM_PATTERNS_FILE = "custom_strip_patterns.json"
CONFIG_FILE_NAME = ".media_renamer_config.json"
VIDEO_EXTENSIONS = ['.mkv', '.mp4', '.avi', '.mov', '.m4v', '.wmv', '.flv', '.webm', '.ts', '.mpg', '.mpeg']

DEFAULT_STRIP_PATTERNS = [
    r'\d+\.?\d*\s?(MB|GB|mb|gb)', r'\d{3,4}MB',
    r'\b(x264|x265|h\.?264|h\.?265|hevc|xvid|divx|avc)\b',
    r'\b(AAC|AC3|DTS|DD5\.1|DD5|DD2\.0|DD2|FLAC|MP3|Atmos)\b',
    r'\bDDP?\s*\d+[\s\.]\d+\b',  # DDP5.1, DD5.1, DDP5 1, DD5 1, etc. (dot or space separated)
    r'\b(WEBRip|Web-Rip|WebRip|WEB-DL|WebDL|WEBDL|WEB|BluRay|Blu-Ray|BRRip|BDRip|DVDRip|HDTV|PDTV)\b',
    r'\b(PROPER|REPACK|INTERNAL|LIMITED|UNRATED|EXTENDED|DC|DIRECTORS\.CUT)\b',
    r'\[(?![12]\d{3}\])[^\]]*\]',
    r'\b(Subs|Subtitle|Subtitles|iP|RARBG|YIFY|YTS)\b',
    r'\b(GalaxyRG|RARBG|YTS|YIFY|ETRG|Pahe|PSA|STUTTERSHIT|CMRG|TGx|EVO|ION10|ION265)\b',
    # Streaming service source tags
    r'\b(PCOK|AMZN|DSNP|HMAX|ATVP|PMTP|NFLX|CRKL|STAN|BCORE|iP)\b',
    # Common release group names
    r'\b(FLUX|NTb|EDITH|MeGusta|mSD|DEFLATE|NTG|Pahe|TOMMY|MIXED|EZTVx?|RAWR|JFF|RiPSaL|CMRG|SiGMA)\b',
    # Website/URL prefixes (e.g. "www.UIndex.org    -    filename")
    r'(?:www\.\S+|https?://\S+)\s*[-–—]\s*',
    r'\.mp4$|\.mkv$|\.avi$|\.ts$|\.m4v$|\.wmv$|\.mov$', r'&#?\w+;', r'\b(Phoenix\s*RG|MVGroup\.org|UIndex\.org)\b',
    # Streaming site file-type artifacts embedded in filename (e.g. "Show Ep Ts.ts" → "Ts" leftover)
    r'\bTs\b',
]

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
        return self.__dict__

    @classmethod
    def from_dict(cls, data):
        scheme = cls()
        for k, v in data.items():
            if hasattr(scheme, k):
                setattr(scheme, k, v)
        return scheme


# =========================
# SHARED UTILITIES
# =========================

class BackNavigationException(Exception):
    pass

def prompt_input(message, default=''):
    """Universal prompt that supports 'back' or 'b'."""
    val = input(f"{message}").strip()
    if val.lower() in ['back', 'b']:
        raise BackNavigationException()
    return val if val else default

def ask_yes_no(prompt, default=True):
    default_str = "y/n"
    response = prompt_input(f"{prompt} ({default_str}): ")
    if not response:
        return default
    return response.upper() in ['Y', 'YES']

def paginated_preview(lines, page_size=20):
    """
    Show a list of lines with page-by-page navigation.
    Returns True if the user wants to proceed, False to abort.
    """
    total = len(lines)
    if total == 0:
        return True

    page = 0
    total_pages = max(1, (total + page_size - 1) // page_size)

    while True:
        start = page * page_size
        end = min(start + page_size, total)
        print(f"\n--- Page {page + 1}/{total_pages}  (items {start + 1}-{end} of {total}) ---")
        for line in lines[start:end]:
            print(line)

        # Navigation options
        at_start = page == 0
        last_page: int = int(total_pages) - 1
        at_end   = page == last_page

        nav_parts = []
        if not at_end:  nav_parts.append("[N]ext")
        if not at_start: nav_parts.append("[P]rev")
        if total_pages > 2: nav_parts.append("[G]o to page")
        if not at_end:  nav_parts.append("[A]ll at once")
        nav_parts.append("[Y]es proceed")
        nav_parts.append("[Q]uit / abort")
        print("  " + "  ".join(nav_parts))

        choice = input("Choice: ").strip().upper()

        if choice in ('N', '') and not at_end:
            page += 1
        elif choice == 'P' and not at_start:
            page -= 1
        elif choice == 'G' and total_pages > 2:
            try:
                pg = int(input(f"Go to page (1-{total_pages}): ").strip()) - 1
                if 0 <= pg < total_pages:
                    page = pg
            except ValueError:
                pass
        elif choice == 'A':
            print(f"\n--- ALL {total} items ---")
            for line in lines:
                print(line)
            choice = input("\n[Y]es proceed  [Q]uit / abort: ").strip().upper()
            if choice == 'Y': return True
            if choice == 'Q': return False
        elif choice == 'Y':
            return True
        elif choice == 'Q':
            return False
        elif at_end and choice not in ('P', 'G', 'A', 'Q'):
            # At last page, any unrecognised input = prompt clearly
            pass

def clean_path_input(path):
    if not path: return path
    path = str(path).strip()
    if (path.startswith('"') and path.endswith('"')) or (path.startswith("'") and path.endswith("'")):
        path = path.strip('"').strip("'")
    return path

def validate_path(path):
    if not path: return None
    cleaned = clean_path_input(path)
    if not cleaned: return None
    try:
        p = Path(str(cleaned))
        if p.exists() and p.is_dir():
            return str(p)
    except Exception: pass
    return None

def browse_for_folder(prompt="Select a folder", allow_skip=True):
    """User-friendly folder selection with GUI picker, CLI browser, and manual paste."""
    print(f"\n📁 {prompt}")
    print("-" * 50)
    print("  [1] 📂 Open folder picker (Windows dialog)")
    print("  [2] 📂 Browse folders in terminal")
    print("  [3] ✏️  Paste a path manually")
    if allow_skip:
        print("  [4] ⏭️  Skip (leave empty)")
    else:
        print("  [4] 🔙  Cancel / Go Back")
    
    choice = input("\nSelect (1-4): ").strip()
    
    if choice == '1':
        # Try tkinter folder dialog
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            folder = filedialog.askdirectory(title=prompt)
            root.destroy()
            if folder:
                print(f"   ✅ Selected: {folder}")
                return folder
            else:
                print("   ⚠️ No folder selected.")
                return browse_for_folder(prompt, allow_skip)
        except Exception as e:
            print(f"   ⚠️ Could not open folder picker: {e}")
            print("   Falling back to terminal browser...")
            return _cli_folder_browser(prompt, allow_skip)
    
    elif choice == '2':
        return _cli_folder_browser(prompt, allow_skip)
    
    elif choice == '3':
        path = input("\nPaste folder path: ").strip()
        result = validate_path(path)
        if result:
            print(f"   ✅ Valid: {result}")
            return result
        else:
            print("   ❌ Invalid path. Try again.")
            return browse_for_folder(prompt, allow_skip)
    
    elif choice == '4':
        return None
    
    else:
        return browse_for_folder(prompt, allow_skip)

def _cli_folder_browser(prompt, allow_skip=True):
    """Interactive terminal folder browser."""
    # Start at a reasonable location
    current = Path.home()
    
    while True:
        print(f"\n📂 Current: {current}")
        print("-" * 50)
        
        try:
            subfolders = sorted([d for d in current.iterdir() if d.is_dir() and not d.name.startswith('.')], 
                               key=lambda x: x.name.lower())
        except PermissionError:
            print("   ❌ Permission denied. Going back up...")
            current = current.parent
            continue
        
        # Show options
        print("  [✓] Select THIS folder")
        if current.parent != current:  # Not root
            print("  [..] Go up one level")
        
        for i, folder in enumerate(subfolders, 1):
            if i > 20: break
            print(f"  [{i}] 📁 {folder.name}")
        if len(subfolders) > 20:
            print(f"  ... and {len(subfolders) - 20} more folders")
        
        if allow_skip:
            print("  [q] Cancel / Skip")
        else:
            print("  [q] Cancel / Go Back")
        
        nav = input("\nNavigate to: ").strip().lower()
        
        if nav == 'q':
            return None
        elif nav in ['v', 'V', 'ok', 'OK', '']:
            print(f"   ✅ Selected: {current}")
            return str(current)
        elif nav == '..':
            current = current.parent
        else:
            try:
                idx = int(nav) - 1
                if 0 <= idx < len(subfolders):
                    current = subfolders[idx]
                else:
                    print("   Invalid number.")
            except ValueError:
                # Maybe they typed a path directly
                typed = validate_path(nav)
                if typed:
                    print(f"   ✅ Selected: {typed}")
                    return typed
                print("   Invalid input. Type a number, '..', or '✓'.")

def get_folder_size(folder_path):
    total_size = 0
    try:
        for dirpath, dirnames, filenames in os.walk(folder_path):
            for filename in filenames:
                try: total_size += os.path.getsize(os.path.join(dirpath, filename))
                except Exception: pass
    except Exception: pass
    return round(float(total_size) / (1024**3), 2)  # type: ignore

def extract_year(text):
    matches = re.findall(r'\b(19\d{2}|20\d{2})\b', text)
    if matches: return matches[0]
    return None

def extract_quality(text):
    text_lower = text.lower()
    quality_map = [
        ('8K', ['8k', '4320p']), ('4K', ['4k', '2160p', 'uhd']), ('2K', ['2k', '1440p', 'qhd']),
        ('1080p', ['1080p', '1080', 'fhd', 'fullhd', 'full-hd']), ('720p', ['720p', '720', 'hd']),
        ('480p', ['480p', '480', 'dvd']), ('360p', ['360p', '360']),
    ]
    for quality_name, patterns in quality_map:
        for pattern in patterns:
            if pattern in text_lower: return quality_name
    return None


# =========================
# SCANNER LOGIC
# =========================

def load_custom_patterns():
    if os.path.exists(CUSTOM_PATTERNS_FILE):
        try:
            with open(CUSTOM_PATTERNS_FILE, 'r') as f: return json.load(f)
        except Exception: return []
    return []

def save_custom_patterns(patterns):
    try:
        with open(CUSTOM_PATTERNS_FILE, 'w') as f:
            json.dump(patterns, f, indent=2)
    except Exception as e:
        print(f"⚠️ Failed to save custom patterns: {e}")

def get_custom_patterns(target_type="patterns"):
    print("\n" + "=" * 80)
    print(f"🎨 CUSTOM {target_type.upper()} CLEANING")
    print("=" * 80)
    if target_type.lower() == "folder":
        print("We are now gathering words/patterns to strip from your FOLDER names.")
    elif target_type.lower() == "file":
        print("We are now gathering words/patterns to strip from your FILE names.")
        
    existing_patterns = load_custom_patterns()
    custom_patterns: list[str] = []
    
    if existing_patterns:
        print(f"\n📋 You have {len(existing_patterns)} saved custom patterns from previous runs.")
        print("Current patterns:", ", ".join(existing_patterns))
        if ask_yes_no("Use these patterns for this step?", default=True):
            custom_patterns = list(existing_patterns)
            if not ask_yes_no("Would you like to add more words to this list?", default=False):
                case_sensitive = ask_yes_no("Should pattern matching be case-sensitive?", default=False)
                return custom_patterns, case_sensitive
            print("\nType one pattern to strip at a time. Type 'NOTHING' or press Enter to finish.")
        else:
            print("\nType one pattern to strip at a time. Type 'NOTHING' or press Enter to finish.")
    else:
        print("\nType one pattern to strip at a time. Type 'NOTHING' or press Enter to finish.")
    
    while True:
        pattern = prompt_input("Pattern to strip: ")
        if not pattern or pattern.upper() == 'NOTHING': break
        if pattern not in custom_patterns:
            custom_patterns.append(pattern)
            print(f"✅ Added: {pattern}")
    
    if custom_patterns:
        save_custom_patterns(custom_patterns)
        print(f"💾 Patterns saved for future runs.")
    
    case_sensitive = ask_yes_no("Should pattern matching be case-sensitive?", default=False)
    return custom_patterns, case_sensitive

def clean_title(folder_name, custom_patterns, case_sensitive=False):
    title = folder_name
    year = extract_year(title)
    quality_patterns = [
        r'\b8K\b', r'\b4K\b', r'\b2K\b', r'\b8k\b', r'\b4k\b', r'\b2k\b',
        r'\b4320p\b', r'\b2160p\b', r'\b1440p\b', r'\b1080p\b', r'\b720p\b', r'\b480p\b', r'\b360p\b',
        r'\b4320P\b', r'\b2160P\b', r'\b1440P\b', r'\b1080P\b', r'\b720P\b', r'\b480P\b', r'\b360P\b',
        r'\b1080\b', r'\b720\b', r'\b480\b', r'\b360\b', r'\b(UHD|FHD|HD|QHD)\b',
    ]
    for pattern in quality_patterns: title = re.sub(pattern, '', title, flags=re.IGNORECASE)
    for pattern in DEFAULT_STRIP_PATTERNS: title = re.sub(pattern, '', title, flags=re.IGNORECASE)
    
    spaced_patterns = [r'\bH\s*264\b', r'\bH\s*265\b', r'\bAAC2?\s*0\b', r'\bDD5?\s*1\b', r'\bDD2?\s*0\b', r'\bDDP?\s*\d+[\s\.]\s*\d+\b']
    for pattern in spaced_patterns: title = re.sub(pattern, '', title, flags=re.IGNORECASE)
    flags = 0 if case_sensitive else re.IGNORECASE
    for pattern in custom_patterns: title = re.sub(re.escape(pattern), '', title, flags=flags)
    
    title = re.sub(r'[\._\-]', ' ', title)
    title = re.sub(r'\s+', ' ', title)
    title = title.replace('039;', "'").replace('&amp;', '&').replace('&quot;', '"')
    title = re.sub(r'\[.*?\]', '', title)
    title = re.sub(r'\((?!\d{4}\))[^)]*\)', '', title)
    title = re.sub(r'\s+', ' ', title).strip()
    title = re.sub(r'\(\s*\)', '', title)
    title = re.sub(r'\s+', ' ', title).strip().title()
    
    # Always ensure the final name starts with a letter — strip any leading
    # punctuation/symbols/spaces left behind after pattern removal (e.g. "- My Show" -> "My Show")
    title = re.sub(r'^[^A-Za-z]+', '', title).strip()
    
    if year:
        title = re.sub(r'\b' + year + r'\b', '', title).strip()
        title = re.sub(r'\(\s*\)', '', title)
        title = re.sub(r'\s+', ' ', title).strip()
        # Strip again after year removal in case year was at the start
        title = re.sub(r'^[^A-Za-z]+', '', title).strip()
        title = f"{title} ({year})"
    return title

def extract_season_episode(filename):
    patterns = [r'[Ss](\d{1,2})[Ee](\d{1,2})', r'(\d{1,2})[xX](\d{1,2})']
    for pattern in patterns:
        match = re.search(pattern, filename)
        if match: return int(match.group(1)), int(match.group(2))
    return None, None

def scan_tv_show_seasons(show_path):
    episodes = []
    try: items = os.listdir(show_path)
    except Exception: return episodes
    
    season_folders = [item for item in items if os.path.isdir(os.path.join(show_path, item)) and re.search(r'season|s\d{1,2}', item, re.IGNORECASE)]
    
    if not season_folders:
        for item in items:
            item_path = os.path.join(show_path, item)
            if os.path.isfile(item_path) and os.path.splitext(item)[1].lower() in VIDEO_EXTENSIONS:
                season, episode = extract_season_episode(item)
                if season and episode:
                    episodes.append({
                        'season_folder': '', 'season_num': season, 'episode_num': episode,
                        'filename': item, 'year': extract_year(item),
                        'size': os.path.getsize(item_path) / (1024**3),
                        'rel_path': item
                    })
        return episodes
    
    # Also collect loose video files sitting directly in show_path (even when season/episode folders exist)
    seen_rel_paths = set()
    for item in items:
        item_path = os.path.join(show_path, item)
        if os.path.isfile(item_path) and os.path.splitext(item)[1].lower() in VIDEO_EXTENSIONS:
            s, e = extract_season_episode(item)
            if s and e and item not in seen_rel_paths:
                seen_rel_paths.add(item)
                episodes.append({
                    'season_folder': '', 'season_num': s, 'episode_num': e,
                    'filename': item, 'year': extract_year(item),
                    'size': os.path.getsize(item_path) / (1024**3),
                    'rel_path': item
                })

    for season_folder in season_folders:
        season_path = os.path.join(show_path, season_folder)
        season_match = re.search(r'(\d{1,2})', season_folder)
        season_num = int(season_match.group(1)) if season_match else None
        season_year = None
        season_episodes = []
        try:
            # Walk all depths below the season folder to find video files
            for dirpath, dirnames, filenames in os.walk(season_path):
                for file in filenames:
                    file_path = os.path.join(dirpath, file)
                    if os.path.splitext(file)[1].lower() in VIDEO_EXTENSIONS:
                        s, e = extract_season_episode(file)
                        if s and e:
                            file_year = extract_year(file)
                            if file_year and not season_year: season_year = file_year
                            # Store relative path from season folder
                            rel_path = os.path.relpath(file_path, os.path.join(show_path))
                            if rel_path not in seen_rel_paths:
                                seen_rel_paths.add(rel_path)
                                season_episodes.append({
                                    'season_folder': season_folder, 'season_num': s, 'episode_num': e,
                                    'filename': file, 'year': file_year,
                                    'size': os.path.getsize(file_path) / (1024**3),
                                    'rel_path': rel_path
                                })
        except Exception: pass
        for ep in season_episodes:
            if not ep['year']: ep['year'] = season_year
            episodes.append(ep)
    return episodes

def organize_loose_files(media_path, media_type="Movies"):
    """Find video files sitting directly in the root media folder and move them into their own folders."""
    try:
        loose_files = [f for f in os.listdir(media_path) 
                       if os.path.isfile(os.path.join(media_path, f)) 
                       and os.path.splitext(f)[1].lower() in VIDEO_EXTENSIONS]
    except Exception:
        return 0
    
    if not loose_files:
        return 0
        
    print(f"\n📂 Found {len(loose_files)} loose {media_type.lower()} file(s) not in folders. Organizing...")
    moved = 0
    for filename in loose_files:
        folder_name = os.path.splitext(filename)[0]
        folder_path = os.path.join(media_path, folder_name)
        file_path = os.path.join(media_path, filename)
        try:
            os.makedirs(folder_path, exist_ok=True)
            shutil.move(file_path, os.path.join(folder_path, filename))
            moved += 1
        except Exception as e:
            print(f"   ⚠️ Could not move '{filename}': {e}")
    
    if moved:
        print(f"   ✅ Organized {moved} file(s) into folders.")
    return moved


def organize_season_structure(show_path):
    """
    Organize a TV show folder by:
    1. Grouping loose episode folders (SxxExx) into Season X folders
    2. Normalizing season folder names (S01 -> Season 1)
    Returns list of changes made: [{type, old_path, new_path, description}]
    """
    changes = []
    try:
        items = os.listdir(show_path)
    except Exception:
        return changes
    
    folders = [f for f in items if os.path.isdir(os.path.join(show_path, f))]
    
    # Step 1: Identify loose episode folders — match both SxxExx AND bare NxN (e.g. 1x2, 9x9)
    # SxxExx: standard format — group 1 = season
    # NxN:    bare number format (like in 'WATCH - Show 1x2 - FREE') — group 2 = season, group 3 = episode
    episode_folder_pattern = re.compile(
        r'(?:[Ss](\d{1,2})[Ee]\d{1,2}|(?<!\d)(\d{1,2})[xX](\d{1,2})(?!\d))'
    )
    season_folder_pattern = re.compile(r'^(?:season\s*|s)(\d{1,2})$', re.IGNORECASE)
    
    # Match "ShowTitle S##" abbreviated form ONLY — e.g. "Snowfall S02"
    # Do NOT match spelled-out "Season N" suffix (that's covered by season_folder_pattern above,
    # and show-root folders like "Bloopers Season 8" also end in "Season 8").
    # Require: no dashes/pipes/dots in name, ends with " S<1-2 digits>"
    season_like_pattern = re.compile(
        r'^([^\-\|.]+?)\s+[Ss](\d{1,2})(?:\s*\(?\d{4}\)?)?\s*$'
    )
    
    def _get_episode_season(match):
        """Extract season number from either SxxExx (group 1) or NxN (group 2) match."""
        return int(match.group(1) or match.group(2))
    
    # Categorize folders
    existing_seasons = {}  # season_num -> folder_name
    loose_episodes = {}    # season_num -> [folder_names]
    
    for folder in folders:
        # Check if it's a pure season folder (Season 1, S01, etc.)
        season_match = season_folder_pattern.match(folder.strip())
        if season_match:
            existing_seasons[int(season_match.group(1))] = folder
            continue
        
        # Check if folder name contains SxxExx or NxN (episode folder sitting loose)
        ep_match = episode_folder_pattern.search(folder)
        if ep_match:
            snum = _get_episode_season(ep_match)
            if snum not in loose_episodes:
                loose_episodes[snum] = []
            loose_episodes[snum].append(folder)
            continue
        
        # Check if it's a show-name-style season folder e.g. "Snowfall S02"
        # Only the abbreviated S## form — spelled-out 'Season N' suffixes belong to show-root folders
        s_match = season_like_pattern.match(folder)
        if s_match and not episode_folder_pattern.search(folder):
            existing_seasons[int(s_match.group(2))] = folder
    
    # Step 1b: Identify loose episode FILES sitting directly in the show root
    loose_episode_files = {}  # season_num -> [filenames]
    try:
        for item in os.listdir(show_path):
            item_path = os.path.join(show_path, item)
            if os.path.isfile(item_path) and os.path.splitext(item)[1].lower() in VIDEO_EXTENSIONS:
                s, e = extract_season_episode(item)
                if s is not None:
                    loose_episode_files.setdefault(s, []).append(item)
    except Exception:
        pass

    # Step 2: Plan moves — loose episode folders into season folders
    for snum, ep_folders in sorted(loose_episodes.items()):
        target_season = f"Season {snum}"
        target_path = os.path.join(show_path, target_season)

        for ep_folder in sorted(ep_folders):
            old_path = os.path.join(show_path, ep_folder)
            new_path = os.path.join(target_path, ep_folder)
            changes.append({
                'type': 'move_to_season',
                'old_path': old_path,
                'new_path': new_path,
                'description': f"  📦 {ep_folder}  →  {target_season}/{ep_folder}"
            })

    # Step 2b: Plan moves — loose episode files into season folders
    for snum, filenames in sorted(loose_episode_files.items()):
        target_season = f"Season {snum}"
        target_path = os.path.join(show_path, target_season)
        for filename in sorted(filenames):
            old_path = os.path.join(show_path, filename)
            new_path = os.path.join(target_path, filename)
            changes.append({
                'type': 'move_to_season',
                'old_path': old_path,
                'new_path': new_path,
                'description': f"  📦 {filename}  →  {target_season}/{filename}"
            })
    
    # Step 3: Plan renames — normalize season folder names
    for snum, folder_name in sorted(existing_seasons.items()):
        target_name = f"Season {snum}"
        if folder_name != target_name:
            old_path = os.path.join(show_path, folder_name)
            new_path = os.path.join(show_path, target_name)
            changes.append({
                'type': 'rename_season',
                'old_path': old_path,
                'new_path': new_path,
                'description': f"  ✏️  {folder_name}  →  {target_name}"
            })

    # Step 4: Flatten episode-named subfolders inside existing season folders.
    # e.g. Season 4/Show S04E10 Title/Show.S04E10.mp4  →  Season 4/Show.S04E10.mp4
    MEDIA_EXTENSIONS = set(VIDEO_EXTENSIONS) | {'.nfo', '.jpg', '.jpeg', '.png', '.srt', '.sub', '.ass', '.ssa'}
    for snum, folder_name in sorted(existing_seasons.items()):
        season_path = os.path.join(show_path, folder_name)
        if not os.path.isdir(season_path):
            continue
        try:
            for ep_dir in sorted(os.listdir(season_path)):
                ep_dir_path = os.path.join(season_path, ep_dir)
                if not os.path.isdir(ep_dir_path):
                    continue
                if not episode_folder_pattern.search(ep_dir):
                    continue
                # Move each media file from the episode subfolder up to the season folder
                try:
                    for media_file in sorted(os.listdir(ep_dir_path)):
                        media_path = os.path.join(ep_dir_path, media_file)
                        if os.path.isfile(media_path) and os.path.splitext(media_file)[1].lower() in MEDIA_EXTENSIONS:
                            new_media_path = os.path.join(season_path, media_file)
                            changes.append({
                                'type': 'flatten_to_season',
                                'old_path': media_path,
                                'new_path': new_media_path,
                                'ep_dir': ep_dir_path,
                                'description': f"  📂→📄 {folder_name}/{ep_dir}/{media_file}  →  {folder_name}/{media_file}"
                            })
                except Exception:
                    pass
        except Exception:
            pass

    return changes


def _folder_has_episodes_or_seasons(path):
    """Check if a folder directly contains episode folders (SxxExx / NxN), season folders, or loose episode files."""
    try:
        items = os.listdir(path)
    except Exception:
        return False
    # Matches SxxExx or bare NxN (e.g. 1x2, 9x9) — both are clear episode indicators
    ep_pat = re.compile(r'(?:[Ss]\d{1,2}[Ee]\d{1,2}|(?<!\d)\d{1,2}[xX]\d{1,2}(?!\d))')
    # Pure season folders: "Season 1", "S01"
    season_pat = re.compile(r'^(?:season\s*|s)\d{1,2}(?:\s*\(?\d{4}\)?)?\s*$', re.IGNORECASE)
    # Show-name abbreviated-season: "Snowfall S02" — abbreviated S## form only, no dashes/pipes
    season_like_pat = re.compile(r'^([^\-\|.]+?)\s+[Ss](\d{1,2})(?:\s*\(?\d{4}\)?)?\s*$')
    def _is_season_like(f):
        return bool(season_like_pat.match(f))
    for f in items:
        full = os.path.join(path, f)
        if os.path.isdir(full):
            if ep_pat.search(f) or season_pat.search(f) or _is_season_like(f):
                return True
        elif os.path.isfile(full):
            # Loose video file with episode marker — e.g. Show.S01E01.mkv in show root
            if os.path.splitext(f)[1].lower() in VIDEO_EXTENSIONS and ep_pat.search(f):
                return True
    return False


def run_organizer(folder=None):
    """Menu option: Organize TV show season structure."""
    try:
        print("\n" + "=" * 80)
        print("📚 TV SHOW SEASON ORGANIZER")
        print("=" * 80)
        print("This will:")
        print("  • Group loose episode folders (S05E07, S05E04) into Season folders")
        print("  • Rename season folders (S01, Snowfall S02 → Season 1, Season 2)")
        
        if not folder:
            folder = browse_for_folder("Select a TV show folder or your TV Shows root folder", allow_skip=False)
            if not folder:
                return
        
        # Auto-detect: did the user point at a show folder or a TV root?
        if _folder_has_episodes_or_seasons(folder):
            # User pointed directly at a show folder — organize it
            show_name = Path(folder).name
            print(f"\n📺 Detected show folder: {show_name}")
            show_changes = organize_season_structure(folder)
            if not show_changes:
                print("✅ Already organized! No changes needed.")
                return
            all_changes = [(show_name, show_changes)]
        else:
            # User pointed at a TV root — check each subfolder
            try:
                shows = sorted([f for f in os.listdir(folder) if os.path.isdir(os.path.join(folder, f))])
            except Exception as e:
                print(f"❌ Error reading folder: {e}")
                return
            
            if not shows:
                print("❌ No show folders found.")
                return
            
            # Scan each show for organizable content
            all_changes = []
            print(f"\n� Scanning {len(shows)} show folder(s)...")
            for show in shows:
                show_path = os.path.join(folder, show)
                # Check if episodes are nested one level deeper (e.g. Snowfall/Snowfall/)
                actual_path = show_path
                if not _folder_has_episodes_or_seasons(show_path):
                    # Check one level deeper
                    try:
                        subs = [s for s in os.listdir(show_path) if os.path.isdir(os.path.join(show_path, s))]
                        for sub in subs:
                            sub_path = os.path.join(show_path, sub)
                            if _folder_has_episodes_or_seasons(sub_path):
                                actual_path = sub_path
                                break
                    except Exception:
                        pass
                
                show_changes = organize_season_structure(actual_path)
                if show_changes:
                    all_changes.append((show, show_changes))
            
            if not all_changes:
                print("\n✅ Everything is already organized! No changes needed.")
                return
        
        # Paginated preview of ALL changes
        total = sum(len(c) for _, c in all_changes)
        print(f"\n--- ORGANIZER PREVIEW: {total} change(s) ---")
        preview_lines = []
        for show, show_changes in all_changes:
            preview_lines.append(f"  Show: {show}")
            for c in show_changes:
                preview_lines.append(c['description'])
        
        if not paginated_preview(preview_lines):
            print("User aborted.")
            return
        
        # Execute ALL changes — moves first, then renames.
        # The undo script is written AFTER each successful move so it reflects actual disk state.
        print("\n🔄 Organizing...")
        success: int = 0
        errors: int = 0
        
        flat_changes = [(show, c) for show, changes in all_changes for c in changes]
        # Order: move season groups first, then flatten files inside season folders, then renames
        def _sort_key(x):
            t = x[1]['type']
            if t == 'move_to_season': return 0
            if t == 'flatten_to_season': return 1
            return 2  # rename_season
        flat_changes.sort(key=_sort_key)

        # Collect successful moves in order so the undo script can reverse them correctly.
        # Renames must be undone BEFORE the moves they depend on — we track them separately.
        undo_renames:  list[tuple[str, str]] = []  # (actual_dst, actual_src)
        undo_moves:    list[tuple[str, str]] = []  # (actual_dst, actual_src)
        undo_flattens: list[tuple[str, str, str]] = []  # (actual_dst, actual_src, ep_dir)
        # Track which ep_dirs had successful flattens so we can rmdir them after
        ep_dirs_to_clean: set[str] = set()

        for show, c in flat_changes:
            old_p = Path(c['old_path'])
            new_p = Path(c['new_path'])
            try:
                new_p.parent.mkdir(parents=True, exist_ok=True)
                if (os.name == 'nt' or sys.platform == 'darwin') and str(old_p).lower() == str(new_p).lower() and str(old_p) != str(new_p):
                    temp = str(old_p) + f'._tmp_{uuid.uuid4().hex[:8]}'
                    shutil.move(str(old_p), temp)
                    shutil.move(temp, str(new_p))
                else:
                    shutil.move(str(old_p), str(new_p))
                # Record what actually happened on disk using resolve()-style absolute strings
                actual_src = str(old_p.resolve()) if old_p.exists() else str(old_p)
                actual_dst = str(new_p.resolve()) if new_p.exists() else str(new_p)
                if c['type'] == 'rename_season':
                    undo_renames.append((actual_dst, str(old_p)))
                elif c['type'] == 'flatten_to_season':
                    undo_flattens.append((actual_dst, str(old_p), c['ep_dir']))
                    ep_dirs_to_clean.add(c['ep_dir'])
                else:
                    undo_moves.append((actual_dst, str(old_p)))
                success += 1  # type: ignore
                print(f"  ✅ {c['description'].strip()}")
            except Exception as e:
                errors += 1  # type: ignore
                print(f"  ❌ {c['description'].strip()}: {e}")

        # Send now-empty episode subfolders to trash (best-effort)
        for ep_dir in ep_dirs_to_clean:
            try:
                if os.path.isdir(ep_dir) and not os.listdir(ep_dir):
                    safe_delete(ep_dir)
            except Exception:
                pass
        
        print(f"\n✅ Done! {success} changes applied, {errors} errors.")
        
        # Write undo script using actual resolved paths, in correct reverse order:
        # undo renames first (they were applied last), then undo moves.
        ext = "bat" if os.name == 'nt' else "sh"
        undo_file = f"undo_organize.{ext}"
        try:
            with open(undo_file, 'w', encoding='utf-8') as f:
                if os.name == 'nt':
                    f.write('@echo off\nchcp 65001\necho Undoing organize changes...\n')
                else:
                    f.write('#!/bin/bash\necho "Undoing organize changes..."\n')
                # Undo renames first (reverse order)
                for dst, src in reversed(undo_renames):
                    if os.name == 'nt':
                        f.write(f'move "{dst}" "{src}"\n')
                    else:
                        f.write(f'mv "{dst}" "{src}"\n')
                # Undo flattens (reverse order) — recreate ep_dir then move file back
                seen_ep_dirs: set[str] = set()
                for dst, src, ep_dir in reversed(undo_flattens):
                    if ep_dir not in seen_ep_dirs:
                        seen_ep_dirs.add(ep_dir)
                        if os.name == 'nt':
                            f.write(f'mkdir "{ep_dir}"\n')
                        else:
                            f.write(f'mkdir -p "{ep_dir}"\n')
                    if os.name == 'nt':
                        f.write(f'move "{dst}" "{src}"\n')
                    else:
                        f.write(f'mv "{dst}" "{src}"\n')
                # Then undo moves (reverse order)
                for dst, src in reversed(undo_moves):
                    if os.name == 'nt':
                        f.write(f'move "{dst}" "{src}"\n')
                    else:
                        f.write(f'mv "{dst}" "{src}"\n')
                if os.name == 'nt':
                    f.write('echo Done!\npause\n')
                else:
                    f.write('echo Done!\n')
            print(f"✅ Undo script: {os.path.abspath(undo_file)}")
        except Exception as e:
            print(f"⚠️  Could not write undo script: {e}")
    
    except BackNavigationException:
        print("\n🔙 Going back...")
        return

def scan_media_folder(media_path, media_type="Movies"):
    custom_patterns = [] # No custom patterns in step 1 anymore
    print(f"\n📁 Scanning {media_type} folder: {media_path}")
    if not os.path.exists(media_path):
        print("   ⚠️ Path does not exist.")
        return []
    
    # Auto-organize any loose files into folders first
    organize_loose_files(media_path, media_type)
    
    items = []
    try: folders = [f for f in os.listdir(media_path) if os.path.isdir(os.path.join(media_path, f))]
    except Exception as e:
        print(f"   ❌ Error accessing path: {e}")
        return []
        
    for folder_name in tqdm(folders, desc=f"Processing {media_type}"):
        folder_path = os.path.join(media_path, folder_name)
        if media_type == "Movies":
            year, quality = extract_year(folder_name), extract_quality(folder_name)
            video_files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f)) and os.path.splitext(f)[1].lower() in VIDEO_EXTENSIONS]
            if (not year or not quality) and video_files:
                for vf in video_files:
                    if not year: year = extract_year(vf)
                    if not quality: quality = extract_quality(vf)
                    if year and quality: break
            
            items.append({
                'Folder Name': folder_name, 'Folder Fixed': '',
                'Title': clean_title(folder_name, custom_patterns), 'Title Fixed': '',
                'Year': year or '', 'Year Fixed': '',
                'Quality': quality or '', 'Quality Fixed': '',
                'Size (GB)': get_folder_size(folder_path),
                'Video Files': ', '.join(video_files), 'Files Fixed': ''
            })
        else:
            episodes = scan_tv_show_seasons(folder_path)
            if not episodes:
                items.append({
                    'Show Folder': folder_name, 'Folder Fixed': '',
                    'Title': clean_title(folder_name, custom_patterns), 'Title Fixed': '',
                    'Season': '', 'Season Year': '', 'Episode': '', 'Episode File': '', 'File Fixed': '',
                    'Quality': '', 'Quality Fixed': '', 'Size (GB)': get_folder_size(folder_path)
                })
                continue
            
            show_year = next((ep['year'] for ep in episodes if ep['season_num'] == 1 and ep['year']), None)
            title = clean_title(folder_name, custom_patterns)
            if show_year and f"({show_year})" not in title: title = f"{title} ({show_year})"
            
            for ep in episodes:
                items.append({
                    'Show Folder': folder_name, 'Folder Fixed': '',
                    'Title': title, 'Title Fixed': '',
                    'Season': ep['season_num'], 'Season Year': ep['year'] or '',
                    'Episode': ep['episode_num'], 
                    'Episode File': ep['rel_path'],
                    'File Fixed': '', 'Quality': extract_quality(ep['filename']) or '', 'Quality Fixed': '',
                    'Size (GB)': round(float(ep['size']), 2)  # type: ignore
                })
    return items

def run_scanner(movies_path=None, tv_path=None, output_file=None):
    try:
        if not output_file:
            default_name = "media_library.xlsx"
            out_input = prompt_input(f"\nName for new Excel File (Enter for default '{default_name}'): ")
            if out_input and not out_input.endswith('.xlsx'):
                out_input += '.xlsx'
            output_file = Path(out_input).resolve() if out_input else Path(default_name).resolve()
        else:
            output_file = Path(output_file).resolve()
        
        append_mode = False
        if output_file.exists():
            print(f"\n⚠️ The file '{output_file.name}' already exists.")
            choice = prompt_input("Do you want to (A)ppend to it or (O)verwrite it? (A/O): ", default='A').upper()
            if choice == 'O':
                append_mode = False
            else:
                append_mode = True
        
        if not movies_path and not tv_path:
            movies_path = browse_for_folder("Select Movies folder", allow_skip=True)
            tv_path = browse_for_folder("Select TV Shows folder", allow_skip=True)
    except BackNavigationException:
        print("\n🔙 Going back...")
        return
        
    if not movies_path and not tv_path:
        print("❌ No valid paths provided to scan.")
        return
        
    movies, tv_shows = [], []
    if movies_path: movies = scan_media_folder(movies_path, "Movies")
    if tv_path: tv_shows = scan_media_folder(tv_path, "TV Shows")
    
    if movies or tv_shows:
        print(f"\n📊 {'Updating' if append_mode else 'Creating'} Excel file: {output_file}")
        
        movies_df = pd.DataFrame()
        tv_df = pd.DataFrame()
        meta_dict = {}

        if append_mode:
            try:
                with pd.ExcelFile(output_file) as excel:
                    if 'Movies' in excel.sheet_names:
                        movies_df = pd.read_excel(excel, sheet_name='Movies')
                    if 'TV Shows' in excel.sheet_names:
                        tv_df = pd.read_excel(excel, sheet_name='TV Shows')
                    if 'Metadata' in excel.sheet_names:
                        df_meta = pd.read_excel(excel, sheet_name='Metadata')
                        meta_dict = dict(zip(df_meta['Key'].astype(str), df_meta['Value'].astype(str)))
                    
                if movies:
                    df_new_m = pd.DataFrame(movies)
                    movies_df = pd.concat([movies_df, df_new_m]).drop_duplicates(subset=['Folder Name'], keep='last')
                if tv_shows:
                    df_new_t = pd.DataFrame(tv_shows)
                    tv_df = pd.concat([tv_df, df_new_t]).drop_duplicates(subset=['Show Folder', 'Season', 'Episode'], keep='last')
            except Exception as e:
                print(f"❌ Error reading existing Excel: {e}")
                return
        else:
            if movies: movies_df = pd.DataFrame(movies)
            if tv_shows: tv_df = pd.DataFrame(tv_shows)

        # Update Metadata
        if movies_path: meta_dict['Movies Path'] = str(movies_path)
        if tv_path: meta_dict['TV Shows Path'] = str(tv_path)
        meta_dict['Last Scan'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        meta_df = pd.DataFrame(list(meta_dict.items()), columns=['Key', 'Value'])

        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Write sheets
                if not movies_df.empty: movies_df.to_excel(writer, sheet_name='Movies', index=False)
                if not tv_df.empty: tv_df.to_excel(writer, sheet_name='TV Shows', index=False)
                meta_df.to_excel(writer, sheet_name='Metadata', index=False)
                
                # Auto-adjust column widths for all sheets
                for sheetname in writer.sheets:
                    worksheet = writer.sheets[sheetname]
                    for col in worksheet.columns:
                        max_length = 0
                        column = col[0].column_letter # type: ignore
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except Exception: pass
                        adjusted_width = (max_length + 2)
                        worksheet.column_dimensions[column].width = min(adjusted_width, 60) # Cap at 60
                        
        except PermissionError:
            print(f"\n❌ Cannot write to '{output_file.name}' — it is open in another program.")
            return
        except Exception as e:
            print(f"❌ Error writing Excel: {e}")
            return
            
        print(f"✅ Saved results to {output_file}")


# =========================
# RENAMER LOGIC
# =========================

def get_config_path():
    return os.path.join(os.getcwd(), CONFIG_FILE_NAME)

def clean_str(val, is_year=False):
    if pd.isna(val) or val == '': return ''
    s = str(val)
    if is_year and s.endswith('.0'): s = s[:-2]  # type: ignore
    return s.strip()

def build_movie_folder_name(title, year, quality, size_gb, scheme):
    title, year, quality, size_gb = clean_str(title), clean_str(year, True), clean_str(quality), clean_str(size_gb)
    parts = [title]
    if not scheme.movie_folder_include_year:
        parts = [re.sub(r'\s*\(\d{4}\)\s*', '', title).strip()]
    if scheme.movie_folder_include_quality and quality: parts.append(f"[{quality}]")
    if scheme.movie_folder_include_size and size_gb: parts.append(f"[{size_gb}GB]")
    return " ".join(parts)

def build_movie_file_name(old_filename, title, year, quality, size_gb, scheme):
    title, year, quality, size_gb = clean_str(title), clean_str(year, True), clean_str(quality), clean_str(size_gb)
    name, ext = os.path.splitext(Path(str(old_filename)).name)
    parts = [title]
    if not scheme.movie_file_include_year:
        parts = [re.sub(r'\s*\(\d{4}\)\s*', '', title).strip()]
    if scheme.movie_file_include_quality and quality: parts.append(f"[{quality}]")
    if scheme.movie_file_include_size and size_gb: parts.append(f"[{size_gb}GB]")
    return " ".join(parts) + ext

def build_tv_show_folder_name(title, quality, scheme):
    title, quality = clean_str(title), clean_str(quality)
    if not scheme.tv_parent_include_year:
        title = re.sub(r'\s*\(\d{4}\)\s*', '', title).strip()
    parts = [title]
    if scheme.tv_parent_include_quality and quality: parts.append(f"[{quality}]")
    return " ".join(parts)

def build_season_folder_name(season_num, season_year, scheme):
    season_num, season_year = clean_str(season_num, True), clean_str(season_year, True)
    name = f"Season {season_num}"
    if scheme.tv_season_include_year and season_year: name += f" ({season_year})"
    return name

def build_episode_file_name(episode_file, season_year, quality, size_gb, scheme):
    season_year, quality, size_gb = clean_str(season_year, True), clean_str(quality), clean_str(size_gb)
    name, ext = os.path.splitext(Path(str(episode_file)).name)
    parts = [name]
    if scheme.tv_episode_include_year and season_year: parts.append(f"({season_year})")
    if scheme.tv_episode_include_quality and quality: parts.append(f"[{quality}]")
    if scheme.tv_episode_include_size and size_gb: parts.append(f"[{size_gb}GB]")
    return " ".join(parts) + ext

def get_val(row, col_fixed, col_auto, is_literal=False):
    val_fixed = row.get(col_fixed)
    if pd.notna(val_fixed) and str(val_fixed).strip() != '': return val_fixed
    return col_auto if is_literal else row.get(col_auto, '')

def find_excel_files(folder_path):
    excel_files = []
    try:
        p = Path(folder_path)
        if p.is_dir():
            for file in p.iterdir():
                if file.is_file() and file.suffix in ['.xlsx', '.xls'] and not file.name.startswith('~$'):
                    excel_files.append(file)
    except Exception: pass
    return sorted(excel_files, key=lambda x: x.stat().st_mtime, reverse=True)

def select_excel_file():
    excel_files = find_excel_files(os.getcwd())
    if not excel_files:
        print("\n❌ No Excel files found in the current directory.")
        print("Please run Step 1 (Scan Library) first to generate a spreadsheet.")
        raise BackNavigationException()
    
    print("\n📁 AVAILABLE SPREADSHEETS")
    print("-" * 80)
    for i, file in enumerate(excel_files, 1):
        mod_time = datetime.fromtimestamp(file.stat().st_mtime)
        print(f"[{i}] {file.name} (Modified: {mod_time.strftime('%Y-%m-%d %H:%M')})")
        
    while True:
        choice = prompt_input(f"\nSelect a spreadsheet (1-{len(excel_files)}): ")
        try:
            choice_num = int(choice)
            if 1 <= choice_num <= len(excel_files):
                return str(excel_files[choice_num - 1].resolve())
            print(f"❌ Please enter a number between 1 and {len(excel_files)}")
        except ValueError:
            print("❌ Invalid input. Please enter a number.")

def detect_changes(df_movies, movies_path, df_tv, tv_path, scheme, folder_patterns, file_patterns, llm_results=None, folder_case_sensitive=False, file_case_sensitive=False):
    if llm_results is None: llm_results = {}
    changes = []
    if df_movies is not None and movies_path:
        for _, row in df_movies.iterrows():
            old_name = row['Folder Name']
            if pd.isna(row.get('Title')): continue
            
            # Use LLM result if available, otherwise regex
            if llm_results and old_name in llm_results:  # type: ignore
                llm_r = llm_results[old_name]  # type: ignore
                title = llm_r['title']
                if llm_r['year']: title = f"{title} ({llm_r['year']})"
            else:
                base_folder = clean_title(old_name, folder_patterns, folder_case_sensitive)
                title = get_val(row, 'Title Fixed', base_folder, is_literal=True)
            year = get_val(row, 'Year Fixed', 'Year')
            quality = get_val(row, 'Quality Fixed', 'Quality')
            size_gb = row.get('Size (GB)', '')
            
            new_name = build_movie_folder_name(title, year, quality, size_gb, scheme)
            if new_name and new_name != old_name:
                old_path = Path(movies_path) / str(old_name)
                new_path = Path(movies_path) / new_name
                changes.append({'type': 'movie', 'old_name': old_name, 'new_name': new_name, 'old_path': str(old_path), 'new_path': str(new_path), 'exists': old_path.exists()})
            
            # Record changes for movie files inside the folder
            video_files = row.get('Video Files', '')
            if pd.notna(video_files) and video_files:
                for vf in str(video_files).split(','):
                    vf = vf.strip()
                    if not vf: continue
                    
                    orig_name, orig_ext = os.path.splitext(vf)
                    # Use LLM result if available, otherwise regex
                    if llm_results and vf in llm_results:  # type: ignore
                        llm_r = llm_results[vf]  # type: ignore
                        vf_clean = llm_r['title']
                        if llm_r['year']: vf_clean = f"{vf_clean} ({llm_r['year']})"
                    else:
                        vf_clean = clean_title(vf, file_patterns, file_case_sensitive)
                    new_file_name = build_movie_file_name(vf, vf_clean, year, quality, size_gb, scheme)
                    # Force the original extension back on in case clean_title stripped it
                    _, new_ext = os.path.splitext(new_file_name)
                    if not new_ext and orig_ext:
                        new_file_name = new_file_name + orig_ext
                    
                    if new_file_name != vf:
                        # old folder path might change, but the rename order handles files first
                        old_file_path = Path(movies_path) / str(old_name) / vf
                        new_file_path = old_file_path.parent / new_file_name
                        changes.append({'type': 'movie_file', 'old_name': vf, 'new_name': new_file_name, 'old_path': str(old_file_path), 'new_path': str(new_file_path), 'exists': old_file_path.exists()})
    
    if df_tv is not None and tv_path:
        # Pattern to detect when the scanner was run on a single show root:
        # in that case each "show folder" in the Excel is actually a Season X folder.
        _season_folder_re = re.compile(r'^Season\s+\d+$', re.IGNORECASE)

        for show_folder, show_episodes in df_tv.groupby('Show Folder'):
            first_ep = show_episodes.iloc[0]

            # If the show_folder matches "Season N", the tv_path IS the show root
            # and each show_folder entry is really a season — don't nest paths further.
            show_is_season = bool(_season_folder_re.match(str(show_folder)))

            # Use LLM result if available for show folder, otherwise regex
            if llm_results and show_folder in llm_results:  # type: ignore
                llm_r = llm_results[show_folder]  # type: ignore
                title = llm_r['title']
                if llm_r['year']: title = f"{title} ({llm_r['year']})"
            else:
                title = get_val(first_ep, 'Title Fixed', show_folder, is_literal=True)
            if not title or pd.isna(title): continue

            new_show_folder = build_tv_show_folder_name(title, None, scheme)
            if new_show_folder != show_folder and not show_is_season:
                changes.append({'type': 'tv_show', 'old_name': show_folder, 'new_name': new_show_folder, 'old_path': str(Path(tv_path) / str(show_folder)), 'new_path': str(Path(tv_path) / new_show_folder), 'exists': (Path(tv_path) / str(show_folder)).exists()})

            for season_num, season_episodes in show_episodes.groupby('Season'):
                if pd.isna(season_num): continue
                season_num = int(season_num)
                season_year = season_episodes.iloc[0].get('Season Year', '')
                old_season_folder = f"Season {season_num}"
                new_season_folder = build_season_folder_name(season_num, season_year, scheme)
                # Only propose season-folder rename when show_folder is a real show (not itself a season)
                if new_season_folder != old_season_folder and not show_is_season:
                    changes.append({'type': 'tv_season', 'old_name': old_season_folder, 'new_name': new_season_folder, 'old_path': str(Path(tv_path) / str(show_folder) / old_season_folder), 'new_path': str(Path(tv_path) / str(show_folder) / new_season_folder), 'exists': (Path(tv_path) / str(show_folder) / old_season_folder).exists()})

                for _, episode in season_episodes.iterrows():
                    episode_file = episode.get('Episode File', '')
                    quality = get_val(episode, 'Quality Fixed', 'Quality')
                    size_gb = episode.get('Size (GB)', '')
                    if not episode_file or pd.isna(episode_file): continue

                    # Apply file patterns to episode filename manually
                    episode_base = Path(str(episode_file)).name
                    orig_name, orig_ext = os.path.splitext(episode_base)
                    # Extract season/episode to protect it during stripping
                    s, e = extract_season_episode(episode_base)
                    se_code = f"S{s:02d}E{e:02d}" if s and e else ""

                    # Use LLM result if available, otherwise regex
                    if llm_results and episode_base in llm_results:  # type: ignore
                        llm_r = llm_results[episode_base]  # type: ignore
                        ep_clean = llm_r['title']
                        if se_code and se_code not in ep_clean:
                            ep_clean = f"{se_code} - {ep_clean}"
                    else:
                        ep_clean = clean_title(episode_base, file_patterns, file_case_sensitive)
                        # Check for any S/E code variant (padded or unpadded) before prepending
                        already_has_se = bool(re.search(r'[Ss]\d{1,2}[Ee]\d{1,2}', ep_clean))
                        if se_code and not already_has_se:
                            ep_clean = f"{se_code} - {ep_clean}"

                    # Build name and ensure original extension is preserved
                    new_episode_name = build_episode_file_name(ep_clean, season_year, quality, size_gb, scheme)
                    # Force the original extension back on in case clean_title stripped it
                    _, new_ext = os.path.splitext(new_episode_name)
                    if not new_ext and orig_ext:
                        new_episode_name = new_episode_name + orig_ext
                    old_episode_name = episode_base
                    if new_episode_name != old_episode_name:
                        old_path = Path(tv_path) / str(show_folder) / str(episode_file)
                        if show_is_season:
                            # tv_path is the show root; show_folder IS the season folder —
                            # place the renamed episode directly inside it (flatten episode subfolders)
                            new_path = Path(tv_path) / str(show_folder) / new_episode_name
                        else:
                            new_path = Path(tv_path) / str(show_folder) / old_season_folder / new_episode_name
                        changes.append({'type': 'tv_episode', 'old_name': old_episode_name, 'new_name': new_episode_name, 'old_path': str(old_path), 'new_path': str(new_path), 'exists': old_path.exists()})
    return changes

def run_renamer(movies_path=None, tv_path=None, excel_path=None, rename_mode='both'):
    # rename_mode: 'files' = episode/movie files only
    #              'folders' = show/season/movie folders only
    #              'both' = everything (default)
    try:
        if not excel_path:
            excel_path = select_excel_file()

        try:
            with pd.ExcelFile(excel_path) as excel:
                has_movies, has_tv = 'Movies' in excel.sheet_names, 'TV Shows' in excel.sheet_names
                df_movies = pd.read_excel(excel, sheet_name='Movies') if has_movies else None
                df_tv = pd.read_excel(excel, sheet_name='TV Shows') if has_tv else None

                # Read paths from metadata (only if not already provided)
                if not movies_path or not tv_path:
                    if 'Metadata' in excel.sheet_names:
                        df_meta = pd.read_excel(excel, sheet_name='Metadata')
                        meta_dict = dict(zip(df_meta['Key'], df_meta['Value']))
                        if not movies_path: movies_path = meta_dict.get('Movies Path')
                        if not tv_path: tv_path = meta_dict.get('TV Shows Path')
                
        except Exception as e:
            print(f"❌ Error reading Excel: {e}")
            return
            
        print(f"\n✅ Loaded Library Locations:")
        if not movies_path and has_movies: 
            print("⚠️ Movies path missing from spreadsheet.")
            movies_path = browse_for_folder("Select Movies folder", allow_skip=True)
        if not tv_path and has_tv: 
            print("⚠️ TV Shows path missing from spreadsheet.")
            tv_path = browse_for_folder("Select TV Shows folder", allow_skip=True)
            
        if not movies_path and not tv_path:
            print("❌ No valid paths provided for renaming.")
            return

        if movies_path: print(f"   Movies: {movies_path}")
        if tv_path: print(f"   TV Shows: {tv_path}")

        # ── LLM vs Regex decision ──────────────────────────────────────
        use_llm = False
        llm_results = None
        folder_patterns = []
        file_patterns = []
        folder_case_sensitive = False
        file_case_sensitive = False
        
        if LLM_AVAILABLE and ask_yes_no("\n🤖 Use AI-powered title cleaning?", default=False):
            print("\n📡 SELECT LLM PROVIDER")
            print("-" * 40)
            print("1. Gemini (Google — free tier available)")
            print("2. OpenAI (requires paid API key)")
            print("3. Ollama (local, free, no API key)")
            
            provider_choice = prompt_input("\nSelect provider (1-3): ")
            provider_map = {'1': 'gemini', '2': 'openai', '3': 'ollama'}
            provider = provider_map.get(provider_choice)
            
            if provider:
                config = load_llm_config()
                api_key = None
                model = None
                
                if provider in ['gemini', 'openai']:
                    saved_key = config.get(f"{provider}_api_key", '')
                    if saved_key:
                        masked = saved_key[:8] + '...' + saved_key[-4:]
                        print(f"\n🔑 Found saved API key: {masked}")
                        if not ask_yes_no("Use this key?", default=True):
                            saved_key = ''
                    if not saved_key:
                        saved_key = prompt_input(f"Enter your {provider.title()} API key: ")
                        config[f"{provider}_api_key"] = saved_key
                        save_llm_config(config)
                        print("💾 API key saved for future runs.")
                    api_key = saved_key
                    
                elif provider == 'ollama':
                    saved_url = config.get('ollama_url', 'http://localhost:11434')
                    ollama_url = prompt_input(f"Ollama host URL (Enter for '{saved_url}'): ") or saved_url
                    config['ollama_url'] = ollama_url
                    save_llm_config(config)
                    
                    print(f"\n🔍 Checking for models at {ollama_url}...")
                    models = list_ollama_models(ollama_url)
                    if not models:
                        print("❌ Could not connect to Ollama or no models installed.")
                        print("   Make sure Ollama is running (ollama serve) and you have a model pulled.")
                        print("   Example: ollama pull llama3")
                        print("\n⚠️ Falling back to regex patterns.")
                    else:
                        print(f"\n📚 INSTALLED OLLAMA MODELS")
                        print("-" * 40)
                        for i, m in enumerate(models, 1):
                            print(f"  [{i}] {m}")
                        model_choice = prompt_input(f"\nSelect a model (1-{len(models)}): ")
                        try:
                            model = models[int(model_choice) - 1]
                            config['ollama_model'] = model
                            save_llm_config(config)
                            print(f"   ✅ Using model: {model}")
                        except (ValueError, IndexError):
                            print("❌ Invalid selection. Falling back to regex patterns.")
                            provider = None
                
                # Collect all names to send to LLM
                if provider:
                    print("\n🤖 Sending filenames to AI for cleaning...")
                    all_names = []
                    if df_movies is not None:
                        all_names.extend(df_movies['Folder Name'].dropna().tolist())
                        for vf_list in df_movies['Video Files'].dropna():
                            all_names.extend([v.strip() for v in str(vf_list).split(',') if v.strip()])
                    if df_tv is not None:
                        all_names.extend(df_tv['Show Folder'].dropna().unique().tolist())
                        for ef in df_tv['Episode File'].dropna():
                            all_names.append(Path(str(ef)).name)
                    
                    # Deduplicate while preserving order
                    seen = set()
                    unique_names = []
                    for name in all_names:
                        if name not in seen:
                            seen.add(name)
                            unique_names.append(name)
                    
                    print(f"   📋 Processing {len(unique_names)} unique names...")
                    ollama_host = ollama_url if provider == 'ollama' else None
                    with tqdm(total=len(unique_names), desc="🤖 AI Cleaning") as pbar:
                        llm_results = clean_titles_with_llm(unique_names, provider, api_key, model, ollama_url=ollama_host, pbar=pbar)
                    
                    if llm_results:
                        print(f"   ✅ AI cleaned {len(llm_results)}/{len(unique_names)} names successfully!")
                        use_llm = True
                    else:
                        print("   ⚠️ AI cleaning failed. Falling back to regex patterns.")
        
        if not use_llm:
            # Existing regex flow
            folder_patterns, folder_case_sensitive = get_custom_patterns("Folder")
            if ask_yes_no("\nWould you also like to strip the same words from your Files?", default=True):
                file_patterns = list(folder_patterns)
                file_case_sensitive = folder_case_sensitive
            else:
                file_patterns, file_case_sensitive = get_custom_patterns("File")

        config_path = get_config_path()
        scheme = NamingScheme()
        run_questionnaire = True
        
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r') as f:
                    scheme = NamingScheme.from_dict(json.load(f))
                
                print("\n💾 Found saved naming preferences:")
                if has_movies:
                    print("   🎬 Movies:")
                    print("      [FOLDER]")
                    print(f"      - Include Year:    {'Yes' if scheme.movie_folder_include_year else 'No'}")
                    print(f"      - Include Quality: {'Yes' if scheme.movie_folder_include_quality else 'No'}")
                    print(f"      - Include Size:    {'Yes' if scheme.movie_folder_include_size else 'No'}")
                    print("      [FILE]")
                    print(f"      - Include Year:    {'Yes' if scheme.movie_file_include_year else 'No'}")
                    print(f"      - Include Quality: {'Yes' if scheme.movie_file_include_quality else 'No'}")
                    print(f"      - Include Size:    {'Yes' if scheme.movie_file_include_size else 'No'}")
                if has_tv:
                    print("   📺 TV Shows:")
                    print("      [PARENT FOLDER]")
                    print(f"      - Include Year:    {'Yes' if scheme.tv_parent_include_year else 'No'}")
                    print(f"      - Include Quality: {'Yes' if scheme.tv_parent_include_quality else 'No'}")
                    print("      [SEASON FOLDER]")
                    print(f"      - Include Year:    {'Yes' if scheme.tv_season_include_year else 'No'}")
                    print("      [EPISODE FILE]")
                    print(f"      - Include Year:    {'Yes' if scheme.tv_episode_include_year else 'No'}")
                    print(f"      - Include Quality:  {'Yes' if scheme.tv_episode_include_quality else 'No'}")
                    print(f"      - Include Size:     {'Yes' if scheme.tv_episode_include_size else 'No'}")

                if ask_yes_no("\nUse these saved preferences?", default=True):
                    run_questionnaire = False
            except Exception: pass
            
        if run_questionnaire:
            print("\n🎨 NAMING SCHEME BUILDER")
            if has_movies:
                print("\n🎬 Movies:\n  [FOLDER]")
                scheme.movie_folder_include_year = ask_yes_no("  Include Year on folder?", True)
                scheme.movie_folder_include_quality = ask_yes_no("  Include Quality on folder?", True)
                scheme.movie_folder_include_size = ask_yes_no("  Include Size on folder?", False)
                print("  [FILE]")
                scheme.movie_file_include_year = ask_yes_no("  Include Year on file?", True)
                scheme.movie_file_include_quality = ask_yes_no("  Include Quality on file?", True)
                scheme.movie_file_include_size = ask_yes_no("  Include Size on file?", False)
            if has_tv:
                print("\n📺 TV Shows:\n  [PARENT FOLDER]")
                scheme.tv_parent_include_year = ask_yes_no("  Include Year on folder?", True)
                scheme.tv_parent_include_quality = ask_yes_no("  Include Quality on folder?", False)
                print("  [SEASON FOLDER]")
                scheme.tv_season_include_year = ask_yes_no("  Include Year on folder?", True)
                print("  [EPISODE FILE]")
                scheme.tv_episode_include_year = ask_yes_no("  Include Year on file?", False)
                scheme.tv_episode_include_quality = ask_yes_no("  Include Quality on file?", True)
                scheme.tv_episode_include_size = ask_yes_no("  Include Size on file?", False)
            
            try:
                with open(config_path, 'w') as f:
                    json.dump(scheme.to_dict(), f, indent=2)
                print("💾 Preferences saved.")
            except Exception: pass

        print("\n🔍 Analyzing changes...")
        changes = detect_changes(df_movies, movies_path, df_tv, tv_path, scheme, folder_patterns, file_patterns, llm_results,
                                 folder_case_sensitive=folder_case_sensitive if not use_llm else False,
                                 file_case_sensitive=file_case_sensitive if not use_llm else False)
        
    except BackNavigationException:
        print("\n🔙 Going back...")
        return
        
    if not changes:
        print("✅ No changes needed!")
        return
        
    print(f"\n--- RENAME PREVIEW: {len(changes)} change(s) ---")
    preview_lines = []
    for c in changes:
        status = "[OK]" if c['exists'] else "[NOT FOUND]"
        preview_lines.append(f"{status}  {c['old_name']}  ->  {c['new_name']}")
    
    if not paginated_preview(preview_lines):
        print("User aborted.")
        return
        
    # Generate Undo scripts — separate for names-only vs full undo
    ext = "bat" if os.name == 'nt' else "sh"
    
    # Categorize changes: name-only (same parent dir) vs location changes
    name_only_changes = []
    location_changes = []
    for c in changes:
        old_parent = str(Path(str(c['old_path'])).parent)
        new_parent = str(Path(str(c['new_path'])).parent)
        if old_parent.lower() == new_parent.lower():
            name_only_changes.append(c)
        else:
            location_changes.append(c)
    
    def write_undo_script(filename, items, label):
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                if os.name == 'nt':
                    f.write("@echo off\nchcp 65001\n")
                    f.write(f"echo Undoing {label}...\n")
                else:
                    f.write("#!/bin/bash\n")
                    f.write(f'echo "Undoing {label}..."\n')
                for c in reversed(items):
                    if os.name == 'nt': f.write(f'move "{c["new_path"]}" "{c["old_path"]}"\n')
                    else: f.write(f'mv "{c["new_path"]}" "{c["old_path"]}"\n')
                if os.name == 'nt':
                    f.write("echo Done!\npause\n")
            if os.name != 'nt': os.chmod(filename, 0o755)
            return True
        except Exception as e:
            print(f"⚠️ Failed to create {filename}: {e}")
            return False
    
    # Always create a full undo script
    write_undo_script(f"undo_all.{ext}", changes, "ALL changes (names + locations)")
    
    # Create names-only undo if there are name changes
    if name_only_changes:
        write_undo_script(f"undo_names.{ext}", name_only_changes, "name changes only")
    
    print(f"✅ Created undo scripts:")
    print(f"   📄 undo_all.{ext} — reverses everything")
    if name_only_changes:
        print(f"   📄 undo_names.{ext} — reverses names only (keeps locations)")

    print("\n🔄 Renaming...")
    success: int = 0
    errors: int = 0
    completed_changes = []
    
    # Filter changes based on rename_mode
    FILE_TYPES = {'tv_episode', 'movie_file'}
    FOLDER_TYPES = {'tv_show', 'tv_season', 'movie'}
    if rename_mode == 'files':
        changes = [c for c in changes if c['type'] in FILE_TYPES]
    elif rename_mode == 'folders':
        changes = [c for c in changes if c['type'] in FOLDER_TYPES]

    # Sort: files first (episodes, movie files), then season folders, then parent folders
    # This ensures we never rename a parent before its children
    type_priority = {'tv_episode': 0, 'movie_file': 0, 'tv_season': 1, 'tv_show': 2, 'movie': 3}
    changes_sorted = sorted(changes, key=lambda x: (type_priority.get(x['type'], 2), -len(x['old_path'])))
    
    for i, c in enumerate(tqdm(changes_sorted)):
        old_p = Path(c['old_path'])
        new_p = Path(c['new_path'])
        
        if not old_p.exists():
            errors += 1
            print(f"\n❌ Not found: {old_p}")
            continue
            
        try:
            new_p.parent.mkdir(parents=True, exist_ok=True)
            
            # Handle case-only renames on Windows (NTFS is case-insensitive)
            if (os.name == 'nt' or sys.platform == 'darwin') and str(old_p).lower() == str(new_p).lower() and str(old_p) != str(new_p):
                temp_path = str(old_p) + f'._tmp_{uuid.uuid4().hex[:8]}'
                shutil.move(str(old_p), temp_path)
                shutil.move(temp_path, str(new_p))
            else:
                shutil.move(str(old_p), str(new_p))
            success += 1
            completed_changes.append(c)
            
            # When a folder is renamed, update all subsequent items that reference it
            old_str = str(old_p)
            new_str = str(new_p)
            for j in range(i + 1, len(changes_sorted)):
                future = changes_sorted[j]  # type: ignore
                old_path_str = str(future['old_path'])
                new_path_str = str(future['new_path'])
                if old_path_str.startswith(old_str + os.sep) or old_path_str.startswith(old_str + '/'):
                    future['old_path'] = new_str + old_path_str[len(old_str):]  # type: ignore
                if new_path_str.startswith(old_str + os.sep) or new_path_str.startswith(old_str + '/'):
                    future['new_path'] = new_str + new_path_str[len(old_str):]  # type: ignore
                    
        except Exception as e:
            errors += 1
            print(f"\n❌ Failed to rename {c['old_name']}: {e}")
            
    print(f"\n✅ Renaming complete! {success} successes, {errors} errors.")
    
    # Log changes to Excel
    if completed_changes:
        try:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            change_records = []
            for c in completed_changes:
                change_records.append({
                    'Type': c['type'],
                    'Old Name': c['old_name'],
                    'New Name': c['new_name'],
                    'Old Path': c['old_path'],
                    'New Path': c['new_path'],
                    'Timestamp': timestamp
                })
            df_changes = pd.DataFrame(change_records)
            
            # Append to existing Changes sheet or create new one
            try:
                existing = pd.read_excel(excel_path, sheet_name='Changes')
                df_changes = pd.concat([existing, df_changes], ignore_index=True)
            except Exception: pass
            
            with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:  # type: ignore
                df_changes.to_excel(writer, sheet_name='Changes', index=False)
            print(f"📊 Logged {len(completed_changes)} changes to '{Path(str(excel_path)).name}' → Changes sheet")
        except Exception as e:
            print(f"⚠️ Could not log changes to Excel: {e}")
    
    # Post-rename verification
    print("\n" + "=" * 80)
    print("🔍 VERIFICATION")
    print("=" * 80)
    print("Please check your file system to make sure everything looks correct.")
    print(f"   📄 Undo scripts are in: {os.getcwd()}")
    print(f"   📊 Change log saved to: {Path(str(excel_path)).name} → Changes sheet")
    
    if ask_yes_no("\n✅ Does everything look correct?", default=True):
        print("🎉 Great! Your media library has been cleaned up.")
    else:
        print("\n🔙 UNDO OPTIONS")
        print("-" * 40)
        print("  [1] Undo names only (keep file locations)")
        print("  [2] Undo everything (names + locations)")
        print("  [3] Don't undo anything right now")
        
        undo_choice = input("\nSelect (1-3): ").strip()
        
        if undo_choice == '1' and name_only_changes:
            print(f"\n🔄 Running undo_names.{ext}...")
            try:
                if os.name == 'nt':
                    os.system(f'"{os.path.abspath(f"undo_names.{ext}")}"')
                else:
                    os.system(f'./{f"undo_names.{ext}"}')
                print("✅ Name changes have been reversed.")
            except Exception as e:
                print(f"❌ Error running undo script: {e}")
                print(f"   You can run it manually: {os.path.abspath(f'undo_names.{ext}')}")
        elif undo_choice == '2':
            print(f"\n🔄 Running undo_all.{ext}...")
            try:
                if os.name == 'nt':
                    os.system(f'"{os.path.abspath(f"undo_all.{ext}")}"')
                else:
                    os.system(f'./{f"undo_all.{ext}"}')
                print("✅ All changes have been reversed.")
            except Exception as e:
                print(f"❌ Error running undo script: {e}")
                print(f"   You can run it manually: {os.path.abspath(f'undo_all.{ext}')}")
        else:
            print(f"👍 No undo performed. You can run the scripts manually anytime:")
            if name_only_changes:
                print(f"   📄 Names only: {os.path.abspath(f'undo_names.{ext}')}")
            print(f"   📄 Everything:  {os.path.abspath(f'undo_all.{ext}')}")


def run_text_export():
    """Export media library to a formatted text file."""
    try:
        print("\n" + "=" * 80)
        print("📝 EXPORT TO TEXT FILE")
        print("=" * 80)
        
        movies_path = browse_for_folder("Select Movies folder", allow_skip=True)
        tv_path = browse_for_folder("Select TV Shows folder", allow_skip=True)
        
        if not movies_path and not tv_path:
            print("❌ No folders selected.")
            return
        
        timestamp = datetime.now().strftime('%Y-%m-%d_%H%M')
        output_file = f"media_library_{timestamp}.txt"
        
        lines = []
        lines.append("=" * 80)
        lines.append("MEDIA LIBRARY")
        lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append("=" * 80)
        
        if movies_path and os.path.exists(movies_path):
            lines.append("\n" + "─" * 40)
            lines.append(f"🎬 MOVIES ({movies_path})")
            lines.append("─" * 40)
            movie_folders = sorted([f for f in os.listdir(movies_path) if os.path.isdir(os.path.join(movies_path, f))])
            for i, folder in enumerate(movie_folders, 1):
                size = get_folder_size(os.path.join(movies_path, folder))
                lines.append(f"  {i:4d}. {folder}  [{size} GB]")
            lines.append(f"\n  Total: {len(movie_folders)} movies")
        
        if tv_path and os.path.exists(tv_path):
            lines.append("\n" + "─" * 40)
            lines.append(f"📺 TV SHOWS ({tv_path})")
            lines.append("─" * 40)
            show_folders = sorted([f for f in os.listdir(tv_path) if os.path.isdir(os.path.join(tv_path, f))])
            for i, show in enumerate(show_folders, 1):
                show_path = os.path.join(tv_path, show)
                size = get_folder_size(show_path)
                lines.append(f"\n  {i:4d}. {show}  [{size} GB]")
                # List seasons
                seasons = sorted([s for s in os.listdir(show_path) if os.path.isdir(os.path.join(show_path, s))])
                for season in seasons:
                    season_path = os.path.join(show_path, season)
                    ep_count = sum(1 for f in os.listdir(season_path) 
                                   if os.path.isfile(os.path.join(season_path, f)) 
                                   and os.path.splitext(f)[1].lower() in VIDEO_EXTENSIONS) if os.path.isdir(season_path) else 0
                    lines.append(f"        └─ {season}  ({ep_count} episodes)")
            lines.append(f"\n  Total: {len(show_folders)} shows")
        
        lines.append("\n" + "=" * 80)
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("\n".join(lines))
        
        print(f"\n✅ Exported to: {os.path.abspath(output_file)}")
        print(f"   {sum(1 for l in lines if l.strip())} lines written.")
    
    except BackNavigationException:
        print("\n🔙 Going back...")
        return



def run_extension_converter():
    """Menu option 4: Batch-rename file extensions within a folder."""
    print("\n" + "=" * 60)
    print("  FILE EXTENSION CONVERTER")
    print("=" * 60)
    print("Rename file extensions in bulk — e.g. .ts -> .mp4")
    print()

    try:
        folder = browse_for_folder("Select the folder to convert files in", allow_skip=False)
        if not folder:
            return

        # List unique extensions found
        found_exts = {}  # type: ignore
        for root, _, files in os.walk(folder):
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                if ext:
                    if ext not in found_exts:
                        found_exts[ext] = []
                    found_exts[ext].append(str(os.path.join(root, f)))  # type: ignore

        if not found_exts:
            print("No files with extensions found in that folder.")
            return

        print("\nExtensions found:")
        for ext, files in sorted(found_exts.items()):
            print(f"  {ext:12}  ({len(files)} file{'s' if len(files) != 1 else ''})")

        from_ext = prompt_input("\nConvert FROM extension (e.g. ts): ").strip().lower()
        if not from_ext:
            return
        if not from_ext.startswith('.'):
            from_ext = '.' + from_ext

        if from_ext not in found_exts:
            print(f"No files with extension '{from_ext}' found.")
            return

        to_ext = prompt_input(f"Convert TO extension (e.g. mp4): ").strip().lower()
        if not to_ext:
            return
        if not to_ext.startswith('.'):
            to_ext = '.' + to_ext

        if from_ext == to_ext:
            print("Source and target extensions are the same. Nothing to do.")
            return

        targets = found_exts[from_ext]
        preview_lines = []
        changes = []
        for old_path in sorted(targets):
            name_no_ext = os.path.splitext(old_path)[0]
            new_path = name_no_ext + to_ext
            preview_lines.append(f"  {os.path.relpath(old_path, folder)}  ->  {os.path.relpath(new_path, folder)}")
            changes.append((old_path, new_path))

        print(f"\n--- EXTENSION CONVERT PREVIEW: {len(changes)} file(s) ---")
        if not paginated_preview(preview_lines):
            print("User aborted.")
            return

        success = 0
        errors = 0
        for old_path, new_path in changes:
            try:
                if os.path.exists(new_path):
                    print(f"  [SKIP] Already exists: {os.path.basename(new_path)}")
                    continue
                os.rename(str(old_path), str(new_path))
                success += 1
            except Exception as e:
                print(f"  [ERROR] {os.path.basename(old_path)}: {e}")
                errors += 1

        print(f"\nDone! {success} converted, {errors} errors.")

    except BackNavigationException:
        print("\nGoing back...")
        return


# =========================
# MAIN ENTRY
# =========================

def run_wizard(args=None):
    """Main interactive menu / wizard."""
    # Handle non-interactive CLI args
    if args and args.action:
        _handle_cli_args(args)
        return

    global USE_EMOJIS
    config_path = os.path.join(os.getcwd(), CONFIG_FILE_NAME)
    
    # Pre-Step 1: Check and ask for emoji preference if not saved
    config_data = {}
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
                if isinstance(loaded, dict):
                    config_data = loaded
        except Exception:
            pass
            
    if 'use_emojis' in config_data:
        USE_EMOJIS = config_data['use_emojis']
    else:
        print("\n" + "="*80)
        print("PRE-STEP 1: UI PREFERENCES")
        print("="*80)
        ans = _builtin_input("Would you like to use Emojis in the interface? (y/n) [default: y]: ").strip().lower()
        if ans == 'n':
            USE_EMOJIS = False
        else:
            USE_EMOJIS = True
            
        config_data['use_emojis'] = USE_EMOJIS
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=4)
        except Exception:
            pass

    while True:
        try:
            print("\n" + "="*80)
            print("🎬 UNIFIED MEDIA ORGANIZER")
            print("="*80)
            print()
            print("STEP 1: What would you like to do?")
            print()
            print('  [0] ⚙️ Toggle UI Emojis (Currently: ' + ('ON' if USE_EMOJIS else 'OFF') + ')')
            print()
            print('  [1] 🧹 Clean file names')
            print('      Strips codec/quality tags from episode and movie file names.')
            print('      Example: "Show.S01E01.1080p.WEBRip.x264-GROUP.mkv" -> "Show S01E01.mkv"')
            print()
            print('  [2] 🗂️  Clean folder names')
            print('      Renames show/movie folders. TV show folders only rename if you set')
            print('      a "Title Fixed" value in the Excel — safe for Plex libraries.')
            print('      Example: Movie folder "The.Matrix.1999.BluRay.x264" -> "The Matrix (1999)"')
            print()
            print('  [3] 🧹🗂️  Clean both file and folder names')
            print()
            print('  [4] Organize file structure (TV Shows only)')
            print('      Example: Loose "S05E07" folders -> grouped into "Season 5/"')
            print()
            print('  [5] Do everything (organize structure, then clean folders and files)')
            print()
            print('  [6] Convert file extensions  (e.g. .ts -> .mp4)')
            print()
            print('  [7] Scan Library only (create/update Excel spreadsheet)')
            print()
            print('  [8] Export library to text file')
            print()
            print('  [9] Exit')
            print()
            print("Tip: Type 'back' or 'b' at any prompt to return to the previous step.")

            choice = prompt_input("\nSelect an option (0-9): ")

            if choice == '0':
                USE_EMOJIS = not USE_EMOJIS
                config_data['use_emojis'] = USE_EMOJIS
                try:
                    with open(config_path, 'w', encoding='utf-8') as f:
                        json.dump(config_data, f, indent=4)
                except Exception:
                    pass
                print(f"\nEmojis are now {'ON' if USE_EMOJIS else 'OFF'}.")
                continue

            if choice == '9':
                break

            if choice == '8':
                run_text_export()
                continue

            if choice == '6':
                run_extension_converter()
                continue

            # For 1, 2, 3, 4, 5, 7: Ask for folders ONCE
            print("\n" + "-"*40)
            print("STEP 2: Where are your files?")
            print("Press Enter to skip a folder if you don't have it.")
            print("-"*40)

            movies_path, tv_path = None, None

            # If they just want to organize TV structures, skip movies prompt
            if choice == '4':
                tv_path = browse_for_folder("Select your TV Shows folder", allow_skip=False)
                if not tv_path: continue
            else:
                movies_path = browse_for_folder("Select Movies folder", allow_skip=True)
                tv_path = browse_for_folder("Select TV Shows folder", allow_skip=True)
                if not movies_path and not tv_path:
                    print("No folders selected.")
                    continue

            # For 1, 2, 3, 5, 7: Ask for Excel file name ONCE
            excel_path = None
            if choice in ['1', '2', '3', '5', '7']:
                print("\n" + "-"*40)
                print("STEP 3: Spreadsheet Name")
                print("-" * 40)
                default_name = "media_library.xlsx"
                out_input = prompt_input(f"Name for Excel File (Enter for default '{default_name}'): ")
                if out_input and not out_input.endswith('.xlsx'):
                    out_input += '.xlsx'
                excel_path = Path(out_input).resolve() if out_input else Path(default_name).resolve()

            # Execute selected flow
            if choice == '1':
                print("\n" + "="*60)
                print("RUNNING PIPELINE: Scan -> Clean File Names")
                print("="*60)
                print("\n[1/2] Scanning library...")
                run_scanner(movies_path=movies_path, tv_path=tv_path, output_file=excel_path)
                print("\n[2/2] Cleaning file names...")
                run_renamer(movies_path=movies_path, tv_path=tv_path, excel_path=excel_path, rename_mode='files')

            elif choice == '2':
                print("\n" + "="*60)
                print("RUNNING PIPELINE: Scan -> Clean Folder Names")
                print("="*60)
                print("\n[1/2] Scanning library...")
                run_scanner(movies_path=movies_path, tv_path=tv_path, output_file=excel_path)
                print("\n[2/2] Cleaning folder names...")
                run_renamer(movies_path=movies_path, tv_path=tv_path, excel_path=excel_path, rename_mode='folders')

            elif choice == '3':
                print("\n" + "="*60)
                print("RUNNING PIPELINE: Scan -> Clean File and Folder Names")
                print("="*60)
                print("\n[1/2] Scanning library...")
                run_scanner(movies_path=movies_path, tv_path=tv_path, output_file=excel_path)
                print("\n[2/2] Cleaning file and folder names...")
                run_renamer(movies_path=movies_path, tv_path=tv_path, excel_path=excel_path, rename_mode='both')

            elif choice == '4':
                print("\n" + "="*60)
                print("RUNNING PIPELINE: Organize TV Structure")
                print("="*60)
                run_organizer(folder=tv_path)

            elif choice == '5':
                print("\n" + "="*60)
                print("RUNNING FULL PIPELINE: Organize -> Scan -> Clean")
                print("="*60)
                if tv_path:
                    print("\n[1/3] Organizing TV Structure...")
                    run_organizer(folder=tv_path)
                else:
                    print("\n[1/3] Organizing TV Structure... (Skipped, no TV folder)")
                print("\n[2/3] Scanning library...")
                run_scanner(movies_path=movies_path, tv_path=tv_path, output_file=excel_path)
                print("\n[3/3] Cleaning file and folder names...")
                run_renamer(movies_path=movies_path, tv_path=tv_path, excel_path=excel_path, rename_mode='both')

            elif choice == '7':
                print("\n" + "="*60)
                print("RUNNING PIPELINE: Scan Library")
                print("="*60)
                run_scanner(movies_path=movies_path, tv_path=tv_path, output_file=excel_path)

            else:
                print("Invalid choice. Please enter 0-9.")
                
        except BackNavigationException:
            pass # Already at top menu

def _handle_cli_args(args):
    """Handle non-interactive CLI execution."""
    print(f"🚀 Running non-interactive action: {args.action}")
    
    movies_path = Path(args.movies).resolve() if args.movies else None
    tv_path = Path(args.tv).resolve() if args.tv else None
    excel_path = Path(args.output or "media_library.xlsx").resolve()
    
    if args.action == 'scan':
        run_scanner(movies_path=movies_path, tv_path=tv_path, output_file=excel_path)
    elif args.action == 'organize':
        if not tv_path:
            print("❌ Error: --tv path is required for 'organize' action.")
            sys.exit(1)
        run_organizer(folder=tv_path)
    elif args.action == 'rename-files':
        run_renamer(movies_path=movies_path, tv_path=tv_path, excel_path=excel_path, rename_mode='files')
    elif args.action == 'rename-folders':
        run_renamer(movies_path=movies_path, tv_path=tv_path, excel_path=excel_path, rename_mode='folders')
    elif args.action == 'rename':
        run_renamer(movies_path=movies_path, tv_path=tv_path, excel_path=excel_path, rename_mode='both')
    elif args.action == 'full':
        if tv_path: run_organizer(folder=tv_path)
        run_scanner(movies_path=movies_path, tv_path=tv_path, output_file=excel_path)
        run_renamer(movies_path=movies_path, tv_path=tv_path, excel_path=excel_path, rename_mode='both')
    else:
        print(f"❌ Unknown action: {args.action}")
        sys.exit(1)

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Unified Media Organizer")
    parser.add_argument('--action', choices=['scan', 'organize', 'rename', 'rename-files', 'rename-folders', 'full'], help="Non-interactive action to perform")
    parser.add_argument('--movies', help="Path to movies folder")
    parser.add_argument('--tv', help="Path to TV shows folder")
    parser.add_argument('--output', help="Excel output file name")
    parser.add_argument('--no-emoji', action='store_true', help="Disable emoji output")
    
    args = parser.parse_args()
    
    if args.no_emoji:
        global USE_EMOJIS
        USE_EMOJIS = False
        
    run_wizard(args)

if __name__ == "__main__":
    if hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(encoding='utf-8')  # type: ignore
        except Exception:
            pass
            
    try:
        main()
    except KeyboardInterrupt:
        print("\n⚠️ Interrupted by user.")
    except Exception as e:
        print(f"\n❌ A critical error occurred: {e}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
