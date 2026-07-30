# Media Organizer

A cross-platform tool to scan media libraries, organize TV show structures, and cleanly rename media files — powered by [guessit](https://github.com/guessit-io/guessit) (the same filename parser the Sonarr/Bazarr ecosystem relies on), with optional LLM help (Ollama, OpenAI, Gemini) for hopeless filenames.

## Features
- **Clean file names:** `The.Matrix.1999.1080p.BluRay.x264-RARBG.mkv` → `The Matrix (1999) [1080p].mkv`. Handles the hard cases: `WALL-E`, `Se7en`, `Blade Runner 2049`, `1917`, multi-episode `S01E01E02`, date-based shows, `The Office (US)`.
- **Organize TV structures:** loose `S01E01` files/folders are grouped into `Season X` folders; duplicate season folders (`S02` + `Season 2`) are merged; subtitles and `.nfo` files move (and rename) together with their episode.
- **Safe by design:** every change is planned first, previewed, and only applied after you confirm. Collisions are never overwritten — conflicting changes are skipped and reported. `--dry-run` shows the plan without touching anything. Scanning never modifies files. Nothing is ever deleted: the only destructive primitive is "remove this directory if it is empty".
- **Journaled undo:** every applied change is recorded (with the actual paths, plus the file's size and mtime) in `mediaorg_journal.jsonl`, which lives **next to the app** — not in whatever directory you happened to launch from. `python run.py --list-runs` shows the history; `--undo` reverses the last run, `--undo-run <id>` a specific one, and `--undo-session` the whole of a `--action full`. An *intent* record is written before every change, so a crash or a half-finished copy across drives is detected and cleaned up rather than left to block future runs.
- **Excel journal:** your library is written to an `.xlsx`. Edit the `… Fixed` columns to override any title/year/quality and the next rename pass uses your values.
- **Fix extensions:** restore stripped video extensions via magic-byte detection, or bulk-convert one extension to another.
- **LLM support (optional):** names guessit can't parse are offered to a local (Ollama) or cloud (OpenAI/Gemini) model.

## Installation & Usage

### 1. Download
```bash
git clone https://github.com/TerminallyChilll/MediaOrganizer.git
cd MediaOrganizer
```

### Windows
1. Install **Python 3.11+** from [python.org/downloads](https://www.python.org/downloads/) — check **"Add python.exe to PATH"** during setup.
2. Double-click `install_and_run.bat`.

### Mac / Linux
1. Install **Python 3.11+** (via `brew`, `apt`, …).
2. `chmod +x install_and_run.sh && ./install_and_run.sh`

### Docker
1. Edit `docker-compose.yml` to map your media folders.
2. `docker compose run --rm media-organizer` (use `run`, not `up` — the wizard is interactive).

## How it works
Launching `python run.py` opens the interactive wizard:

```
[1] Clean file names        (scan -> preview -> rename)
[2] Organize TV structure   (loose episodes -> Season folders)
[3] Do both                 (organize -> scan -> rename)
[4] Fix file extensions     (restore missing / bulk convert)
[5] Scan library only       (create/update the Excel spreadsheet)
[6] Export library to text file
[7] Undo last run
```

Every flow follows the same shape: **plan → preview → confirm → apply → journal**. Nothing touches your files until you've seen the full list of changes and said yes.

### Non-interactive use
```bash
python run.py --action scan     --movies /media/Movies --tv /media/TV --output lib.xlsx
python run.py --action organize --tv /media/TV
python run.py --action rename   --output lib.xlsx
python run.py --action full     --movies /media/Movies --tv /media/TV
# add --dry-run to any of the above to preview without changing anything
```

### Undo
```bash
python run.py --list-runs         # what happened, and what can still be reversed
python run.py --undo              # reverse the most recent run
python run.py --undo-session      # reverse every run of the last action
python run.py --undo-last 3       # reverse the newest three runs
python run.py --undo-run a1b2c3d4 # reverse one specific run
python run.py --undo --dry-run    # show what undo would do
```
Undo refuses to move a file back if it was modified or replaced after the fact,
and refuses to reverse a run out of order when a newer run touched the same
paths. `--force` overrides either check. The journal location can be pinned
with the `MEDIAORG_JOURNAL` environment variable.

### The Fixed-columns workflow
1. Run a scan (`[5]`). Open the `.xlsx`.
2. Anywhere the auto-detected `Title` / `Year` / `Quality` is wrong, type the correct value in the matching `… Fixed` column (or put a complete name in `Folder Fixed` / `File Fixed`).
3. Run the rename (`[1]`) — your values win.

## Cross-platform notes
The tool runs on Windows, Linux and macOS, and CI runs the test suite on all
three. A few filesystem realities are worth knowing, because they explain
messages you may see:

- **Case-insensitive volumes** (Windows/NTFS, macOS/APFS, and any exFAT or NTFS
  drive mounted on Linux). Case-only renames such as `wall-e` → `WALL-E` go
  through a temporary name so they work on those volumes, and collision
  detection probes the destination volume rather than assuming the host's
  behaviour.
- **macOS filename normalization.** macOS may hand back decomposed (NFD)
  filenames where the rename scheme produces composed (NFC) ones. All name
  comparisons normalize first, so accented titles like `Amélie` settle after
  one pass instead of being "renamed" on every run.
- **OS junk files.** `._AppleDouble` sidecars, `.DS_Store`, `Thumbs.db` and
  `desktop.ini` are ignored as media, and if one is all that is left in a
  folder being cleaned up it is moved to a `.mediaorg_trash` folder (a normal,
  reversible move) instead of blocking the cleanup.
- **Long paths.** Windows rejects paths over 260 characters unless
  `LongPathsEnabled` is set. Destinations that would exceed the limit are
  skipped up front with a clear reason rather than failing partway through.
- **Moves across drives** are byte copies, not renames. The preview says so,
  each copy is size-verified before the original is removed, and an interrupted
  copy is cleaned up rather than left behind.
- **Files in use** (Plex/Jellyfin/VLC streaming an episode) are retried briefly
  and then reported by name — on Windows this is the most common cause of a
  partially-renamed show.
- **Docker:** the journal records container paths, so run undo inside the
  container, not on the host.

## Requirements
Python 3.11+. Dependencies (`pandas`, `openpyxl`, `tqdm`, `guessit`) are installed automatically by `run.py`.

## Development
```bash
python -m venv .venv
# Linux / macOS
.venv/bin/pip install -r requirements-dev.txt && .venv/bin/python -m pytest
# Windows
.venv\Scripts\pip install -r requirements-dev.txt && .venv\Scripts\python -m pytest
```
`python run.py --doctor` diagnoses environment problems and reports any
interrupted changes; `--doctor --fix` attempts repair.
