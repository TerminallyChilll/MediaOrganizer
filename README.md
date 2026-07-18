# Media Organizer

A cross-platform tool to scan media libraries, organize TV show structures, and cleanly rename media files — powered by [guessit](https://github.com/guessit-io/guessit) (the same filename parser the Sonarr/Bazarr ecosystem relies on), with optional LLM help (Ollama, OpenAI, Gemini) for hopeless filenames.

## Features
- **Clean file names:** `The.Matrix.1999.1080p.BluRay.x264-RARBG.mkv` → `The Matrix (1999) [1080p].mkv`. Handles the hard cases: `WALL-E`, `Se7en`, `Blade Runner 2049`, `1917`, multi-episode `S01E01E02`, date-based shows, `The Office (US)`.
- **Organize TV structures:** loose `S01E01` files/folders are grouped into `Season X` folders; duplicate season folders (`S02` + `Season 2`) are merged; subtitles and `.nfo` files move (and rename) together with their episode.
- **Safe by design:** every change is planned first, previewed, and only applied after you confirm. Collisions are never overwritten — conflicting changes are skipped and reported. `--dry-run` shows the plan without touching anything. Scanning never modifies files.
- **Journaled undo:** every applied change is recorded (with the actual paths) in `mediaorg_journal.jsonl`. Menu option `[7]` or `python run.py --undo` reverses the last run — run it again to unwind earlier runs. Works even after a crash mid-run.
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
python run.py --undo            # reverse the last run
# add --dry-run to any of the above to preview without changing anything
```

### The Fixed-columns workflow
1. Run a scan (`[5]`). Open the `.xlsx`.
2. Anywhere the auto-detected `Title` / `Year` / `Quality` is wrong, type the correct value in the matching `… Fixed` column (or put a complete name in `Folder Fixed` / `File Fixed`).
3. Run the rename (`[1]`) — your values win.

## Requirements
Python 3.11+. Dependencies (`pandas`, `openpyxl`, `tqdm`, `guessit`) are installed automatically by `run.py`.

## Development
```bash
python -m venv .venv && .venv/bin/pip install -r requirements.txt pytest
.venv/bin/python -m pytest
```
