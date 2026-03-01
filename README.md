# Unified Media Organizer

A cross-platform tool to scan media libraries, organize TV show structures, and cleanly rename media files using Regex or local/cloud LLMs (Gemini, OpenAI, Ollama).

## Features
- **Clean File Names:** Remove scene release garbage, resolution tags, and years to create clean folder and file names.
- **Organize TV Structures:** Automatically grab loose `S01E01` files and group them into appropriate `Season X` folders.
- **Scan & Journal:** Outputs everything to a `.xlsx` spreadsheet so you always have a backup record of your library.
- **Undo Scripts:** Every rename action generates an `undo_all.bat`/`.sh` script in case you don't like the results.
- **Fix Extensions:** Detects missing video extensions (e.g. accidentally stripped `.mkv`) and restores them via magic bytes.
- **LLM Support:** Optionally use local (Ollama) or remote (OpenAI/Gemini) AI to parse incredibly complex/messy filenames perfectly.

## Installation & Usage

### 🪟 Windows
1. Double-click `install_and_run.bat`.
2. The script will automatically check for Python, ensure dependencies are installed, and launch the interactive wizard.

### 🍎 Mac / 🐧 Linux
1. Open a terminal and navigate to this folder.
2. Run: `chmod +x install_and_run.sh`
3. Run: `./install_and_run.sh`
4. The script will install dependencies via pip and launch the wizard.

### 🐳 Docker (Advanced)
If you prefer running inside a container to avoid installing Python or pip packages on your host system:

1. Edit `docker-compose.yml` to map your real media folders to `/media` volume mounts.
2. Run the application interactively:
   ```bash
   docker-compose run --rm media-organizer
   ```
   *(Note: You must use `run` instead of `up` so you can interact with the terminal prompts!)*

## Requirements
* Python 3.9 or higher
* `pandas`, `openpyxl`, `tqdm` (automatically installed by `run.py`)
