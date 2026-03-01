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

### 📥 1. Download the App
You can download the code using Git in your terminal:
```bash
git clone https://github.com/TerminallyChilll/MediaOganizer.git
cd MediaOganizer
```
*(Alternatively, you can just click the green "Code" button on GitHub and select "Download ZIP")*

---
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
   docker compose run --rm media-organizer
   ```
   *(Note: This uses Docker Compose V2. You must use `run` instead of `up` so you can interact with the terminal prompts!)*

## How to Use the App
After launching the script via one of the methods above, you will be greeted by the **Interactive Wizard**. The wizard will hold your hand through the entire process:

1. **Choose an Action:**
   * **Clean file names:** Renames files to be clean and readable.
   * **Organize file structure:** Groups loose TV episodes into `Season X` folders.
   * **Do both:** Runs both of the above actions back-to-back.
   * **Scan Library:** Just scans your folders and updates the `.xlsx` journal without changing any files.

2. **Select your Media Folders:**
   * You'll be prompted to pick your `Movies` folder and your `TV Shows` folder (you can press Enter to skip one).
   * You can use the built-in folder picker popups, browse directories directly inside the terminal, or paste the folder paths manually.

3. **Choose a Renaming Engine:**
   * **Regex (Fastest):** Strips out year tags, resolution (`1080p`), and scene groups instantly using built-in rules.
   * **Ollama (Free LLM):** If you run Ollama locally, the AI will perfectly extract clean names without breaking a sweat, even on badly obfuscated files.
   * **Cloud LLM (Paid/API):** Connect to OpenAI or Google Gemini if you want cloud-powered AI intelligence.

4. **Review & Execute:**
   * The app will scan your library, generate a preview of all the changes it's about to make, and ask for confirmation before it touches any files.
   * **Safe Undo:** After files are renamed, the app automatically generates an `undo_all.bat` (Windows) or `undo_all.sh` (Mac/Linux) script right next to it. If you ever realize you made a mistake, just run it, and all of your files will be instantly reverted back to their original names and locations!

## Requirements
* Python 3.9 or higher
* `pandas`, `openpyxl`, `tqdm` (automatically installed by `run.py`)
