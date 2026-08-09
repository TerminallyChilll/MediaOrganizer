# Media Organizer

A cross-platform tool to scan media libraries, organize TV show structures, and cleanly rename media files — powered by [guessit](https://github.com/guessit-io/guessit) (the same filename parser the Sonarr/Bazarr ecosystem relies on), with optional LLM help (Ollama, OpenAI, Gemini) for hopeless filenames.

> **Prefer plain text?** [README.txt](README.txt) is this same document with the
> formatting removed — double-click it and it opens in Notepad or any text
> editor. ([REVIEW.txt](REVIEW.txt) likewise.) Both are generated; see
> [Development](#development).

## Features
- **Clean file names:** `The.Matrix.1999.1080p.BluRay.x264-RARBG.mkv` → `The Matrix (1999) [1080p].mkv`. Handles the hard cases: `WALL-E`, `Se7en`, `Blade Runner 2049`, `1917`, multi-episode `S01E01E02`, date-based shows, `The Office (US)`.
- **Organize TV structures:** loose `S01E01` files/folders are grouped into `Season X` folders; duplicate season folders (`S02` + `Season 2`) are merged; subtitles and `.nfo` files move (and rename) together with their episode.
- **Handles nested libraries:** shows are found wherever they sit — `TV/Show`, `TV/Genre/Show`, `TV/Genre/SubGenre/Show` — and season folders are flattened however deep the episode is buried (`Season 1/Disc 1/ep.mkv`, `Season 1/Show.S01E01/Subs/ep.mkv`). Episodes dumped in a non-season subfolder (`Show/Downloads/Show.S01E01.mkv`) are routed into the right season folder, one season at a time. See [Nested folders](#nested-folders).
- **Safe by design:** every change is planned first, previewed, and only applied after you confirm. Collisions are never overwritten — conflicting changes are skipped and reported. `--dry-run` shows the plan without touching anything. Scanning never modifies files. Nothing is ever deleted: the only destructive primitive is "remove this directory if it is empty".
- **Review before, decide after:** the preview is a numbered, paged working list, not a wall of text — page through every change, exclude anything you don't want (`x 3,7-9`), or retype a name that came out wrong (`e 3`, and its subtitles follow). Enter pages; it never applies. Then, once the changes are on disk, the wizard stops and asks you to go look at them. Say no and every file goes straight back where it was. See [Review and accept](#review-and-accept).
- **Journaled undo:** every applied change is recorded (with the actual paths, plus the file's size and mtime) in `mediaorg_journal.jsonl`, which lives **next to the app** — not in whatever directory you happened to launch from. `python run.py --list-runs` shows the history; `--undo` reverses the last run, `--undo-run <id>` a specific one, and `--undo-session` the whole of a `--action full`. An *intent* record is written before every change, so a crash or a half-finished copy across drives is detected and cleaned up rather than left to block future runs.
- **Excel journal:** your library is written to an `.xlsx`. Edit the `… Fixed` columns to override any title/year/quality and the next rename pass uses your values.
- **Fix extensions:** restore stripped video extensions via magic-byte detection, or bulk-convert one extension to another.
- **Self-updating:** the wizard tells you how many commits behind you are and prints the command (and the folder to run it in) to catch up — or press `[U]` and it updates itself. See [Updating](#updating).
- **LLM support (optional):** names guessit can't parse are offered to a local (Ollama) or cloud (OpenAI/Gemini) model. Configure it with environment variables (`OLLAMA_URL`, `OPENAI_API_KEY`, `GEMINI_API_KEY`) or answer the prompt once — see [LLM setup](#llm-setup).

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

## Updating
Launching the wizard tells you when your copy is behind, and exactly what to
type to catch up:

```
----------------------------------------------------------------------
  Update available: you are 3 commits behind origin/main.
  (you have a1b2c3d, latest is e4f5g6h)

  To update, run this in a terminal in the MediaOrganizer folder:
      cd "C:\Users\you\MediaOrganizer"
      python run.py --update

  ...or press [U] here and the wizard will do it for you.
----------------------------------------------------------------------
```

So there are two ways to update — pick either:

- **From the menu:** press **`[U]`**. It updates and then asks you to relaunch.
- **From a terminal**, in the folder you cloned into:
  ```bash
  cd path/to/MediaOrganizer     # the folder containing run.py
  python run.py --update        # Windows: py run.py --update
  ```

Other commands:
```bash
python run.py --version        # which version and commit you are on
python run.py --check-update   # how far behind you are, and what you're missing
python run.py --update --dry-run   # show what would be pulled, change nothing
python run.py --update --yes       # skip the confirmation (required when
                                   #   run unattended - see below)
```

### Coming from a pre-updater install
Self-updating arrived in **v2.1.0**. An install older than that has no `[U]`
key and no `--update` flag, so it cannot pull its own updater — catching up is
a one-time manual step, after which everything above applies.

**1. Find out which kind of install you have.** In the folder containing
`run.py`:

```bash
cd path/to/MediaOrganizer
git status
```

If that says `not a git repository`, you have a ZIP download — skip to step 5.
Otherwise carry on.

**2. Put back what the old launcher deleted.** Launchers before v2.0.0 ran
`rm -f install_and_run.bat Dockerfile docker-compose.yml` (and the Windows
equivalent) on every single run, to save a few KB. Those files have changed
since, so git will not merge over the deletions and the pull fails before it
starts. If `git status` lists deleted files:

```bash
git checkout -- .
```

If it still reports changes — or you edited something yourself and don't want
to keep it — use `git reset --hard` instead. Both throw away local
modifications, so read `git status` first if you might have changed something
on purpose.

**3. Pull.**

```bash
git pull
```

**4. Launch it once.** `run.py` installs anything missing or outdated:

```bash
python run.py        # Windows: py run.py
```

That's the whole thing — `[U]` and `--update` work from here on.

**5. ZIP downloads only: re-clone.** Updating needs a real clone, so make one
and bring your state across by hand. Your journal, settings and word list sit
*inside* the old folder — they are untracked, which is why `git pull` never
touches them, and also why a fresh clone doesn't have them:

```bash
git clone https://github.com/TerminallyChilll/MediaOrganizer.git MediaOrganizer-new
cd MediaOrganizer-new
cp ../MediaOrganizer/mediaorg_journal*.jsonl .      # undo history
cp ../MediaOrganizer/custom_strip_patterns.json .   # your word list
cp ../MediaOrganizer/.media_llm_config.json .       # saved API key
cp ../MediaOrganizer/.media_renamer_config.json .   # remembered folders
cp ../MediaOrganizer/*.xlsx .                       # scans + Fixed-column edits
```

Skip any that don't exist. On Windows use `copy`, and note that the files
starting with a dot are hidden in Explorer. Keep the old folder until the new
one has run once.

The `.xlsx` files matter more than they look: the `… Fixed` columns you typed
live in the workbook and nowhere else, so leaving them behind means the next
scan writes a fresh one and your corrections are gone. If you kept scans
somewhere other than the app folder, copy them from there instead.

**6. Check it worked.**

```bash
python run.py --version        # 2.1.0 or newer
python run.py --check-update   # should say you are up to date
```

If anything looks wrong: `python run.py --doctor --fix`.

**What the update does:** fetches, fast-forwards your clone to the latest
commit, and reinstalls dependencies only if `requirements.txt` changed. It
never discards your work — if you have edited files locally, or made your own
commits, it stops and prints the git command to resolve that first. Your
library, journal, spreadsheets and word list are untracked files and are never
touched.

Like every other flow in this app, it asks before it changes anything. If there
is no terminal to ask (a cron job, a script, `--update < /dev/null`) it stops
and tells you to opt in with `--yes` rather than deciding for you. Ctrl-C or
Ctrl-D at the prompt means "no".

For scripting, `--update` exits:

| Code | Meaning |
| --- | --- |
| `0` | Updated, already current, previewed with `--dry-run`, or you declined |
| `1` | Refused (local edits, diverged, not a clone, no terminal), the pull failed, or the code came down but the dependency install did not |
| `2` | The arguments could not be understood |

**Notes**
- Updating needs `git` and an install made with `git clone`, in a folder that
  is the clone itself. If you downloaded a ZIP instead — or dropped the folder
  inside some other project's repository, where it would otherwise update
  *that* project — the app says so and prints the `git clone` command to switch
  to a real clone. Your journal, settings and word list live in the folder you
  are replacing, so copy them across as well — see [Coming from a pre-updater
  install](#coming-from-a-pre-updater-install).
- Updates follow the branch you are on. On `main` that is `origin/main`; on
  your own branch it is whatever that branch tracks. A branch that tracks
  nothing is reported as such rather than quietly fast-forwarded onto `main`.
- **The app contacts github.com.** On launch it runs a `git fetch` in the
  background, at most once a day (and on a first launch, before any cache
  exists); `--check-update` and `--doctor` each fetch once when you run them.
  No other data leaves the machine, and nothing is sent about your library. On a media server, an air-gapped box or any host with restricted
  egress, turn it off:
  ```bash
  MEDIAORG_NO_UPDATE_CHECK=1 python run.py     # this launch
  setx MEDIAORG_NO_UPDATE_CHECK 1              # Windows, permanently
  ```
  With it off nothing is sent on its own, and `--doctor` compares against
  what was last fetched instead of reaching out. `python run.py --update`
  still works whenever you ask for it explicitly.
- The answer is cached next to the app (`.mediaorg_update_check.json`), so a
  normal launch is offline and instant. `MEDIAORG_UPDATE_INTERVAL=6` checks
  every 6 hours instead of every 24.
- No update available means no message — the notice only appears when there is
  actually something to get.

## How it works
Launching `python run.py` opens the interactive wizard:

```
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
```

Every flow that changes anything follows the same shape:

**pick the folder → pick what to do → check the word list → choose the naming
engine (guessit or an LLM) → see the full before/after list → review it →
apply → keep it or have it all put back.**

Nothing touches your files until you've seen the full list of changes and said
yes, and nothing *stays* changed until you've looked at the result and said yes
again. `[6]`, `[7]` and `[8]` never touch your media at all.

### Review and accept

Every change is on screen before anything is asked. The list is numbered and
paged, the current path above the proposed one:

```
--- renames: items 1-10 of 14, page 1/2 (14 will be applied) ---
  [ 1] BEFORE  /media/TV/My Show/My.Show.S01E01.mkv
       AFTER   /media/TV/My Show/Season 1/My Show S01E01.mkv
  [ 2] BEFORE  /media/TV/My Show/My.Show.S01E02.mkv
       AFTER   /media/TV/My Show/Season 1/My Show S01E02.mkv
  [Enter/N]ext  [P]rev  [G] page  [A]ll  [R] review & edit  [Y]es apply  [Q]uit
```

Enter pages forward — it never applies anything, so a stray keypress cannot
commit a rename you have not read. `[Y]` applies, `[Q]` cancels.

`[R]` turns on editing, on the same list:

```
  [Enter/N]ext  [P]rev  [G] page  [A]ll  [x N] exclude  [k N] keep  [e N] rename  [Y]es apply  [Q]uit
  (x and k take ranges too: 'x 3,7-9')
```

- `x 3` drops a change; `x 3,7-9` drops several; `k 3` puts one back.
- `e 3` lets you retype that one name. Subtitles and `.nfo` sidecars follow the
  video automatically, keeping any `.en`/`.fr` language tag; changing the
  extension is questioned rather than done silently; and a name that isn't
  usable on every filesystem is offered back cleaned up. You are editing the
  *name*, not the folder it lands in, so slashes are refused. Renaming a
  subtitle does **not** rename the film it belongs to.
- Anything you change is re-checked for collisions before it runs, exactly like
  a name the tool generated itself. A folder cleanup your exclusion made
  impossible is dropped rather than left to fail, and so is a folder that would
  now be created empty.
- Anything the planner had to skip is listed with the reason, so a conflict is
  never just a number.

After the changes are applied the wizard stops and tells you to go look:

```
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  CHECK YOUR FILES NOW - 14 change(s) are already on disk.
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
...
Keep these changes? (y/n) [n]:
```

Answer **no** — or just press Enter, which is the default — and every file goes
back exactly where it was, through the same journal `--undo` uses. Answer yes
and they stay; you can still reverse them later, you just have to ask.

`[4] Do it all` asks this **once, for the whole action**: saying no there puts
back the TV organizing as well as the renaming, not only the last step.

The folders you pick are remembered between runs in
`.media_renamer_config.json`, kept next to the app (override with
`MEDIAORG_CONFIG`) so that launching from any directory offers the same
defaults — like the journal and the word list.

### Non-interactive use
```bash
python run.py --action scan            --movies /media/Movies --tv /media/TV --output lib.xlsx
python run.py --action organize        --tv /media/TV
python run.py --action organize-movies --movies /media/Movies
python run.py --action rename          --output lib.xlsx
python run.py --action full            --movies /media/Movies --tv /media/TV
python run.py --action inventory       --path /media --output inventory.xlsx
# add --dry-run to any of the above to preview without changing anything
# add --review to be asked, afterwards, whether to keep the changes
```

Run from a terminal, these behave like the wizard: the full list first, then a
`[Y]`. They do **not** ask the keep-or-put-back question afterwards — pass
`--review` for that, and with `--action full` it covers every phase at once.

**When stdin is not a terminal** — cron, a Docker service, anything piped —
there is nobody to answer either question, so the plan is printed to the log
and applied without prompting. The journal is still written, and the exact
`--undo-run` command is printed, so an unattended run is as reversible as any
other. `--dry-run` still changes nothing at all.

### Two ways to list what you have
Both are read-only.

- **`[6] Scan library only`** understands media. It writes the `Movies` and
  `TV Shows` sheets the renamer reads, so this is the one to run before
  editing the `… Fixed` columns.
- **`[7] Inventory every file`** understands nothing. It records *every* file
  under a folder — video, subtitles, artwork, a stray PDF — one row each with
  path, extension, type, size and last-modified date, as `.xlsx`, `.csv` or a
  plain `.txt` tree. Use it for an audit, a backup list, or a before/after
  comparison.

### Custom word list
The rename flows (`[1]`, `[4]`) and the scan (`[6]`) show you this list on the
way past and offer to edit it, so you do not have to know to visit `[8]` first.
`[8]` manages the words stripped out of a name before it is parsed — release
groups, tracker tags, whatever your library is littered with. You can add
entries, **remove them individually, or clear the list**, and test a pattern
against a real filename before committing to it. Each entry is a
case-insensitive regular expression; a typo that isn't valid regex is offered
back escaped, and a pattern that would match whole names is refused rather
than silently ignored. The list lives next to the app (override with
`MEDIAORG_PATTERNS`), so it follows you regardless of where you launch from.

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

These are for changes you accepted earlier and have since thought better of.
For the run you just made, the wizard already offers to reverse it — see
[Review and accept](#review-and-accept) — so you rarely need to reach for these
in the same sitting.

### Nested folders
A **show folder** is the shallowest folder that directly contains season
folders (`Season 1`, `S01`, `Specials`) or `SxxEyy` episode files. Everything
above it is treated as a wrapper and left alone, so `TV/Genre/SubGenre/Show`
is organized in place rather than being flattened into the library root. The
search descends four levels below the folder you select; shows nested deeper
than that are not found.

Inside a show, a season folder is expected to hold episodes and nothing else,
so anything nested below one is lifted out and the emptied folders removed:

```
Show/Season 1/Disc 1/ep.mkv           ->  Show/Season 1/ep.mkv
Show/Season 1/Show.S01E01/Subs/ep.mkv ->  Show/Season 1/ep.mkv
Show/Downloads/Show.S02E03.mkv        ->  Show/Season 2/Show.S02E03.mkv
```

A file is only ever moved when something says where it goes — its own episode
code, or the code in the name of the folder holding it. Nothing is placed by
position alone, so `Artwork/poster.jpg` stays put, and `Season 1/Extras/
Show.S02E05.mkv` goes to `Season 2`, not to the season folder it happened to
be sitting in. Aspect-ratio names are not read as episode codes:
`banner-16x9.jpg` is artwork, not season 16.

These are left alone entirely, even when the files inside them carry episode
codes: local-extras folders (`Specials`, `Extras`, `Trailers`, `Featurettes`,
`Behind The Scenes`, `Deleted Scenes`, `Interviews`, `Bonus`, `Other`), and
OS/NAS bookkeeping directories (`@eaDir`, `.trashes`, `$RECYCLE.BIN`, …). A
folder is only removed once everything in it has been moved out.

**Wrappers vs. shows.** A folder that contains a show is a wrapper and is left
where it is. That is decided by what is on disk, not by the folder's name, so
two flat shows under one bucket stay two shows:

```
Genre/ShowA/ShowA.S01E01.mkv   ->  Genre/ShowA/Season 1/…
Genre/ShowB/ShowB.S01E01.mkv   ->  Genre/ShowB/Season 1/…
```

**One case the tool cannot resolve for you.** `Show/Downloads/Show.S01E01.mkv`
and `Genre/Show/Show.S01E01.mkv` are the same shape on disk — nothing in the
names says which folder is the show. The inner one always wins, so a show
behind a wrapper is found correctly, and a dump folder is named as the show
(`Downloads`, `Complete Series`). Fix those by typing the real name in the
`Folder Fixed` column and re-running the rename.

### LLM setup
Optional, and only ever consulted for names guessit could not parse (those
show up as `source="raw"`); you can also ask for every name to be sent. The
renaming-engine prompt appears during `[1]` and `[4]`.

| Variable | Used for |
|---|---|
| `OLLAMA_URL` | Ollama endpoint, default `http://localhost:11434` |
| `OLLAMA_MODEL` | Preselects a model instead of choosing from the list |
| `OPENAI_API_KEY` | OpenAI (`gpt-4o-mini`) |
| `GEMINI_API_KEY` | Google Gemini (`gemini-2.0-flash`) |

The environment takes precedence over a saved key, which is what makes the
Docker services work — `docker-compose.yml` passes all of these through. A
key supplied that way is **never written to disk**. A key you type at the
prompt is saved to `.media_llm_config.json` (mode `0600`, next to the app,
override with `MEDIAORG_LLM_CONFIG`) so you only type it once; the tool tells
you when it does this. An unset variable is ignored rather than blanking a
key you already saved.

Local models are batched 15 names at a time (40 for cloud) and asked for JSON
specifically, and the response parser copes with the usual local-model output
— markdown fences, single quotes, trailing commas, a wrapper object, or
alternative field names. If a batch fails, the run continues with the names
that did come back.

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

Every `.md` file at the repo root ships a generated plain-text twin, so anyone
who does not know what a Markdown file is can still read the instructions. After
editing a `.md`, regenerate:

```bash
python tools/md_to_txt.py           # rewrite the .txt copies
python tools/md_to_txt.py --check   # verify they are current (CI does this)
```

Never edit a `.txt` by hand — the next regenerate discards it, and
`tests/test_docs.py` fails the build if the two ever disagree.
