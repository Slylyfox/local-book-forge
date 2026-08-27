# Local Book Forge — Dashboard

A local control panel for the book generator pipeline. Instead of juggling
terminal windows for Ollama, AUTOMATIC1111, and the generator script, you
get one page with buttons and a live terminal output panel — plus a genre
dropdown so runs are deliberate instead of random.

Everything still runs as real local processes on your machine. The
dashboard is just a thin local web UI (Flask) sitting in front of them —
nothing leaves your computer.

## How it works

- `dashboard_server.py` — a small Flask server that runs on `127.0.0.1:8765`.
  When you click a button in the browser, it starts the real command
  (e.g. `ollama serve`, `webui-user.bat`, or
  `python local-book-generator.py --genre "epic fantasy"`) as a subprocess
  and streams its stdout back to the page live via Server-Sent Events —
  so you see the same output you'd see in a terminal, just inside the page.
- `templates/dashboard.html` — the page itself: status pills for Ollama/A1111,
  launch buttons, a genre dropdown, and the terminal panel.
- `local-book-generator.py` — your existing pipeline script, with one change:
  it now accepts `--genre "<genre>"` on the command line. If you omit it (or
  select "Random" in the dashboard), it falls back to the original random
  behavior.

## One-time setup

1. Put `dashboard_server.py`, the `templates/` folder, and
   `local-book-generator.py` in the same directory (they already are, if
   you keep this folder structure).

2. Install the dashboard's own tiny dependency (separate from your book
   generator's `requirements.txt`, which you should already have installed
   in your venv):
   ```
   pip install -r requirements_dashboard.txt
   ```
   You can install this into the same venv you use for the generator, or
   into your system Python — the dashboard server itself doesn't touch
   Ollama, A1111, or CrewAI directly, it just launches your existing
   Python (from whichever venv you point it at) as a subprocess.

3. Start the dashboard:
   ```
   python dashboard_server.py
   ```
   Then open **http://127.0.0.1:8765** in your browser.

4. In the dashboard's **Settings** panel, fill in and save:
   - **Python executable** — full path to the Python you use for the
     generator, e.g. `C:\path\to\your\project\venv\Scripts\python.exe`.
     If you don't set this, it defaults to whatever Python the dashboard
     server itself is running under, which may not have `crewai`, `docx`,
     etc. installed.
   - **local-book-generator.py path** — full path to the script.
   - **webui-user.bat path** — full path to AUTOMATIC1111's launcher, e.g.
     `C:\stable-diffusion-webui\webui-user.bat`.
   - **Ollama executable** — usually just `ollama` if it's on your PATH;
     otherwise the full path to `ollama.exe`.

   These are saved to `dashboard_config.json` next to the server, so you
   only need to do this once.

## Day-to-day use

1. Open the dashboard, check the **Dependencies** panel:
   - Green "running" pill = detected and reachable.
   - Click **Launch** next to a service to start it (its boot log streams
     into the terminal panel on the right — for A1111 this can take a
     minute or two on first load as it loads the checkpoint).
2. Pick a **genre** from the dropdown (or leave it on Random) and click
   **Run**.
3. Watch the terminal panel — it's the exact same stdout you'd see running
   the script directly, just streamed into the page: outline generation,
   each chapter as it's written, cover generation, and manuscript assembly.
4. Click **Stop** if you need to kill a run mid-way (e.g. the outline JSON
   looks broken, or a chapter has clearly gone off the rails).

Output files land in the same place as before:
`output_books/<book-title>-<timestamp>/`.

## Notes & caveats

- **This is single-user and unauthenticated by design.** It binds to
  `127.0.0.1` only (not `0.0.0.0`), so it isn't reachable from other
  machines on your network — keep it that way. Don't expose port 8765
  externally.
- **Launching AUTOMATIC1111 via the dashboard** runs `webui-user.bat`
  through `cmd /c` with output captured, rather than opening its own
  console window, so everything funnels into one place. A1111 keeps
  running in the background after its boot log finishes streaming — you
  don't need to keep watching it.
- **Ollama** usually auto-starts as a background service after install; the
  **Check** button is what you'll use most. **Launch** runs `ollama serve`
  explicitly, which is mainly useful if it isn't already running as a
  service.
- **Streaming granularity**: the generator script now force-enables
  line-buffered stdout, so its own `print()` statements stream immediately.
  CrewAI's internal logging may still arrive in slightly larger bursts
  depending on how it buffers internally — this is a CrewAI-side detail,
  not something the dashboard can fully control.
- If you ever want to fall back to plain terminal usage, nothing about the
  script changed except the addition of `--genre`; running it with no
  arguments still works exactly as before (random genre).

## Ideas for later, if useful

- A history panel that lists past runs from `output_books/` with quick
  links to open `manuscript.docx` / `cover.png`.
- A "dry run" button that only runs Stage A (outline) so you can review the
  premise/outline JSON before committing to the slow chapter-writing stage.
- Per-genre default word-count/chapter-count overrides in the dashboard
  instead of editing `MIN_CHAPTERS`/`MAX_CHAPTERS`/`MAX_TOTAL_WORDS` in the
  script directly.
- A small badge showing which SDXL checkpoint is currently loaded in A1111
  (via `GET /sdapi/v1/options`), so you can confirm it before a long run.

Happy to build any of these next if they'd help.
