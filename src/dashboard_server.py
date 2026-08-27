"""
Local Book Forge — Dashboard Server
------------------------------------
A small local-only Flask server that gives the dashboard.html page real
buttons that do real things: check/launch Ollama, check/launch AUTOMATIC1111,
and run local-book-generator.py with a chosen genre — all while streaming
the actual terminal output back to the browser live.

This is meant to run ONLY on your own machine (127.0.0.1). It happily runs
arbitrary local commands you've configured, so don't expose it to a network.

Run it with:
    python dashboard_server.py

Then open:
    http://127.0.0.1:8765
"""

import json
import os
import queue
import random  # added 2026-08-14 for the Book Structure "Random" preset roll
import re
import shutil
import subprocess
import sys
import threading
import time
import uuid
import zipfile

from flask import Flask, Response, jsonify, request, send_file, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "dashboard_config.json")
LOGS_DIR = os.path.join(BASE_DIR, "logs")
os.makedirs(LOGS_DIR, exist_ok=True)

app = Flask(__name__, static_folder=None)

# Windows-only flag so launched processes don't pop open their own console
# window — everything is captured and shown in the dashboard's terminal
# panel instead. On non-Windows this constant doesn't exist, so guard it.
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

GENRE_POOL = [
    "sci-fi thriller",
    "cozy mystery",
    "epic fantasy",
    "romantic suspense",
    "psychological horror",
    "dystopian survival",
    "noir detective",
    "space opera",
]

# Every Ollama tag the pipeline actually requires (dashboard section 6.9's
# system test checks these are all pulled/built). Duplicated here rather than
# imported — same "each script runs standalone" convention as the pipeline
# scripts themselves — so if these ever change over there, update here too.
REQUIRED_OLLAMA_MODELS = [
    "llama3.1",                                    # outline (all scripts) + editor/scorer base
    "llama3.1-16k",                                # editor/scorer + writer 16k-context alt (6.12)
    "natsumura-storytelling-rp-llama-3.1-16k",      # default writer model
]

# Section 6.12 model-switch buttons — fixed set per product decision, 2026-08-13.
# "" means omit --writer-model entirely (script's own default: Natsumura).
# Deliberately NOT offering bare "llama3.1" as the comparison model — that's
# the un-fixed 2048-token-context base pull section 5 identified as the root
# cause of late-chapter coherence breakdown, so a button for it would let a
# one-click A/B test silently reintroduce a bug that took real debugging to
# find. llama3.1-16k (already built for the editor/scorer, see section 5.2)
# gives the same "stock instruct model" comparison point without that risk.
WRITER_MODEL_OPTIONS = [
    {"value": "", "label": "Natsumura (default)"},
    {"value": "llama3.1-16k", "label": "Llama 3.1 (stock, 16k — comparison)"},
]
WRITER_MODEL_VALUES = {opt["value"] for opt in WRITER_MODEL_OPTIONS}

# =====================================================================
# BOOK STRUCTURE PRESETS (2026-08-14)
# =====================================================================
# Length was previously fixed in local-book-generator.py's constants, which meant
# changing it required editing code. Now it's a dashboard panel, persisted to
# dashboard_config.json and forwarded to the generator as CLI flags.
#
# Note MAX_CHAPTER_TARGET_WORDS stays at 3000 in every preset including
# Full-length. Extra length comes from more chapters, never longer ones —
# measured on real runs, chapters past ~3,000 words are where repetition
# collapse starts appearing. A "longer book" preset that raised the per-chapter
# ceiling would be trading the exact quality problem this pipeline just spent a
# day fixing for a bigger word count.
BOOK_STRUCTURE_PRESETS = {
    "novelette": {
        "label": "Novelette (12-20k)",
        "min_chapters": 5, "max_chapters": 8,
        "min_total_words": 12000, "max_total_words": 20000,
        "min_chapter_words": 1500, "max_chapter_words": 3000,
        "note": "Finished length after editing. ~30-50 min. Lands in the Short band.",
    },
    "novella": {
        "label": "Novella (20-32k)",
        "min_chapters": 8, "max_chapters": 12,
        "min_total_words": 20000, "max_total_words": 32000,
        "min_chapter_words": 1800, "max_chapter_words": 3000,
        "note": "Finished length after editing. ~50-80 min. Lands in the Standard band.",
    },
    "full": {
        "label": "Full-length (32-40k)",
        "min_chapters": 12, "max_chapters": 15,
        "min_total_words": 32000, "max_total_words": 40000,
        "min_chapter_words": 2000, "max_chapter_words": 3000,
        "note": "Finished length after editing. ~80-110 min. Lands in the Full-length band.",
    },
}
# =====================================================================
# EDITORIAL SHRINKAGE ALLOWANCE (2026-08-14)
# =====================================================================
# The word counts in the presets above are what you want the FINISHED book to
# be. The outline has to be told a bigger number, because the editorial pass
# legitimately removes wordiness on its way through.
#
# Measured on two real runs with the current editor:
#     Rainy Night Requiem   planned 19,994 -> delivered 19,056  (95%)
#     Fractured Earth       planned 21,150 -> delivered 18,313  (87%)
#
# Both books were aimed at the 20,000-word Standard boundary and both landed
# under it — by 944 and 1,687 words — so both came out a band short. The presets
# were asking the outline for the finished length and then losing 5-13% of it
# downstream.
#
# 0.87 is the worst retention observed, used deliberately rather than the 91%
# average: overshooting the band costs nothing (a slightly longer book is still
# in the same band), while undershooting it costs a dollar a sale.
#
# Keeping the compensation here rather than baking bigger numbers into the
# presets means the panel keeps showing honest finished lengths, and if the
# editor's behaviour changes there is exactly one number to update.
EDITOR_RETENTION = 0.87

STRUCTURE_FIELDS = ("min_chapters", "max_chapters", "min_total_words",
                    "max_total_words", "min_chapter_words", "max_chapter_words")


def resolve_book_structure(preset, custom=None):
    """Turn a preset name (or 'random', or 'custom' + explicit values) into a
    concrete dict of the six structure numbers.

    'random' rolls a preset. Called once PER BOOK rather than per batch, per
    Decision 2026-08-14 — so a 3-book batch can produce one novelette, one
    novella and one full-length, which is closer to how a real imprint's list
    looks than three books of identical length."""
    if preset == "custom" and custom:
        # Custom values are taken literally. If you typed a number, that number
        # is what the outline is asked for — second-guessing an explicit setting
        # would be worse than the shrinkage it corrects for.
        return {f: int(custom[f]) for f in STRUCTURE_FIELDS if custom.get(f) is not None}
    if preset == "random" or preset not in BOOK_STRUCTURE_PRESETS:
        preset = random.choice(list(BOOK_STRUCTURE_PRESETS))
    chosen = BOOK_STRUCTURE_PRESETS[preset]
    resolved = {f: chosen[f] for f in STRUCTURE_FIELDS}

    # Inflate the word budget so the FINISHED book lands where the preset says.
    # See EDITOR_RETENTION for the measurements behind this.
    resolved["min_total_words"] = int(round(chosen["min_total_words"] / EDITOR_RETENTION))
    resolved["max_total_words"] = int(round(chosen["max_total_words"] / EDITOR_RETENTION))

    # The inflated ceiling can exceed what the chapter budget can actually
    # produce, which would leave an unsatisfiable range. Clamp it to what's
    # reachable, and keep the floor below it.
    reachable = resolved["max_chapters"] * resolved["max_chapter_words"]
    resolved["max_total_words"] = min(resolved["max_total_words"], reachable)
    resolved["min_total_words"] = min(resolved["min_total_words"],
                                      resolved["max_total_words"])
    return resolved


def validate_book_structure(s):
    """Return a human-readable error, or None. Mirrors the generator's own checks
    so a bad combination is rejected at the click rather than after a run starts
    and dies at the outline stage."""
    if s["min_chapters"] > s["max_chapters"]:
        return f"Min chapters ({s['min_chapters']}) is above max chapters ({s['max_chapters']})."
    if s["min_chapter_words"] > s["max_chapter_words"]:
        return (f"Min words per chapter ({s['min_chapter_words']}) is above max "
                f"({s['max_chapter_words']}).")
    if s["min_total_words"] > s["max_total_words"]:
        return (f"Min total words ({s['min_total_words']:,}) is above max "
                f"({s['max_total_words']:,}).")
    reachable = s["max_chapters"] * s["max_chapter_words"]
    if s["min_total_words"] > reachable:
        return (f"Min total words ({s['min_total_words']:,}) can't be reached: at most "
                f"{s['max_chapters']} chapters x {s['max_chapter_words']:,} words = "
                f"{reachable:,}. Raise max chapters or max words per chapter.")
    if any(s[f] < 1 for f in STRUCTURE_FIELDS):
        return "All structure values must be positive."
    return None


def structure_to_args(s):
    """Structure dict -> CLI flags for local-book-generator.py / pipeline_chain.py."""
    flags = []
    for field in STRUCTURE_FIELDS:
        if s.get(field) is not None:
            flags += [f"--{field.replace('_', '-')}", str(s[field])]
    return flags


DEFAULT_CONFIG = {
    "python_exe": sys.executable,
    "script_path": os.path.join(BASE_DIR, "local-book-generator.py"),
    "pipeline_script_path": os.path.join(BASE_DIR, "pipeline_chain.py"),
    "a1111_bat_path": "",  # e.g. C:\\stable-diffusion-webui\\webui-user.bat
    "ollama_exe": "ollama",  # assumes it's on PATH; change to full path if not
    "ollama_url": "http://localhost:11434",
    "a1111_url": "http://127.0.0.1:7860",
    # Book Structure panel state. "random" is the default so a fresh install
    # produces a varied catalogue rather than N identical-length books; the panel
    # persists whatever you last chose.
    "book_structure_preset": "random",
    "book_structure_custom": dict(BOOK_STRUCTURE_PRESETS["novella"],
                                  **{k: v for k, v in ()}),
}
# Strip the label/note out of the persisted custom block — those are display
# strings that belong to the preset definition, not settings worth saving.
DEFAULT_CONFIG["book_structure_custom"] = {
    f: BOOK_STRUCTURE_PRESETS["novella"][f] for f in STRUCTURE_FIELDS
}


def load_config():
    if os.path.isfile(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            merged = dict(DEFAULT_CONFIG)
            merged.update(cfg)
            return merged
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)


# =====================================================================
# Section 6.14 — GPU-task ETA / countdown timers.
#
# Two independent estimates, both driven by the SAME historical-duration
# idea (median of recent past runs, not a hardcoded guess):
#
#   - "task" ETA: how long the CURRENT stage (outline / chapter N write /
#     cover gen / editorial pass / scoring pass / ...) is expected to take,
#     keyed by the stage's stable `stage=` key from the [TIMING] marker
#     lines the pipeline scripts print (see their _emit_timing() helpers).
#   - "book" ETA: how long the WHOLE current command (one book's worth of
#     work) is expected to take, keyed by a normalized job "kind" (Full
#     Pipeline vs. Book Generator) rather than the exact job label, so a
#     batch run's per-book history isn't scattered across N differently-
#     named buckets.
#
# Deliberately NOT summing instrumented sub-stage durations to build the
# book-level estimate — that would require accounting for every
# un-instrumented gap (metadata pass, report writing, etc.) and
# would silently drift low as new un-timed steps get added later. A whole-
# job historical lookup is simpler and self-correcting: whatever actually
# happens between "job started" and "job finished" is what gets recorded,
# instrumented or not.
# =====================================================================

TIMING_HISTORY_PATH = os.path.join(BASE_DIR, "job_timing_history.json")
TIMING_HISTORY_LOCK = threading.Lock()

# How many recent samples to keep per bucket. Capped (not unbounded) so the
# estimate tracks recent hardware/config reality (e.g. after a model switch)
# rather than being dragged down by runs from months ago.
TIMING_HISTORY_MAX_SAMPLES = 10

# Human-readable label per stage key, used for the "current task" ETA line.
# Falls back to a title-cased version of the raw key if a new stage key
# shows up here before this dict is updated.
STAGE_LABELS = {
    "outline": "Generating outline",
    "chapter_write": "Writing chapter",
    "cover_gen": "Generating cover art",
    "manuscript_build": "Building manuscript files",
    "chapter_edit": "Editorial pass",
    "chapter_score": "Scoring pass",
    "whole_book_synthesis": "Whole-book synthesis",
}

# [TIMING] lines arrive either bare (script run directly) or prefixed with
# pipeline_chain.py's "[label] " relay prefix (e.g. "[writer+cover] ") when
# run through the full pipeline — tolerate both.
_TIMING_LINE_RE = re.compile(r"^(?:\[[^\]]+\]\s*)?\[TIMING\]\s+(.*)$")


def _load_timing_history():
    if os.path.isfile(TIMING_HISTORY_PATH):
        try:
            with open(TIMING_HISTORY_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                data.setdefault("stage_durations", {})
                data.setdefault("job_durations", {})
                return data
        except Exception:
            pass
    return {"stage_durations": {}, "job_durations": {}}


def _save_timing_history(history):
    try:
        with open(TIMING_HISTORY_PATH, "w", encoding="utf-8") as f:
            json.dump(history, f, indent=2)
    except Exception:
        pass  # ETA is a nice-to-have; never let a write failure break a job


def _median(values):
    if not values:
        return None
    s = sorted(values)
    n = len(s)
    mid = n // 2
    if n % 2 == 1:
        return s[mid]
    return (s[mid - 1] + s[mid]) / 2.0


def _record_sample(bucket_dict, key, duration):
    samples = bucket_dict.setdefault(key, [])
    samples.append(duration)
    if len(samples) > TIMING_HISTORY_MAX_SAMPLES:
        del samples[: len(samples) - TIMING_HISTORY_MAX_SAMPLES]


def _record_stage_duration(stage, duration):
    if duration is None or duration < 0:
        return
    with TIMING_HISTORY_LOCK:
        history = _load_timing_history()
        _record_sample(history["stage_durations"], stage, duration)
        _save_timing_history(history)


def _record_job_duration(job_kind, duration):
    if not job_kind or duration is None or duration < 0:
        return
    with TIMING_HISTORY_LOCK:
        history = _load_timing_history()
        _record_sample(history["job_durations"], job_kind, duration)
        _save_timing_history(history)


def _stage_estimate(stage):
    with TIMING_HISTORY_LOCK:
        history = _load_timing_history()
    return _median(history["stage_durations"].get(stage, []))


def _job_kind_estimate(job_kind):
    if not job_kind:
        return None
    with TIMING_HISTORY_LOCK:
        history = _load_timing_history()
    return _median(history["job_durations"].get(job_kind, []))


def _normalize_job_kind(name):
    """Buckets a job's display name into a stable history key, ignoring the
    batch-count/genre/model decoration so 'Batch: 3x Full Pipeline (noir
    detective, model: llama3.1-16k)' lands in the same bucket as a plain
    'Full Pipeline (cozy mystery)' run. Order matters: check the more
    specific 'Full Pipeline' before the generic 'Book Generator' text, since
    neither label ever contains the other's exact phrase but this keeps the
    check unambiguous if wording changes later."""
    if "Full Pipeline" in name:
        return "Full Pipeline"
    if "Book Generator" in name:
        return "Book Generator"
    return None


def _display_task_label(stage, ch=None, total=None, label=None):
    base = STAGE_LABELS.get(stage) or (stage.replace("_", " ").title() if stage else "Working")
    if ch and total:
        base = f"{base} ({ch}/{total})"
        if label:
            base += f": {label.replace('_', ' ')}"
    return base


# =====================================================================
# Job management — each button click that runs a command becomes a Job.
# Output lines are pushed into a per-job queue; the /stream endpoint
# (Server-Sent Events) drains that queue live to the browser.
# =====================================================================

class Job:
    def __init__(self, job_id, name, cmds, cwd=None):
        """cmds: a list of argv lists, run one after another in the same job.
        Almost every job is a single command (len(cmds) == 1) — callers just
        wrap their one cmd in a list. Batch runs (see /api/run/batch) pass
        multiple commands so 2-5 books can queue as one job with one Stop
        button and one combined log, instead of the dashboard having to
        juggle several Job objects and figure out which one is "current"."""
        self.id = job_id
        self.name = name
        self.cmds = cmds
        self.cwd = cwd
        # Bounded, added 2026-08-13. This was an unbounded queue.Queue(), drained
        # only by an attached SSE consumer in /api/jobs/<id>/stream. When the
        # browser tab died mid-run on 2026-08-13 (79,237 lines overwhelmed the
        # DOM — see TERM_MAX_LINES in dashboard.html) nothing was draining it
        # anymore, so the rest of the run's output accumulated in this server's
        # memory on top of the browser already struggling. The two compounded at
        # exactly the moment the machine was under the most pressure.
        #
        # maxsize with drop-oldest-on-full means a detached or dead viewer costs
        # a bounded amount of RAM instead of a growing one. Log fidelity is
        # unaffected — _log_write() writes every line to disk independently of
        # this queue, so nothing is lost, only the live view skips ahead.
        self.q = queue.Queue(maxsize=self.QUEUE_MAXSIZE)
        self.dropped_lines = 0
        self.status = "starting"  # starting -> running -> finished / error / stopped
        self.returncode = None
        self.process = None  # whichever stage's subprocess is currently running
        self.started_at = time.time()

        # Section 6.14 — ETA/countdown state. job_kind buckets this job's
        # name for whole-book duration history (None if the name doesn't
        # match a known kind, e.g. "Launch Ollama" — those jobs just don't
        # get a book-level ETA). cmd_started_at is reset per command in a
        # batch (see _run) since ETA history is recorded per BOOK, not per
        # whole batch. current_stage tracks the most recent [TIMING] stage
        # seen, so a stray "end" for a stage that was never "start"ed
        # (shouldn't happen, but subprocess output is subprocess output)
        # doesn't clobber state for a different in-progress stage.
        self.job_kind = _normalize_job_kind(name)
        self.current_stage = None
        self.cmd_started_at = None

        # Persistent log file, independent of the browser's terminal panel —
        # survives page reloads, tab closes, and manual "Clear" clicks.
        safe_name = re.sub(r"[^a-zA-Z0-9._-]+", "_", name)[:60]
        stamp = time.strftime("%Y%m%d-%H%M%S")
        self.log_filename = f"{stamp}_{safe_name}_{job_id}.log"
        self.log_path = os.path.join(LOGS_DIR, self.log_filename)
        self._log_file = None

    def start(self):
        thread = threading.Thread(target=self._run, daemon=True)
        thread.start()

    def _run(self):
        try:
            self._log_file = open(self.log_path, "w", encoding="utf-8")
        except Exception:
            self._log_file = None
        try:
            self.status = "running"
            total = len(self.cmds)
            for idx, cmd in enumerate(self.cmds, start=1):
                if self._was_killed:
                    break
                if total > 1:
                    header = (
                        f"\n{'=' * 70}\n[batch] Starting book {idx}/{total}\n"
                        f"$ {' '.join(cmd)}\n{'=' * 70}\n"
                    )
                else:
                    header = f"$ {' '.join(cmd)}\n"
                self._qput(header)
                self._log_write(header)

                # Reset per-command (per-book) timing state. This is reset
                # HERE, not just once in __init__, because a batch job runs
                # several books back to back under one Job — book-level
                # duration history needs to measure each book individually,
                # not the whole batch's total wall-clock.
                self.cmd_started_at = time.time()
                self.current_stage = None

                self.process = subprocess.Popen(
                    cmd,
                    cwd=self.cwd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    bufsize=1,
                    creationflags=CREATE_NO_WINDOW,
                )
                for line in self.process.stdout:
                    self._qput(line)
                    self._log_write(line)
                    self._maybe_handle_timing(line)
                self.process.wait()
                self.returncode = self.process.returncode

                if self._was_killed:
                    break
                if self.returncode != 0:
                    if total > 1:
                        msg = (
                            f"[batch] Book {idx}/{total} failed (exit {self.returncode}) — "
                            f"stopping the batch before starting the next one.\n"
                        )
                        self._qput(msg)
                        self._log_write(msg)
                    break

                # Cooldown between books, added 2026-08-13 after a 5-book batch
                # took the machine down mid-run. The loop previously launched
                # book N+1 the instant book N's process exited, but process exit
                # is not resource release: the CUDA context teardown lags by a
                # few seconds, and Ollama's server (which outlives the Python
                # script that called it) holds its model for a ~5 minute
                # keep_alive unless told otherwise. Book N+1 therefore started
                # while book N's model was still resident, and on an 8GB card
                # Ollama's response to not fitting is to silently offload to
                # system RAM — so each book in the batch left the next one
                # deeper in host memory until the machine gave out.
                #
                # The pipeline scripts now unload their own models on exit, so
                # this is a backstop rather than the primary mechanism: pause,
                # then confirm VRAM actually came back before continuing.
                if idx < total and not self._was_killed:
                    self._batch_cooldown(idx, total)

                # This book/command finished cleanly — record how long it
                # took into the job-kind history so future ETAs improve.
                if self.job_kind:
                    _record_job_duration(self.job_kind, time.time() - self.cmd_started_at)

            if self._was_killed:
                self.status = "stopped"
            elif self.returncode == 0:
                self.status = "finished"
            else:
                self.status = "error"
        except FileNotFoundError as e:
            msg = f"[dashboard] Could not start process: {e}\n"
            self._qput(msg)
            self._log_write(msg)
            self.status = "error"
        except Exception as e:
            msg = f"[dashboard] Unexpected error: {e}\n"
            self._qput(msg)
            self._log_write(msg)
            self.status = "error"
        finally:
            self._log_write(f"\n[dashboard] Job finished with status: {self.status}\n")
            if self._log_file:
                self._log_file.close()
            self._qput(None)  # sentinel: stream is done

    # Seconds to pause between books in a batch before even checking VRAM.
    # Deliberately generous: this runs at most (N-1) times per batch, so on a
    # 3-book run it costs 60s total against books that take 20+ minutes each —
    # a rounding error next to the cost of the batch dying at book 4.
    BATCH_COOLDOWN_SECONDS = 30
    # Free VRAM the next book wants before it starts. Roughly what the writer
    # model needs to stay fully GPU-resident on an 8GB card; below this Ollama
    # quietly falls back to CPU rather than erroring.
    BATCH_COOLDOWN_TARGET_VRAM_MB = 6000
    BATCH_COOLDOWN_MAX_VRAM_WAIT_SECONDS = 90

    def _batch_cooldown(self, idx, total):
        """Pause and wait for resources to actually come back between books.
        See the call site in _run() for why process exit alone isn't enough."""
        msg = (f"\n[batch] Book {idx}/{total} finished. Cooling down "
               f"{self.BATCH_COOLDOWN_SECONDS}s before starting the next one...\n")
        self._qput(msg)
        self._log_write(msg)

        # Sleep in 1s slices so a Stop click during the cooldown is honoured
        # promptly instead of sitting out the whole pause.
        for _ in range(self.BATCH_COOLDOWN_SECONDS):
            if self._was_killed:
                return
            time.sleep(1)

        vram = _query_vram()
        if not vram.get("available"):
            return  # no nvidia-smi to gate on; the fixed pause is all we have
        waited = 0
        while (vram.get("free_mb", 0) < self.BATCH_COOLDOWN_TARGET_VRAM_MB
               and waited < self.BATCH_COOLDOWN_MAX_VRAM_WAIT_SECONDS
               and not self._was_killed):
            if waited == 0:
                m = (f"[batch] Only {vram['free_mb']}MB VRAM free "
                     f"(want {self.BATCH_COOLDOWN_TARGET_VRAM_MB}MB) — waiting for it to clear.\n")
                self._qput(m)
                self._log_write(m)
            time.sleep(2)
            waited += 2
            vram = _query_vram()

        if self._was_killed:
            return
        free = vram.get("free_mb", 0)
        if free >= self.BATCH_COOLDOWN_TARGET_VRAM_MB:
            m = f"[batch] {free}MB VRAM free — starting book {idx + 1}/{total}.\n"
        else:
            m = (f"[batch] WARNING: still only {free}MB VRAM free after waiting "
                 f"{self.BATCH_COOLDOWN_MAX_VRAM_WAIT_SECONDS}s. Starting book {idx + 1}/{total} "
                 f"anyway, but it may run slowly — check Maintenance & Diagnostics for a leftover "
                 f"A1111 or Ollama process holding VRAM.\n")
        self._qput(m)
        self._log_write(m)

    # Bounded live-view buffer — see the comment on self.q in __init__. Sized to
    # comfortably cover a normal viewing gap (page reload, brief tab switch)
    # while capping worst-case memory if nothing ever reattaches.
    QUEUE_MAXSIZE = 5000

    def _qput(self, item):
        """Put an item on the live-view queue, discarding the OLDEST item if the
        queue is full rather than blocking the reader thread.

        Blocking would be worse than dropping: this is called from the thread
        reading the subprocess's stdout, so a stalled queue would stall that
        read, which would fill the OS pipe buffer, which would block the
        pipeline script itself. A dead browser tab must never be able to halt a
        running book. Every line still reaches the on-disk log regardless."""
        try:
            self.q.put_nowait(item)
        except queue.Full:
            try:
                self.q.get_nowait()      # discard oldest
                self.dropped_lines += 1
            except queue.Empty:
                pass
            try:
                self.q.put_nowait(item)
            except queue.Full:
                pass  # reader is gone entirely; the disk log is still authoritative

    def _log_write(self, text):
        if self._log_file:
            try:
                self._log_file.write(text if text.endswith("\n") else text + "\n")
                self._log_file.flush()
            except Exception:
                pass

    def _maybe_handle_timing(self, line):
        """Parses a raw stdout line for a [TIMING] marker (see the pipeline
        scripts' _emit_timing() helpers). The raw line is ALSO still sent to
        the terminal/log as usual (see the caller in _run) — this just adds
        a second, structured "eta" event on top for the frontend's countdown
        display. Never lets a malformed marker line raise; ETA is a nice-to-
        have, not something that should ever break a running job."""
        try:
            m = _TIMING_LINE_RE.match(line.strip())
            if not m:
                return
            fields = {}
            for tok in m.group(1).split():
                if "=" in tok:
                    k, v = tok.split("=", 1)
                    fields[k] = v
            event = fields.get("event")
            stage = fields.get("stage")
            if not event or not stage:
                return

            if event == "start":
                self.current_stage = stage
            elif event == "end":
                elapsed = fields.get("elapsed")
                if elapsed is not None:
                    try:
                        _record_stage_duration(stage, float(elapsed))
                    except ValueError:
                        pass
                if self.current_stage == stage:
                    self.current_stage = None

            self._qput(self._eta_payload(stage, fields))
        except Exception:
            pass

    def _eta_payload(self, stage, fields):
        ch = fields.get("ch")
        total = fields.get("total")
        label = fields.get("label")
        task_label = _display_task_label(stage, ch=ch, total=total, label=label)
        task_remaining = _stage_estimate(stage)

        book_remaining = None
        book_estimate = _job_kind_estimate(self.job_kind)
        if book_estimate is not None and self.cmd_started_at is not None:
            elapsed_in_book = time.time() - self.cmd_started_at
            book_remaining = max(0.0, book_estimate - elapsed_in_book)

        return {
            "__eta__": True,
            "book_label": self.name,
            "book_remaining": book_remaining,
            "task_label": task_label,
            "task_remaining": task_remaining,
        }

    _was_killed = False

    def stop(self):
        self._was_killed = True
        if self.process and self.process.poll() is None:
            if sys.platform == "win32":
                # process.terminate() only kills this one process. That's fine
                # for a plain "Run" (local-book-generator.py IS the job), but
                # the full-pipeline job wraps pipeline_chain.py, which spawns
                # its own writer/editor/scorer subprocess per stage — those
                # children would otherwise be orphaned and keep running (and
                # keep holding VRAM) after Stop is clicked. taskkill /T kills
                # the whole process tree instead, so Stop works the same way
                # regardless of which job type is running.
                try:
                    subprocess.run(
                        ["taskkill", "/F", "/T", "/PID", str(self.process.pid)],
                        creationflags=CREATE_NO_WINDOW,
                        capture_output=True,
                    )
                    return
                except Exception:
                    pass  # fall through to the plain terminate() below
            self.process.terminate()


JOBS = {}
JOBS_LOCK = threading.Lock()


def start_job(name, cmds, cwd=None):
    """cmds: a list of argv lists (see Job.__init__). Callers running a single
    command pass a one-item list, e.g. start_job(name, [cmd], cwd=...)."""
    job_id = uuid.uuid4().hex[:12]
    job = Job(job_id, name, cmds, cwd=cwd)
    with JOBS_LOCK:
        JOBS[job_id] = job
    job.start()
    return job_id


# =====================================================================
# Routes — static page
# =====================================================================

@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "dashboard.html")


# =====================================================================
# Routes — config
# =====================================================================

@app.route("/api/config", methods=["GET"])
def get_config():
    return jsonify(load_config())


@app.route("/api/config", methods=["POST"])
def set_config():
    incoming = request.get_json(force=True, silent=True) or {}
    cfg = load_config()
    for key in DEFAULT_CONFIG:
        if key in incoming and incoming[key] is not None:
            cfg[key] = incoming[key]
    save_config(cfg)
    return jsonify(cfg)


@app.route("/api/book-structure", methods=["GET"])
def get_book_structure():
    """Presets, the saved selection, and the custom values, for the panel."""
    cfg = load_config()
    return jsonify({
        "presets": {k: dict(v) for k, v in BOOK_STRUCTURE_PRESETS.items()},
        # Explicit order, because jsonify sorts object keys alphabetically and
        # that put Full-length before Novelette in the button row. Shortest to
        # longest is the order these actually make sense in.
        "preset_order": list(BOOK_STRUCTURE_PRESETS),
        "fields": list(STRUCTURE_FIELDS),
        "selected": cfg.get("book_structure_preset", "random"),
        "custom": cfg.get("book_structure_custom",
                          {f: BOOK_STRUCTURE_PRESETS["novella"][f] for f in STRUCTURE_FIELDS}),
    })


@app.route("/api/book-structure", methods=["POST"])
def set_book_structure():
    """Persist the panel's selection. Custom values are validated before saving —
    refusing an impossible combination here means it can't sit in the config
    quietly waiting to break the next run."""
    payload = request.get_json(force=True, silent=True) or {}
    preset = payload.get("preset", "random")
    if preset not in list(BOOK_STRUCTURE_PRESETS) + ["random", "custom"]:
        return jsonify({"ok": False, "detail": f"Unknown preset '{preset}'."}), 400

    cfg = load_config()
    cfg["book_structure_preset"] = preset

    if preset == "custom":
        custom = payload.get("custom") or {}
        parsed = {}
        for field in STRUCTURE_FIELDS:
            try:
                parsed[field] = int(custom[field])
            except (KeyError, TypeError, ValueError):
                return jsonify({"ok": False,
                                "detail": f"'{field.replace('_', ' ')}' must be a whole number."}), 400
        error = validate_book_structure(parsed)
        if error:
            return jsonify({"ok": False, "detail": error}), 400
        cfg["book_structure_custom"] = parsed

    save_config(cfg)
    return jsonify({"ok": True, "preset": preset,
                    "custom": cfg.get("book_structure_custom")})


@app.route("/api/writer-models", methods=["GET"])
def get_writer_models():
    return jsonify(WRITER_MODEL_OPTIONS)


@app.route("/api/genres", methods=["GET"])
def get_genres():
    return jsonify(GENRE_POOL)


# =====================================================================
# Routes — status checks (quick, non-streaming)
# =====================================================================

@app.route("/api/status/ollama", methods=["GET"])
def status_ollama():
    cfg = load_config()
    try:
        import urllib.request
        with urllib.request.urlopen(cfg["ollama_url"], timeout=2) as resp:
            resp.read()
        return jsonify({"running": True})
    except Exception as e:
        return jsonify({"running": False, "detail": str(e)})


@app.route("/api/status/a1111", methods=["GET"])
def status_a1111():
    cfg = load_config()
    try:
        import urllib.request
        url = cfg["a1111_url"].rstrip("/") + "/sdapi/v1/sd-models"
        with urllib.request.urlopen(url, timeout=3) as resp:
            data = resp.read()
        return jsonify({"running": True, "model_count": len(json.loads(data))})
    except Exception as e:
        return jsonify({"running": False, "detail": str(e)})


def _query_vram():
    """Shared by /api/status/vram (6.6) and /api/system-test (6.9) so there's
    one nvidia-smi call site instead of two copies drifting apart. Same query
    already used by get_free_vram_mb() in local-book-generator.py /
    editorial_agent.py / scoring_agent.py, plus the total so callers can show
    'X / Y MB free' instead of a bare number."""
    try:
        result = subprocess.run(
            ["nvidia-smi", "--query-gpu=memory.free,memory.total", "--format=csv,noheader,nounits"],
            capture_output=True, text=True, timeout=10,
            creationflags=CREATE_NO_WINDOW,
        )
        if result.returncode == 0 and result.stdout.strip():
            parts = result.stdout.strip().splitlines()[0].split(",")
            free_mb = int(parts[0].strip())
            total_mb = int(parts[1].strip())
            return {"available": True, "free_mb": free_mb, "total_mb": total_mb}
        return {"available": False, "detail": (result.stderr or "nvidia-smi returned no output").strip()}
    except (subprocess.SubprocessError, FileNotFoundError, ValueError, OSError) as e:
        return {"available": False, "detail": str(e)}


@app.route("/api/status/vram", methods=["GET"])
def status_vram():
    """Passive, always-visible free-VRAM readout (6.6)."""
    return jsonify(_query_vram())


# =====================================================================
# Routes — conflicting-process check/kill for AUTOMATIC1111 (6.7)
#
# Reimplements the manual diagnostic pattern (netstat -ano | findstr
# "7860 7861 7862" -> taskkill /PID <pid> /F) in Python instead of shelling
# out to findstr, so the port numbers and LISTENING-state filter are parsed
# reliably rather than string-matched — same ports, same taskkill call,
# just driven from a button with a confirm step instead of a terminal.
# Windows-only, like Job.stop()'s taskkill /T usage above.
# =====================================================================

A1111_CONFLICT_PORTS = ("7860", "7861", "7862")
_NETSTAT_LISTEN_RE = re.compile(r"TCP6?\s+\S*:(\d+)\s+\S+\s+LISTENING\s+(\d+)", re.IGNORECASE)


def _find_a1111_conflicts():
    """Returns a list of {"port", "pid"} dicts for anything LISTENING on
    A1111's ports, deduped by (port, pid). Returns None (not an empty list)
    on non-Windows, so callers can distinguish 'checked, found nothing' from
    'can't check on this OS' rather than reporting a false all-clear."""
    if sys.platform != "win32":
        return None
    try:
        result = subprocess.run(
            ["netstat", "-ano"], capture_output=True, text=True, timeout=10,
            creationflags=CREATE_NO_WINDOW,
        )
    except (subprocess.SubprocessError, OSError):
        return []
    conflicts = []
    seen = set()
    for line in result.stdout.splitlines():
        m = _NETSTAT_LISTEN_RE.search(line)
        if not m:
            continue
        port, pid = m.group(1), m.group(2)
        if port in A1111_CONFLICT_PORTS and (port, pid) not in seen:
            seen.add((port, pid))
            conflicts.append({"port": port, "pid": pid})
    return conflicts


@app.route("/api/a1111/conflicts", methods=["GET"])
def a1111_conflicts():
    conflicts = _find_a1111_conflicts()
    if conflicts is None:
        return jsonify({
            "supported": False, "conflicts": [],
            "detail": "Conflict check is Windows-only (uses netstat/taskkill).",
        })
    return jsonify({"supported": True, "conflicts": conflicts})


@app.route("/api/a1111/kill", methods=["POST"])
def a1111_kill():
    """Re-checks (rather than trusting a stale list from the browser) and
    force-kills every PID currently listening on A1111's ports. A PID
    listening on more than one of the 3 ports is only taskkill'd once, but
    every matching port is still reported back in the result."""
    conflicts = _find_a1111_conflicts()
    if conflicts is None:
        return jsonify({
            "supported": False, "killed": [],
            "detail": "Conflict check is Windows-only (uses netstat/taskkill).",
        }), 400

    killed = []
    pid_results = {}
    for c in conflicts:
        pid = c["pid"]
        if pid not in pid_results:
            try:
                result = subprocess.run(
                    ["taskkill", "/PID", pid, "/F"],
                    capture_output=True, text=True, timeout=10,
                    creationflags=CREATE_NO_WINDOW,
                )
                pid_results[pid] = {
                    "ok": result.returncode == 0,
                    "detail": (result.stdout or result.stderr or "").strip(),
                }
            except (subprocess.SubprocessError, OSError) as e:
                pid_results[pid] = {"ok": False, "detail": str(e)}
        killed.append({"port": c["port"], "pid": pid, **pid_results[pid]})
    return jsonify({"supported": True, "killed": killed})


def _taskkill_pid(pid: str) -> dict:
    """Shared by /api/a1111/kill (implicitly, via its own inline copy above)
    and /api/tasks/kill (6.8) below. Kept as its own function here rather
    than also refactoring a1111_kill's copy — that route already shipped and
    is working, not worth touching for a pure dedupe."""
    try:
        result = subprocess.run(
            ["taskkill", "/PID", pid, "/F"],
            capture_output=True, text=True, timeout=10,
            creationflags=CREATE_NO_WINDOW,
        )
        return {"ok": result.returncode == 0, "detail": (result.stdout or result.stderr or "").strip()}
    except (subprocess.SubprocessError, OSError) as e:
        return {"ok": False, "detail": str(e)}


# =====================================================================
# Routes — kill-task dropdown (6.8): broader than 6.7's A1111-only button.
# Lists live candidates across every process type the pipeline can leave
# running (Ollama, AUTOMATIC1111, and the standalone pipeline scripts
# themselves if one gets orphaned outside the dashboard's own Job tracking —
# e.g. started by hand in a terminal, or a crash that detached the child).
# Windows-only, like 6.7. dashboard_server.py itself is deliberately never a
# candidate — killing the process serving this very page out from under
# yourself is a footgun, not a feature.
# =====================================================================

# Basenames matched against each python.exe's full command line (via
# PowerShell, since plain tasklist doesn't expose command lines). Deliberately
# excludes dashboard_server.py — see note above. Kept in sync by hand with
# the actual script filenames; update here if one is ever renamed.
PIPELINE_SCRIPT_BASENAMES = (
    "local-book-generator.py",
    "editorial_agent.py",
    "scoring_agent.py",
    "pipeline_chain.py",
    "repolish_agent.py",
)


# Ollama on Windows is TWO processes, not one, and the distinction is the whole
# reason "ending task" on the rogue process didn't stick on 2026-08-13:
#
#   ollama app.exe — the system-tray supervisor. Owns the tray icon and, more
#                    importantly, RESTARTS the server whenever it sees it exit.
#   ollama.exe     — the actual server (`ollama serve`) that loads models and
#                    holds the VRAM/RAM.
#
# Killing ollama.exe alone gets it immediately respawned by the supervisor,
# which is exactly the "rogue ollama service kept popping up in Task Manager
# despite ending task" behaviour reported — and it reappears without a tray
# icon because the respawned server is a bare child process, which is why the
# tray looked empty while Task Manager showed it running.
#
# The supervisor must be killed FIRST, then the server. _kill_ollama_stack()
# below does that in the right order.
OLLAMA_SUPERVISOR_IMAGE = "ollama app.exe"
OLLAMA_SERVER_IMAGE = "ollama.exe"


def _tasklist_by_image(image_name: str) -> list:
    """Windows-only: PIDs for one image name via tasklist. [] on any failure."""
    try:
        result = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {image_name}", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=10,
            creationflags=CREATE_NO_WINDOW,
        )
    except (subprocess.SubprocessError, OSError):
        return []
    found = []
    for line in result.stdout.splitlines():
        # CSV columns: "Image Name","PID","Session Name","Session#","Mem Usage"
        parts = [p.strip('"') for p in line.strip().split('","')]
        if len(parts) >= 2 and parts[0].lower() == image_name.lower() and parts[1].isdigit():
            found.append({"pid": parts[1], "image": parts[0]})
    return found


def _find_ollama_processes():
    """Windows-only: every Ollama process — supervisor AND server. Returns None
    off Windows (can't check), [] if Windows but nothing found/tasklist failed.

    Previously only looked for ollama.exe, which made the supervisor invisible
    in the kill dropdown and any kill of the server pointless (see the comment
    block above). Supervisor entries are listed first so the UI's ordering
    matches the order they need to be killed in."""
    if sys.platform != "win32":
        return None
    return _tasklist_by_image(OLLAMA_SUPERVISOR_IMAGE) + _tasklist_by_image(OLLAMA_SERVER_IMAGE)


def _unload_all_ollama_models(ollama_url: str) -> list:
    """Ask Ollama to drop every currently-loaded model from memory (keep_alive:0).

    This is the *graceful* answer to "ollama is eating 50% CPU and 50% RAM" and
    should be preferred over killing the process: it frees the memory without
    tearing down the server, so the next pipeline stage doesn't have to wait for
    a restart. Returns a list of human-readable result lines."""
    # urllib rather than requests, matching the other outbound calls in this
    # file — the dashboard deliberately has no dependency beyond flask.
    import urllib.request

    base = ollama_url.rstrip("/")
    lines = []
    try:
        with urllib.request.urlopen(f"{base}/api/ps", timeout=10) as resp:
            loaded = [m.get("name") or m.get("model")
                      for m in (json.loads(resp.read()).get("models") or [])]
        loaded = [m for m in loaded if m]
    except Exception as e:
        return [f"Could not list loaded models: {e}"]
    if not loaded:
        return ["No models are currently loaded — nothing to unload."]
    for name in loaded:
        try:
            req = urllib.request.Request(
                f"{base}/api/generate",
                data=json.dumps({"model": name, "keep_alive": 0}).encode("utf-8"),
                headers={"Content-Type": "application/json"},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=15):
                pass
            lines.append(f"Unloaded '{name}' — its VRAM/RAM is released.")
        except Exception as e:
            lines.append(f"Failed to unload '{name}': {e}")
    return lines


def _kill_ollama_stack() -> dict:
    """Kill the tray supervisor first, then the server, so the server stays dead.

    Killing them in the other order (or killing only the server) gets the server
    instantly respawned by the supervisor — the 2026-08-13 symptom. After the
    supervisor is gone, the server is killed with /T so any model subprocess it
    spawned goes with it."""
    killed, failed = [], []
    for image in (OLLAMA_SUPERVISOR_IMAGE, OLLAMA_SERVER_IMAGE):
        for proc in _tasklist_by_image(image):
            try:
                result = subprocess.run(
                    ["taskkill", "/PID", proc["pid"], "/F", "/T"],
                    capture_output=True, text=True, timeout=15,
                    creationflags=CREATE_NO_WINDOW,
                )
                if result.returncode == 0:
                    killed.append(f"{image} (PID {proc['pid']})")
                else:
                    failed.append(f"{image} (PID {proc['pid']}): {result.stderr.strip()}")
            except (subprocess.SubprocessError, OSError) as e:
                failed.append(f"{image} (PID {proc['pid']}): {e}")
        # Give the supervisor a moment to actually exit before killing the
        # server — otherwise a still-alive supervisor can win the race and
        # respawn the server we're about to kill.
        time.sleep(1.5)
    return {"ok": not failed, "killed": killed, "failed": failed}


def _find_pipeline_python_processes():
    """Windows-only: every python.exe/pythonw.exe whose command line mentions
    one of PIPELINE_SCRIPT_BASENAMES, via PowerShell (tasklist alone can't see
    command lines). Best-effort: any failure (PowerShell missing/blocked by
    policy/timeout) degrades to an empty list rather than breaking the whole
    candidates endpoint — Ollama/A1111 detection still work independently."""
    if sys.platform != "win32":
        return None
    ps_script = (
        "Get-CimInstance Win32_Process "
        "-Filter \"Name='python.exe' OR Name='pythonw.exe'\" "
        "| Select-Object ProcessId,CommandLine | ConvertTo-Json -Compress"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            capture_output=True, text=True, timeout=15,
            creationflags=CREATE_NO_WINDOW,
        )
        if result.returncode != 0 or not result.stdout.strip():
            return []
        data = json.loads(result.stdout)
    except (subprocess.SubprocessError, OSError, json.JSONDecodeError, ValueError):
        return []
    if isinstance(data, dict):  # ConvertTo-Json returns a bare object, not an array, for 1 match
        data = [data]
    found = []
    for entry in data:
        cmdline = (entry.get("CommandLine") or "")
        pid = entry.get("ProcessId")
        if pid is None:
            continue
        for basename in PIPELINE_SCRIPT_BASENAMES:
            if basename in cmdline:
                found.append({"pid": str(pid), "script": basename})
                break
    return found


@app.route("/api/tasks/candidates", methods=["GET"])
def tasks_candidates():
    candidates = []

    ollama_procs = _find_ollama_processes()
    supported = ollama_procs is not None
    if ollama_procs:
        for p in ollama_procs:
            # Label names the actual image so the supervisor and the server are
            # distinguishable in the dropdown. Killing either one now takes the
            # whole stack down in the right order (see _kill_ollama_stack), so
            # the selection can't produce the respawn loop hit on 2026-08-13.
            image = p.get("image", "ollama.exe")
            role = "tray supervisor" if image.lower() == OLLAMA_SUPERVISOR_IMAGE else "server"
            candidates.append({"id": f"ollama-{p['pid']}", "pid": p["pid"], "kind": "ollama",
                                "label": f"Ollama {role} — {image} (PID {p['pid']}) "
                                         f"[kills both]"})

    a1111_conflicts = _find_a1111_conflicts()
    if a1111_conflicts:
        seen_pids = set()
        for c in a1111_conflicts:
            if c["pid"] in seen_pids:
                continue
            seen_pids.add(c["pid"])
            ports = sorted({x["port"] for x in a1111_conflicts if x["pid"] == c["pid"]})
            candidates.append({"id": f"a1111-{c['pid']}", "pid": c["pid"], "kind": "a1111",
                                "label": f"AUTOMATIC1111 (PID {c['pid']}, port {'/'.join(ports)})"})

    pipeline_procs = _find_pipeline_python_processes()
    if pipeline_procs:
        for p in pipeline_procs:
            candidates.append({"id": f"pipeline-{p['pid']}", "pid": p["pid"], "kind": "pipeline",
                                "label": f"{p['script']} (PID {p['pid']})"})

    if not supported:
        return jsonify({
            "supported": False, "candidates": [],
            "detail": "Task list is Windows-only (uses tasklist/PowerShell/taskkill).",
        })
    return jsonify({"supported": True, "candidates": candidates})


@app.route("/api/tasks/kill", methods=["POST"])
def tasks_kill():
    if sys.platform != "win32":
        return jsonify({"ok": False, "detail": "Task kill is Windows-only (uses taskkill)."}), 400
    payload = request.get_json(force=True, silent=True) or {}
    pid = str(payload.get("pid", "")).strip()
    if not pid.isdigit():
        return jsonify({"ok": False, "detail": f"Invalid PID '{pid}'."}), 400

    # Killing an Ollama process by bare PID doesn't stick: the tray supervisor
    # ("ollama app.exe") respawns the server the moment it exits. So for any
    # Ollama candidate, take down the whole stack in the correct order instead
    # of the single PID the user happened to select. See _kill_ollama_stack().
    if str(payload.get("kind", "")).strip().lower() == "ollama":
        result = _kill_ollama_stack()
        detail = "Killed: " + (", ".join(result["killed"]) or "nothing")
        if result["failed"]:
            detail += " | Failed: " + ", ".join(result["failed"])
        return jsonify({"pid": pid, "ok": result["ok"], "detail": detail})

    result = _taskkill_pid(pid)
    return jsonify({"pid": pid, **result})


@app.route("/api/ollama/unload", methods=["POST"])
def ollama_unload():
    """Graceful alternative to killing Ollama: drop every loaded model from
    memory but leave the server running. Added 2026-08-13 — this is what you
    actually want when ollama.exe is sitting on 50% CPU / 50% RAM between runs,
    since it reclaims the memory without the restart (and without the tray
    supervisor fighting you)."""
    cfg = load_config()
    lines = _unload_all_ollama_models(cfg.get("ollama_url", "http://localhost:11434"))
    return jsonify({"ok": True, "lines": lines})


# =====================================================================
# Route — "Launch System Test" hardware/config diagnostics (6.9). Read-only
# status report, not a fix-it action — doesn't launch, stop, or change
# anything, and doesn't leave any process running. Returns a flat list of
# report lines; the dashboard prints them straight into the terminal panel
# like a job's output, without actually going through the Job/subprocess
# streaming machinery (every check here is quick enough to just do inline).
# =====================================================================

@app.route("/api/system-test", methods=["GET"])
def system_test():
    cfg = load_config()
    lines = []

    def add(line):
        lines.append(line)

    add("=" * 70)
    add("SYSTEM TEST — read-only diagnostics, nothing was started, stopped, or changed")
    add("=" * 70)

    # --- Paths & config ---
    add("")
    add("-- Paths & config --")
    py_exe = cfg.get("python_exe", "")
    add(f"Python executable: {py_exe or '(not set)'}"
        + ("" if py_exe and (os.path.isfile(py_exe) or shutil.which(py_exe)) else "  [!] not found"))
    for key, label in (
        ("script_path", "local-book-generator.py"),
        ("pipeline_script_path", "pipeline_chain.py"),
        ("a1111_bat_path", "webui-user.bat"),
    ):
        path = cfg.get(key, "")
        ok = bool(path) and os.path.isfile(path)
        add(f"{label}: {path or '(not set)'}" + ("" if ok else "  [!] not found"))
    for extra_name in ("editorial_agent.py", "scoring_agent.py", "repolish_agent.py"):
        extra_path = os.path.join(BASE_DIR, extra_name)
        ok = os.path.isfile(extra_path)
        add(f"{extra_name}: {extra_path}" + ("" if ok else "  [!] not found"))

    # --- Ollama ---
    add("")
    add("-- Ollama --")
    try:
        import urllib.request
        with urllib.request.urlopen(cfg["ollama_url"], timeout=3) as resp:
            resp.read()
        add(f"Reachable at {cfg['ollama_url']}: yes")
        try:
            with urllib.request.urlopen(cfg["ollama_url"].rstrip("/") + "/api/tags", timeout=5) as resp:
                tags_data = json.loads(resp.read())
            pulled = {m.get("name", "") for m in tags_data.get("models", [])}
            pulled_bare = {n.split(":")[0] for n in pulled if n}
            for model in REQUIRED_OLLAMA_MODELS:
                have = model in pulled or model in pulled_bare
                add(f"  Model '{model}': {'present' if have else 'MISSING'}")
        except Exception as e:
            add(f"  [!] Could not list models: {e}")
    except Exception as e:
        add(f"Reachable at {cfg['ollama_url']}: NO ({e})")

    # --- AUTOMATIC1111 ---
    add("")
    add("-- AUTOMATIC1111 --")
    try:
        import urllib.request
        url = cfg["a1111_url"].rstrip("/") + "/sdapi/v1/options"
        with urllib.request.urlopen(url, timeout=3) as resp:
            opts = json.loads(resp.read())
        add(f"Reachable at {cfg['a1111_url']}: yes")
        add(f"  Loaded checkpoint: {opts.get('sd_model_checkpoint', '(unknown)')}")
    except Exception as e:
        add(f"Reachable at {cfg['a1111_url']}: NO ({e})")
    conflicts = _find_a1111_conflicts()
    if conflicts:
        add("  [!] Something is listening on A1111's ports right now: "
            + ", ".join(f"PID {c['pid']} (port {c['port']})" for c in conflicts))
        add("      Use the 'Kill conflicting process' button in Dependencies if this is unexpected.")

    # --- VRAM ---
    add("")
    add("-- GPU VRAM --")
    vram = _query_vram()
    if vram["available"]:
        add(f"Free: {vram['free_mb']:,} MB / {vram['total_mb']:,} MB total")
    else:
        add(f"Unavailable: {vram.get('detail', '(no detail)')}")

    # --- Disk space ---
    add("")
    add("-- Disk space --")
    try:
        usage = shutil.disk_usage(BASE_DIR)
        add(f"Free on this drive: {usage.free / (1024 ** 3):.1f} GB / {usage.total / (1024 ** 3):.1f} GB total")
    except OSError as e:
        add(f"[!] Could not check disk space: {e}")

    # --- Book library ---
    add("")
    add("-- Book library (output_books/) --")
    output_root = os.path.join(BASE_DIR, "output_books")
    if os.path.isdir(output_root):
        book_count = 0
        total_bytes = 0
        for name in os.listdir(output_root):
            book_path = os.path.join(output_root, name)
            if os.path.isdir(book_path) and os.path.isfile(os.path.join(book_path, "outline.json")):
                book_count += 1
                for dirpath, _, filenames in os.walk(book_path):
                    for fn in filenames:
                        try:
                            total_bytes += os.path.getsize(os.path.join(dirpath, fn))
                        except OSError:
                            pass
        add(f"{book_count} book folder(s), {total_bytes / (1024 ** 2):.1f} MB total")
    else:
        add("(no output_books/ folder yet — nothing generated so far)")

    add("")
    add("=" * 70)
    add("SYSTEM TEST COMPLETE")
    return jsonify({"lines": lines})


# =====================================================================
# Routes — title history (6.11): local-book-generator.py's dedupe-title
# guard (see title_history.json / dedupe_title() there) grows forever by
# design — this just gives you a way to see/clear it from the dashboard
# instead of deleting the file by hand.
# =====================================================================

TITLE_HISTORY_PATH = os.path.join(BASE_DIR, "title_history.json")


def _load_title_history() -> dict:
    # Mirrors the type check added to local-book-generator.py 2026-08-14: a file
    # that parses as valid JSON of the wrong type (e.g. "[]") must not be handed
    # back as a dict, or the summary/clear routes report nonsense.
    if os.path.isfile(TITLE_HISTORY_PATH):
        try:
            with open(TITLE_HISTORY_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except (json.JSONDecodeError, OSError):
            pass
    return {}


@app.route("/api/title-history/summary", methods=["GET"])
def title_history_summary():
    history = _load_title_history()
    return jsonify({"count": len(history)})


@app.route("/api/title-history/clear", methods=["POST"])
def title_history_clear():
    payload = request.get_json(force=True, silent=True) or {}
    backup = bool(payload.get("backup", False))
    history = _load_title_history()
    count = len(history)

    backup_path = None
    if backup and count > 0:
        stamp = time.strftime("%Y%m%d-%H%M%S")
        backup_path = os.path.join(BASE_DIR, f"title_history_backup_{stamp}.zip")
        try:
            with zipfile.ZipFile(backup_path, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.write(TITLE_HISTORY_PATH, arcname="title_history.json")
        except OSError as e:
            return jsonify({"ok": False, "detail": f"Backup failed, nothing was cleared: {e}"}), 500

    try:
        with open(TITLE_HISTORY_PATH, "w", encoding="utf-8") as f:
            json.dump({}, f, indent=2)
    except OSError as e:
        return jsonify({"ok": False, "detail": f"Could not clear title history: {e}"}), 500

    return jsonify({"ok": True, "cleared_count": count, "backup_path": backup_path})


# =====================================================================
# Routes — jobs (launching things, streaming output)
# =====================================================================

@app.route("/api/jobs", methods=["GET"])
def list_jobs():
    with JOBS_LOCK:
        return jsonify([
            {
                "id": j.id,
                "name": j.name,
                "status": j.status,
                "started_at": j.started_at,
            }
            for j in JOBS.values()
        ])


@app.route("/api/jobs/<job_id>/status", methods=["GET"])
def job_status(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
    if not job:
        return jsonify({"error": "unknown job"}), 404
    return jsonify({"id": job.id, "name": job.name, "status": job.status, "returncode": job.returncode})


# =====================================================================
# Routes — persisted logs (survive Clear/reload/tab close)
# =====================================================================

_SAFE_LOG_NAME = re.compile(r"^[a-zA-Z0-9._-]+\.log$")


@app.route("/api/logs", methods=["GET"])
def list_logs():
    entries = []
    for fname in os.listdir(LOGS_DIR):
        if not _SAFE_LOG_NAME.match(fname):
            continue
        fpath = os.path.join(LOGS_DIR, fname)
        try:
            stat = os.stat(fpath)
            entries.append({
                "filename": fname,
                "size_bytes": stat.st_size,
                "modified_at": stat.st_mtime,
            })
        except OSError:
            continue
    entries.sort(key=lambda e: e["modified_at"], reverse=True)
    return jsonify(entries)


@app.route("/api/logs/<path:filename>", methods=["GET"])
def get_log(filename):
    if not _SAFE_LOG_NAME.match(filename):
        return jsonify({"error": "invalid filename"}), 400
    fpath = os.path.join(LOGS_DIR, filename)
    if not os.path.isfile(fpath):
        return jsonify({"error": "log not found"}), 404
    with open(fpath, "r", encoding="utf-8", errors="replace") as f:
        content = f.read()
    return jsonify({"filename": filename, "content": content})


@app.route("/api/jobs/<job_id>/stop", methods=["POST"])
def job_stop(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
    if not job:
        return jsonify({"error": "unknown job"}), 404
    job.stop()
    return jsonify({"ok": True})


@app.route("/api/jobs/<job_id>/stream")
def job_stream(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
    if not job:
        return jsonify({"error": "unknown job"}), 404

    def generate():
        reported_drops = 0
        while True:
            item = job.q.get()
            # Tell the viewer when the live buffer skipped ahead, so a gap in the
            # terminal never reads as the job having gone quiet. The on-disk log
            # is always complete — see Job._qput().
            if job.dropped_lines > reported_drops:
                skipped = job.dropped_lines - reported_drops
                reported_drops = job.dropped_lines
                yield (f"data: [dashboard] ...{skipped} line(s) skipped in this live view to keep "
                       f"the page responsive. The full log is on disk.\n\n")
            if item is None:
                yield f"event: done\ndata: {job.status}\n\n"
                break
            if isinstance(item, dict):
                # Section 6.14 — structured ETA update (see Job._eta_payload),
                # distinct from the plain log-line events below.
                yield f"event: eta\ndata: {json.dumps(item)}\n\n"
                continue
            # SSE: escape newlines already handled since we send one line per event
            yield f"data: {item.rstrip(chr(10))}\n\n"

    return Response(generate(), mimetype="text/event-stream", headers={
        "Cache-Control": "no-cache",
        "X-Accel-Buffering": "no",
    })


@app.route("/api/launch/ollama", methods=["POST"])
def launch_ollama():
    cfg = load_config()
    job_id = start_job("Launch Ollama", [[cfg["ollama_exe"], "serve"]], cwd=BASE_DIR)
    return jsonify({"job_id": job_id})


@app.route("/api/launch/a1111", methods=["POST"])
def launch_a1111():
    cfg = load_config()
    bat_path = cfg.get("a1111_bat_path", "").strip()
    if not bat_path or not os.path.isfile(bat_path):
        return jsonify({"error": f"webui-user.bat path not found: '{bat_path}'. Set it in Settings first."}), 400
    cwd = os.path.dirname(bat_path)
    # Run via cmd /c so output is captured instead of opening a separate window.
    job_id = start_job("Launch AUTOMATIC1111", [["cmd", "/c", bat_path]], cwd=cwd)
    return jsonify({"job_id": job_id})


def _structure_for_request(cfg, payload):
    """Resolve the Book Structure for one book. Returns (structure, error_string).

    Preset comes from the request if present, otherwise from saved config, so the
    panel's persisted choice applies even to a Run click that predates it.
    Validated here so a bad combination is refused at the click rather than
    failing at the outline stage several seconds into a run."""
    preset = (payload.get("structure_preset")
              or cfg.get("book_structure_preset", "random"))
    custom = payload.get("structure_custom") or cfg.get("book_structure_custom")
    try:
        structure = resolve_book_structure(preset, custom)
    except (TypeError, ValueError, KeyError):
        return None, "Book Structure values must all be whole numbers."
    missing = [f for f in STRUCTURE_FIELDS if f not in structure]
    if missing:
        return None, f"Book Structure is missing: {', '.join(missing)}."
    error = validate_book_structure(structure)
    return (None, error) if error else (structure, None)


@app.route("/api/run/bookgen", methods=["POST"])
def run_bookgen():
    cfg = load_config()
    payload = request.get_json(force=True, silent=True) or {}
    genre = payload.get("genre", "random")
    writer_model = (payload.get("writer_model") or "").strip()

    script_path = cfg.get("script_path", "")
    if not script_path or not os.path.isfile(script_path):
        return jsonify({"error": f"Script not found: '{script_path}'. Set it in Settings first."}), 400
    if writer_model not in WRITER_MODEL_VALUES:
        return jsonify({"error": f"Unknown writer model '{writer_model}'."}), 400

    structure, structure_error = _structure_for_request(cfg, payload)
    if structure_error:
        return jsonify({"error": structure_error}), 400

    cmd = [cfg["python_exe"], "-u", script_path]
    if genre and genre != "random":
        if genre not in GENRE_POOL:
            return jsonify({"error": f"Unknown genre '{genre}'."}), 400
        cmd += ["--genre", genre]
    if writer_model:
        cmd += ["--writer-model", writer_model]
    cmd += structure_to_args(structure)

    label = f"Run Book Generator ({genre})" + (f" [model: {writer_model}]" if writer_model else "")
    job_id = start_job(label, [cmd], cwd=os.path.dirname(script_path))
    return jsonify({"job_id": job_id})


@app.route("/api/run/pipeline", methods=["POST"])
def run_pipeline():
    """Full chain: writer+cover -> editor -> scorer, via pipeline_chain.py,
    in one job — the dashboard equivalent of the copy-paste terminal commands
    this used to require. Confirmed working end-to-end on real hardware
    2026-08-12 before being wired in here."""
    cfg = load_config()
    payload = request.get_json(force=True, silent=True) or {}
    genre = payload.get("genre", "random")
    suggest_only = bool(payload.get("suggest_only", False))
    writer_model = (payload.get("writer_model") or "").strip()

    script_path = cfg.get("pipeline_script_path", "")
    if not script_path or not os.path.isfile(script_path):
        return jsonify({"error": f"pipeline_chain.py not found: '{script_path}'. Set it in Settings first."}), 400
    if writer_model not in WRITER_MODEL_VALUES:
        return jsonify({"error": f"Unknown writer model '{writer_model}'."}), 400
    structure, structure_error = _structure_for_request(cfg, payload)
    if structure_error:
        return jsonify({"error": structure_error}), 400

    cmd = [cfg["python_exe"], "-u", script_path]
    if genre and genre != "random":
        if genre not in GENRE_POOL:
            return jsonify({"error": f"Unknown genre '{genre}'."}), 400
        cmd += ["--genre", genre]
    if suggest_only:
        cmd.append("--suggest-only")
    if writer_model:
        cmd += ["--writer-model", writer_model]
    cmd += structure_to_args(structure)

    label = f"Full Pipeline ({genre}{', suggest-only' if suggest_only else ''}{f', model: {writer_model}' if writer_model else ''})"
    job_id = start_job(label, [cmd], cwd=os.path.dirname(script_path))
    return jsonify({"job_id": job_id})


# =====================================================================
# Routes — run the editor or the scorer against ONE already-generated book
# (Book Preview panel's "Run editor" / "Run scorer" buttons).
#
# These exist because the preview panel is where you actually notice a book
# is missing a stage: you open a rough copy, see "Not available." under
# Edited or "Not scored" on the score pill, and until now the only way to
# act on that was to leave the dashboard and run the script by hand from a
# terminal. Both routes go through the same start_job() machinery as every
# other run, so they stream into the same terminal panel, respect the same
# Stop button, and record the same ETA history.
#
# Deliberately NOT reusing pipeline_chain.py --book-dir here: that chains
# editor AND scorer together, but these buttons are per-stage by design —
# the whole point is filling in the one stage a given book is missing.
# =====================================================================

def _sibling_script(name):
    """Resolve a pipeline script that sits next to dashboard_server.py. Same
    convention /api/system-test already uses for these three scripts — they
    have no config entry of their own, unlike local-book-generator.py and
    pipeline_chain.py."""
    path = os.path.join(BASE_DIR, name)
    return path if os.path.isfile(path) else None


def _start_single_stage_job(book_slug, script_name, extra_args, label_verb):
    """Shared body of the two routes below — resolve the book folder, resolve
    the script, launch it as a job. Returns a Flask response tuple."""
    cfg = load_config()
    book_dir = resolve_book_dir(get_output_root(cfg), book_slug)
    if not book_dir:
        return jsonify({"error": "unknown book"}), 404

    script_path = _sibling_script(script_name)
    if not script_path:
        return jsonify({
            "error": f"{script_name} not found next to dashboard_server.py "
                     f"(looked in '{BASE_DIR}')."
        }), 400

    # One stage at a time on one GPU. Refusing here rather than queueing is
    # deliberate: a book generation already in flight is holding the VRAM
    # this stage needs, and starting anyway is precisely the overlap that
    # pushed Ollama onto the CPU before the 2026-08-13 fixes.
    with JOBS_LOCK:
        busy = [j for j in JOBS.values() if j.status == "running"]
    if busy:
        return jsonify({
            "error": f"'{busy[0].name}' is still running. Wait for it to finish (or press Stop) "
                     f"before starting another stage — they'd otherwise compete for the same VRAM."
        }), 409

    cmd = [cfg["python_exe"], "-u", script_path, "--book-dir", book_dir] + list(extra_args)
    label = f"{label_verb}: {os.path.basename(book_dir)}"
    job_id = start_job(label, [cmd], cwd=BASE_DIR)
    return jsonify({"job_id": job_id, "book_dir": book_dir})


@app.route("/api/run/editor", methods=["POST"])
def run_editor_on_book():
    """Editorial pass over one existing book, writing edited/ (auto-apply)."""
    payload = request.get_json(force=True, silent=True) or {}
    book_slug = str(payload.get("slug", "")).strip()
    if not book_slug:
        return jsonify({"error": "slug is required"}), 400
    # --auto-apply, because the button's whole purpose is producing an edited/
    # version the preview panel can show. Suggest-only would write a report and
    # leave the Edited pane just as empty as before the click.
    return _start_single_stage_job(book_slug, "editorial_agent.py", ["--auto-apply"], "Editor")


@app.route("/api/run/scorer", methods=["POST"])
def run_scorer_on_book():
    """Scoring + Finished Product Notes over one existing book.

    No version argument: scoring_agent.py picks its own source (edited/ if it
    exists, else the raw draft). Exposing a choice here would be a lie — the
    script doesn't take one."""
    payload = request.get_json(force=True, silent=True) or {}
    book_slug = str(payload.get("slug", "")).strip()
    if not book_slug:
        return jsonify({"error": "slug is required"}), 400
    return _start_single_stage_job(book_slug, "scoring_agent.py", [], "Scorer")


# Length bands used only to label a word count in the Book Preview ("11,604
# words · Short"). Kept here rather than imported because scoring_agent.py runs
# its whole pipeline at import time and can't be imported — the same convention
# as get_free_vram_mb / _load_title_history, which are also duplicated across
# these standalone scripts.
PREVIEW_LENGTH_BANDS = [
    (32000, "Full-length"),
    (20000, "Standard"),
    (0, "Short"),
]

_PREVIEW_WORD_RE = re.compile(r"[A-Za-z']+")


def _count_book_words(book_dir, version):
    """Total words across a book's chapter files. Returns None if that version
    doesn't exist, so the UI can distinguish 'no edited version' from '0 words'."""
    folder = book_dir if version == "raw" else os.path.join(book_dir, "edited")
    if not os.path.isdir(folder):
        return None
    names = [n for n in os.listdir(folder)
             if n.startswith("chapter_") and n.endswith(".txt")]
    if not names:
        return None
    total = 0
    for name in sorted(names):
        try:
            with open(os.path.join(folder, name), "r", encoding="utf-8", errors="replace") as f:
                total += len(_PREVIEW_WORD_RE.findall(f.read()))
        except OSError:
            continue
    return total


def _length_band(words):
    for threshold, label in PREVIEW_LENGTH_BANDS:
        if words >= threshold:
            return label
    return "Short"


@app.route("/api/books/<book_slug>/wordcount", methods=["GET"])
def book_wordcount(book_slug):
    """Word count for both versions of a book, for the Book Preview pills.

    Seeing "11,604 words · Short" immediately says what kind of book came out,
    without opening the score report."""
    cfg = load_config()
    book_dir = resolve_book_dir(get_output_root(cfg), book_slug)
    if not book_dir:
        return jsonify({"error": "unknown book"}), 404
    out = {}
    for version in ("raw", "edited"):
        words = _count_book_words(book_dir, version)
        out[version] = None if words is None else {
            "words": words,
            "band": _length_band(words),
        }
    return jsonify(out)


@app.route("/api/books/<book_slug>/stages", methods=["GET"])
def book_stages(book_slug):
    """What this book is missing, so the preview panel can decide which action
    buttons to offer. Kept as its own tiny endpoint rather than folded into
    /api/books so the preview can refresh just this after a stage finishes."""
    cfg = load_config()
    book_dir = resolve_book_dir(get_output_root(cfg), book_slug)
    if not book_dir:
        return jsonify({"error": "unknown book"}), 404
    edited_dir = os.path.join(book_dir, "edited")
    # An edited/ folder that exists but holds no chapter files means a previous
    # editor run died partway — treat that as "not edited" so the button stays
    # available rather than the panel claiming a version that isn't there.
    has_edited = os.path.isdir(edited_dir) and any(
        name.startswith("chapter_") and name.endswith(".txt")
        for name in os.listdir(edited_dir)
    )
    has_root_score = os.path.isfile(os.path.join(book_dir, "book_score.json"))
    has_edited_score = os.path.isfile(os.path.join(edited_dir, "book_score.json"))
    return jsonify({
        "has_edited": has_edited,
        "has_score": has_root_score or has_edited_score,
        # Which version the scorer WOULD score if run right now — mirrors
        # scoring_agent.py's own choose_chapter_source() preference, so the
        # button label can tell the truth about what's about to happen.
        "would_score": "edited" if has_edited else "raw",
    })


@app.route("/api/run/batch", methods=["POST"])
def run_batch():
    """Queue 2-5 books to run sequentially as ONE job — same "Run" action as
    /api/run/bookgen or /api/run/pipeline, just repeated N times back to
    back. One job means one Stop button (stops the batch, not just the
    current book) and one combined log, rather than the dashboard having to
    track and label N separate jobs. Each book gets its own identical
    command; when genre is "random", each run independently randomizes its
    own genre (local-book-generator.py's own default behavior whenever
    --genre is omitted), so a random batch isn't N copies of the same genre."""
    cfg = load_config()
    payload = request.get_json(force=True, silent=True) or {}
    genre = payload.get("genre", "random")
    full_pipeline = bool(payload.get("full_pipeline", True))
    suggest_only = bool(payload.get("suggest_only", False))
    writer_model = (payload.get("writer_model") or "").strip()

    try:
        count = int(payload.get("count", 1))
    except (TypeError, ValueError):
        return jsonify({"error": "Batch count must be a number."}), 400
    if not 1 <= count <= 5:
        return jsonify({"error": "Batch count must be between 1 and 5."}), 400

    if genre and genre != "random" and genre not in GENRE_POOL:
        return jsonify({"error": f"Unknown genre '{genre}'."}), 400
    if writer_model not in WRITER_MODEL_VALUES:
        return jsonify({"error": f"Unknown writer model '{writer_model}'."}), 400

    if full_pipeline:
        script_path = cfg.get("pipeline_script_path", "")
        script_label = "pipeline_chain.py"
    else:
        script_path = cfg.get("script_path", "")
        script_label = "local-book-generator.py"
    if not script_path or not os.path.isfile(script_path):
        return jsonify({"error": f"{script_label} not found: '{script_path}'. Set it in Settings first."}), 400

    base_cmd = [cfg["python_exe"], "-u", script_path]
    if genre and genre != "random":
        base_cmd += ["--genre", genre]
    if full_pipeline and suggest_only:
        base_cmd.append("--suggest-only")
    if writer_model:
        base_cmd += ["--writer-model", writer_model]

    # Structure resolved once PER BOOK, not once for the batch — that's what makes
    # a "random" 3-book batch produce three different lengths (product decision,
    # 2026-08-14). Validate every roll, since a custom preset is the same for all
    # of them and should fail the whole click rather than book 2.
    cmds = []
    for _ in range(count):
        structure, structure_error = _structure_for_request(cfg, payload)
        if structure_error:
            return jsonify({"error": structure_error}), 400
        cmds.append(list(base_cmd) + structure_to_args(structure))

    mode = "Full Pipeline" if full_pipeline else "Book Generator"
    suffix = ", suggest-only" if (full_pipeline and suggest_only) else ""
    model_suffix = f", model: {writer_model}" if writer_model else ""
    label = f"Batch: {count}x {mode} ({genre}{suffix}{model_suffix})"
    job_id = start_job(label, cmds, cwd=os.path.dirname(script_path))
    return jsonify({"job_id": job_id})


# =====================================================================
# Routes — book preview (item 6.2: raw vs. edited, side by side)
# =====================================================================

# Matches local-book-generator.py's own folder naming exactly:
#   f"{slug}-{run_stamp}"  where slug is lowercase/hyphenated (re.sub of
#   anything non [a-z0-9] to "-") and run_stamp is "%Y%m%d-%H%M%S". Anything
# that doesn't match this shape is rejected outright — both because a real
# book folder never looks like anything else, and as defense-in-depth
# against path traversal via a crafted book_slug in the URL.
_SAFE_BOOK_SLUG = re.compile(r"^[a-z0-9][a-z0-9-]*-\d{8}-\d{6}$")


def get_output_root(cfg):
    script_path = cfg.get("script_path", "")
    base = os.path.dirname(script_path) if script_path else BASE_DIR
    return os.path.join(base, "output_books")


def resolve_book_dir(output_root, book_slug):
    """Returns the book's absolute folder path if book_slug both matches the
    expected naming shape AND actually resolves to somewhere inside
    output_root — belt-and-suspenders against path traversal. Returns None
    (caller treats as 404) for anything that fails either check."""
    if not _SAFE_BOOK_SLUG.match(book_slug or ""):
        return None
    candidate = os.path.realpath(os.path.join(output_root, book_slug))
    root_real = os.path.realpath(output_root)
    if candidate != root_real and not candidate.startswith(root_real + os.sep):
        return None
    if not os.path.isdir(candidate):
        return None
    return candidate


def _dir_size(path):
    """Total bytes of every file under path, walked recursively. Shared by
    list_books() (per-book size, for the library panel — item 7) and
    /api/system-test's whole-library total (section 6.9), so there's one
    walk-and-sum implementation instead of two copies drifting apart."""
    total = 0
    for dirpath, _, filenames in os.walk(path):
        for fn in filenames:
            try:
                total += os.path.getsize(os.path.join(dirpath, fn))
            except OSError:
                pass
    return total


def _book_overall_score(book_dir):
    """Best-available overall score for the library list (item 7) — prefers
    edited/book_score.json (what scoring_agent.py itself prefers whenever
    both a raw and an edited version exist), falls back to the root
    book_score.json (covers raw-only books and repolish-scored books, where
    scored_version is "raw_draft" or "repolish" respectively). Returns None
    if the book hasn't been scored at all yet — the frontend shows
    'unscored' for that case rather than a bogus 0."""
    for score_path in (
        os.path.join(book_dir, "edited", "book_score.json"),
        os.path.join(book_dir, "book_score.json"),
    ):
        if os.path.isfile(score_path):
            try:
                with open(score_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                return data.get("overall_score")
            except Exception:
                pass
    return None


@app.route("/api/books", methods=["GET"])
def list_books():
    cfg = load_config()
    output_root = get_output_root(cfg)
    if not os.path.isdir(output_root):
        return jsonify([])

    entries = []
    for name in os.listdir(output_root):
        book_dir = os.path.join(output_root, name)
        outline_path = os.path.join(book_dir, "outline.json")
        if not os.path.isdir(book_dir) or not os.path.isfile(outline_path):
            continue
        try:
            with open(outline_path, "r", encoding="utf-8") as f:
                outline = json.load(f)
        except Exception:
            outline = {}
        try:
            modified_at = os.path.getmtime(book_dir)
        except OSError:
            modified_at = 0
        entries.append({
            "slug": name,
            "title": outline.get("title", name),
            "genre": outline.get("genre", ""),
            "modified_at": modified_at,
            "has_edited": os.path.isdir(os.path.join(book_dir, "edited")),
            # Added for item 7 (library management panel) — size_bytes and
            # overall_score are extra fields, so this stays backward
            # compatible with the Book Preview dropdown, which only ever
            # read slug/title/modified_at from this same response.
            "size_bytes": _dir_size(book_dir),
            "overall_score": _book_overall_score(book_dir),
        })
    entries.sort(key=lambda e: e["modified_at"], reverse=True)
    return jsonify(entries)


# =====================================================================
# Item 7 — Book Library management (list is /api/books above; this is the
# delete side). Always backs the book's folder up to a timestamped zip
# under output_books/deleted_book_backups/ before removing it — same
# backup-then-clear shape as section 6.11's title-history Clear button,
# per your call on this pass. There's deliberately no "skip the backup"
# option here.
# =====================================================================

DELETED_BOOK_BACKUPS_DIRNAME = "deleted_book_backups"


@app.route("/api/library/delete", methods=["POST"])
def library_delete_book():
    cfg = load_config()
    payload = request.get_json(force=True, silent=True) or {}
    slug = payload.get("slug", "")

    output_root = get_output_root(cfg)
    book_dir = resolve_book_dir(output_root, slug)
    if not book_dir:
        return jsonify({"error": "unknown book"}), 404

    backups_dir = os.path.join(output_root, DELETED_BOOK_BACKUPS_DIRNAME)
    try:
        os.makedirs(backups_dir, exist_ok=True)
    except OSError as e:
        return jsonify({"error": f"Could not create backup folder: {e}"}), 500

    stamp = time.strftime("%Y%m%d-%H%M%S")
    backup_path = os.path.join(backups_dir, f"{slug}_{stamp}.zip")
    try:
        with zipfile.ZipFile(backup_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for dirpath, _, filenames in os.walk(book_dir):
                for fn in filenames:
                    fpath = os.path.join(dirpath, fn)
                    arcname = os.path.join(slug, os.path.relpath(fpath, book_dir))
                    zf.write(fpath, arcname)
    except Exception as e:
        # Nothing removed yet at this point — safe to just report and stop.
        return jsonify({"error": f"Backup failed, nothing was deleted: {e}"}), 500

    try:
        shutil.rmtree(book_dir)
    except Exception as e:
        # The backup zip DID succeed here, so the book isn't actually lost —
        # say so explicitly rather than leaving the operator to wonder.
        return jsonify({
            "error": f"Backup saved to {backup_path}, but removing the original folder failed: {e}"
        }), 500

    return jsonify({"ok": True, "backup_path": backup_path})


def _covers_available(book_dir):
    """Which of the 3 cover options (see local-book-generator.py's Stage C)
    actually exist for this book — checks for either the .jpg (upload-ready,
    added 2026-08-12) or the original .png, since a book generated before
    that fix will only have the .png. Covers aren't per-version (the editor
    only copies the same 3 images into edited/, never regenerates them), so
    there's one shared set regardless of which pane you're previewing."""
    return [
        n for n in (1, 2, 3)
        if os.path.isfile(os.path.join(book_dir, f"cover_option_{n}.jpg"))
        or os.path.isfile(os.path.join(book_dir, f"cover_option_{n}.png"))
    ]


@app.route("/api/books/<book_slug>/chapters", methods=["GET"])
def book_chapters(book_slug):
    cfg = load_config()
    book_dir = resolve_book_dir(get_output_root(cfg), book_slug)
    if not book_dir:
        return jsonify({"error": "unknown book"}), 404
    try:
        with open(os.path.join(book_dir, "outline.json"), "r", encoding="utf-8") as f:
            outline = json.load(f)
    except Exception as e:
        return jsonify({"error": f"Could not read outline.json: {e}"}), 500
    chapters = [
        {"chapter_number": c.get("chapter_number"), "title": c.get("title", "")}
        for c in outline.get("chapters", [])
    ]
    return jsonify({
        "title": outline.get("title", book_slug),
        "genre": outline.get("genre", ""),
        "has_edited": os.path.isdir(os.path.join(book_dir, "edited")),
        "chapters": chapters,
        "covers_available": _covers_available(book_dir),
    })


@app.route("/api/books/<book_slug>/cover/<int:option_number>", methods=["GET"])
def book_cover_image(book_slug, option_number):
    cfg = load_config()
    book_dir = resolve_book_dir(get_output_root(cfg), book_slug)
    if not book_dir:
        return jsonify({"error": "unknown book"}), 404
    if option_number not in (1, 2, 3):
        return jsonify({"error": "option_number must be 1, 2, or 3"}), 400

    # Prefer .jpg (the upload-ready format, added 2026-08-12) but fall back to
    # .png for books generated before that fix shipped.
    for ext, mimetype in ((".jpg", "image/jpeg"), (".png", "image/png")):
        fpath = os.path.join(book_dir, f"cover_option_{option_number}{ext}")
        if os.path.isfile(fpath):
            return send_file(fpath, mimetype=mimetype)
    return jsonify({"error": f"cover_option_{option_number} not found"}), 404


@app.route("/api/books/<book_slug>/chapter/<int:chapter_number>", methods=["GET"])
def book_chapter_text(book_slug, chapter_number):
    cfg = load_config()
    book_dir = resolve_book_dir(get_output_root(cfg), book_slug)
    if not book_dir:
        return jsonify({"error": "unknown book"}), 404
    version = request.args.get("version", "raw")
    if version not in ("raw", "edited"):
        return jsonify({"error": "version must be 'raw' or 'edited'"}), 400

    version_dir = os.path.join(book_dir, "edited") if version == "edited" else book_dir
    if version == "edited" and not os.path.isdir(version_dir):
        return jsonify({"available": False, "reason": "No edited/ version yet — run the editor pass first."})

    fpath = os.path.join(version_dir, f"chapter_{chapter_number:02d}.txt")
    if not os.path.isfile(fpath):
        return jsonify({"available": False, "reason": f"chapter_{chapter_number:02d}.txt not found."})
    with open(fpath, "r", encoding="utf-8", errors="replace") as f:
        text = f.read()
    return jsonify({"available": True, "text": text})


@app.route("/api/books/<book_slug>/score", methods=["GET"])
def book_score(book_slug):
    """Only one version of a book ever actually gets scored per scoring run
    (scoring_agent.py prefers edited/ if present, else the raw draft — see
    project plan item 6.2) — so this deliberately does NOT synthesize a
    score for whichever version wasn't the one scored. version=edited only
    returns data if edited/book_score.json exists (only written when that
    was the scored version); version=raw only returns data if the root
    book_score.json's own scored_version field says "raw_draft" — a root
    book_score.json that actually describes the edited version does NOT
    get reported as the raw draft's score."""
    cfg = load_config()
    book_dir = resolve_book_dir(get_output_root(cfg), book_slug)
    if not book_dir:
        return jsonify({"error": "unknown book"}), 404
    version = request.args.get("version", "raw")
    if version not in ("raw", "edited"):
        return jsonify({"error": "version must be 'raw' or 'edited'"}), 400

    if version == "edited":
        score_path = os.path.join(book_dir, "edited", "book_score.json")
        if not os.path.isfile(score_path):
            return jsonify({"available": False})
        with open(score_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return jsonify({"available": True, **data})

    score_path = os.path.join(book_dir, "book_score.json")
    if not os.path.isfile(score_path):
        return jsonify({"available": False})
    with open(score_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if data.get("scored_version") != "raw_draft":
        return jsonify({"available": False, "reason": "Only the edited version was scored for this book."})
    return jsonify({"available": True, **data})


if __name__ == "__main__":
    if not os.path.isfile(CONFIG_PATH):
        save_config(DEFAULT_CONFIG)
    print("Local Book Forge dashboard running at http://127.0.0.1:8765")
    app.run(host="127.0.0.1", port=8765, debug=False, threaded=True)
