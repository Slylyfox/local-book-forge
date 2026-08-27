# Architecture

## Pipeline stages

```
Stage A   outline generation          CrewAI + Ollama    → outline.json
Stage B   chapter writing             Ollama direct      → chapter_NN.txt
Stage C   cover art                   AUTOMATIC1111/SDXL → cover.png + options
Stage D   manuscript assembly         python-docx        → manuscript.docx

          editorial pass              Ollama direct      → edited/, editorial_review.txt
          quality scoring             deterministic+LLM  → book_score.json, product notes
```

Each stage runs standalone against a book folder, which matters more than it
sounds: a failed scoring run can be re-run without regenerating the book, and
the editorial pass can be pointed at any existing draft.

## Why the LLM calls are split across two mechanisms

**CrewAI** handles the outline and the analysis passes. These return JSON and
benefit from its retry handling — a malformed outline is retried and repaired
rather than crashing the run.

**Direct `/api/chat` calls** handle chapter writing and revision. These are one
prompt in, prose out, and CrewAI adds nothing except a proven bug: litellm
silently drops `max_tokens` for Ollama models. A chapter capped at 2,730 words
came back at 8,759. On the revision side the same defect truncated long rewrites,
which the length check then rejected as "too short".

Direct calls also enable streaming, which is what makes early collapse abort
possible — the writer can kill a degenerating generation partway rather than
waiting for a full-length bad chapter.

## VRAM as an architectural constraint

The reference machine has 8GB of VRAM, and the image model and language model
cannot both be resident. This is not a tuning detail; it shapes stage ordering.

Stage transitions explicitly:

1. unload the other service's model (`unload_a1111_checkpoint()` /
   `unload_ollama_model()` — Ollama holds a model ~5 minutes after last use by
   default via `keep_alive`)
2. poll `nvidia-smi` until VRAM actually frees (`wait_for_free_vram()`)
3. only then start the next stage

Without this, **Ollama silently falls back to CPU inference rather than
erroring**. See `vram-and-local-inference.md`.

## Book folder layout

```
output_books/<slug>-<timestamp>/
  outline.json              chapter beats, characters, word targets, emotional pitch
  model_info.json           which models produced this book
  chapter_NN.txt            raw draft
  _raw_art*.png             SDXL output before typography
  cover_option_N.png/.jpg   three cover candidates
  cover.png / cover.jpg     the selected cover
  manuscript_raw.txt        assembled raw draft
  manuscript.docx           formatted manuscript
  editorial_review.txt      every issue found, by category and severity
  style_sheet.json          accumulated continuity facts across chapters
  edited/                   post-editorial chapters + its own score + docx
  book_score.json           all 18 sub-metrics
  book_score_report.txt     human-readable score with reasoning
  Finished Product Notes.txt listing metadata — blurb, categories, keywords
```

The `edited/` subfolder mirrors the book rather than overwriting it, so the raw
draft and the edited version can be compared and scored independently. That
comparison is how several of the findings in the README were made.

## Continuity between chapters

Chapter N's prompt receives the last 1,500 characters of chapter N−1, so scenes
connect. The seed is **vetted before use** on two independent axes:

- **Coherence** — the last 800 words are run through the repetition-collapse
  detector. A degenerate ending is not carried forward. This was added after a
  single bad chapter poisoned every chapter after it: chapter 4 drifted onto one
  word, chapter 5 was handed that prose and told to continue naturally, and
  reproduced the same collapse on all three retries. Rerolling cannot escape a
  bad seed.
- **Style** — the tail's average sentence length is checked against a ceiling.
  See `continuity-seed-style-ratchet.md`.

Failing either check, continuity falls back to the previous chapter's outline
summary, which carries the events without carrying the prose.

A separate `style_sheet.json` accumulates established facts, character
descriptions and style terms across the editorial pass, giving the reviewer
book-level context that a per-chapter pass would otherwise lack.

## Dashboard

Flask with server-sent events. Two implementation details worth noting:

**Bounded queue with drop-oldest.** Job output goes through a queue capped at
5,000 entries. A runaway job that outpaces the consumer drops old lines rather
than growing memory without limit.

**Terminal rendering.** A DOM node cap alone was insufficient — setting
`scrollTop = scrollHeight` per line forced a synchronous reflow, and a
30,000-line replay still froze the tab (a Playwright probe timed out at 30s).
Fixed by batching appends through `requestAnimationFrame`, which took the same
probe to 2.3ms.

## Testing approach

There is no unit-test suite in this repo. Verification was done with:

- **Mock HTTP servers** standing in for Ollama, to confirm exactly what options
  reach the model — this is how the dropped-`max_tokens` bug was proven.
- **Replay harnesses** running new logic over the real generated corpus, to
  check a change fires where intended before spending an hour on a live run.
- **Calibration against labelled data** — the collapse detector was tuned on 19
  real chapters and validated at 5/5 known-bad flagged, 0/14 known-good.
- **Playwright** against a fixture server for the dashboard.

These live outside the repo as working harnesses. The honest position: this
project is verified empirically against its own output corpus rather than by
a maintained test suite, which suits a single-operator system where the real
oracle is "does the book read well" — but it is the first thing that would need
to change to make this maintainable by anyone else.
