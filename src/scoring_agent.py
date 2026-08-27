import sys

# Force UTF-8 stdout/stderr before anything else runs — same reasoning as
# local-book-generator.py and editorial_agent.py: on Windows the default console
# codepage (cp1252) can't encode unicode CrewAI's console output uses, and this
# must happen before crewai is imported since some of its output can fire on import.
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
    sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)

import argparse
import glob
import json
import math
import os
import re
import shutil
import subprocess
import time
from collections import Counter

import requests  # added 2026-08-13 for the A1111/Ollama VRAM unload calls

from crewai import Agent, Crew, Process, Task, LLM

try:
    sys.stdout.reconfigure(line_buffering=True, encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(line_buffering=True, encoding="utf-8", errors="replace")
except AttributeError:
    pass  # Python < 3.7 fallback, not expected here

# =====================================================================
# 0. CONFIG
# =====================================================================

OUTPUT_ROOT = "output_books"

# Same model as the editor — this is analytical/judgment work (rating craft
# against a rubric), not open-ended creative generation, so the stock instruct
# model is the right fit, consistent with the editorial_agent.py decision.
#
# Also the same custom 16384-context tag as the editor (num_ctx raised from
# Ollama's 2048-token default — see Modelfile.llama31-16k and project plan
# section 5). The scorer reads full chapter text per-chapter the same way the
# editor does, so it's exposed to the identical input-truncation risk on long
# chapters. Requires the same one-time
# `ollama create llama3.1-16k -f Modelfile.llama31-16k` — shared with the
# editor, only needs to be run once total, not once per script.
SCORER_MODEL_NAME = "llama3.1-16k"

# timeout added 2026-08-13 — same reasoning as editorial_agent.py's: without a
# ceiling, one stalled Ollama request hangs the stage indefinitely with no
# output and no failure, instead of failing an attempt the retry logic can
# handle. The scorer's calls are much smaller than the editor's (short JSON out,
# not a whole rewritten chapter), so a tighter ceiling is fine here.
SCORER_REQUEST_TIMEOUT_SECONDS = 600

scorer_llm = LLM(
    model=f"ollama/{SCORER_MODEL_NAME}",
    base_url="http://localhost:11434",
    timeout=SCORER_REQUEST_TIMEOUT_SECONDS,
)


def check_ollama_models_available(
    required_models: list, ollama_url: str = "http://localhost:11434", fix_hints: dict = None
) -> None:
    """Preflight check, run once at startup before any CrewAI work begins.
    Identical to local-book-generator.py's check_ollama_models_available() — see
    that copy for the full reasoning (a missing model otherwise surfaces as a
    wall of CrewAI retry-panel noise ending in a raw traceback instead of one
    clear message, confirmed 2026-08-12).

    fix_hints: optional {model_name: fix_command} overrides for models not
    fixable with a plain `ollama pull` — added 2026-08-12 for the custom
    -16k context-window tag (see Modelfile.llama31-16k), which is built
    locally with `ollama create`, not pulled from a registry."""
    import urllib.request as _urllib_request
    import json as _json

    fix_hints = fix_hints or {}

    try:
        with _urllib_request.urlopen(f"{ollama_url.rstrip('/')}/api/tags", timeout=5) as resp:
            data = _json.loads(resp.read())
    except Exception as e:
        print(f"[FATAL] Could not reach Ollama at {ollama_url} to verify required models: {e}")
        print("        Is Ollama running? Start it (e.g. 'ollama serve') and try again.")
        sys.exit(1)

    available = {m.get("name", "") for m in data.get("models", [])}
    available_bare = {name.split(":")[0] for name in available if name}

    missing = [
        m for m in required_models
        if m not in available and m.split(":")[0] not in available_bare
    ]
    if missing:
        print(f"[FATAL] Required Ollama model(s) not found on this machine: {', '.join(missing)}")
        for m in missing:
            print(f"        Fix: {fix_hints.get(m, f'ollama pull {m}')}")
        print(f"        Currently pulled models: {', '.join(sorted(available)) or '(none)'}")
        sys.exit(1)

    print(f"[Config] Ollama models verified present: {', '.join(required_models)}")


check_ollama_models_available(
    [SCORER_MODEL_NAME],
    fix_hints={
        SCORER_MODEL_NAME: (
            "one-time local setup, not a registry pull: "
            "ollama create llama3.1-16k -f Modelfile.llama31-16k "
            "(requires 'ollama pull llama3.1' first if you haven't already)"
        ),
    },
)

scorer = Agent(
    role="Manuscript Quality Assessor",
    goal=(
        "Score genre fiction manuscripts against a professional craft rubric the way "
        "an acquisitions editor or contest judge would — precise, consistent scores "
        "with a brief justification, not vague praise or vague criticism."
    ),
    backstory=(
        "You have judged hundreds of commercial genre fiction submissions across "
        "thriller, fantasy, mystery, horror, and romance for contests and small "
        "presses. You score strictly against the criteria given, you are consistent "
        "run to run, and you always return the exact numeric fields asked for — "
        "never prose in place of a number, never a range, never skipping a field."
    ),
    llm=scorer_llm,
    verbose=True,
)

MAX_SCORE_ATTEMPTS = 3  # JSON parse retries per pass, same pattern as editorial_agent.py

FICTION_FRAMING = (
    "This is entirely fictional, original genre fiction being scored for quality — "
    "not real-world advice, instructions, or commentary.\n\n"
)

# =====================================================================
# 1. CLI
# =====================================================================

parser = argparse.ArgumentParser(description="Quality scoring pass over a local-book-generator.py output.")
parser.add_argument(
    "--book-dir",
    default=None,
    help="Path to a specific output_books/<slug-timestamp> folder. Defaults to the most "
         "recently modified book folder under output_books/.",
)
parser.add_argument(
    "--score-dir",
    default=None,
    help="Score chapters from this folder instead of the book's own raw/edited/ chapters "
         "(e.g. a repolish/ folder built by repolish_agent.py). outline.json and the style "
         "sheet are still read from --book-dir; only chapter discovery and every output file "
         "(book_score.json, book_score_report.txt, Finished Product Notes.txt) are redirected "
         "to this folder. When set, the book's own edited/ score files are left untouched.",
)
args = parser.parse_args()


def find_latest_book_dir(output_root: str) -> "str | None":
    """Most recently modified subfolder of output_root that actually looks like a
    book (has an outline.json) — filters out stray folders so a typo'd/incomplete
    directory doesn't get silently picked."""
    if not os.path.isdir(output_root):
        return None
    candidates = [
        os.path.join(output_root, name)
        for name in os.listdir(output_root)
        if os.path.isdir(os.path.join(output_root, name))
        and os.path.isfile(os.path.join(output_root, name, "outline.json"))
    ]
    if not candidates:
        return None
    candidates.sort(key=os.path.getmtime, reverse=True)
    return candidates[0]


BOOK_DIR = args.book_dir or find_latest_book_dir(OUTPUT_ROOT)
if not BOOK_DIR or not os.path.isdir(BOOK_DIR) or not os.path.isfile(os.path.join(BOOK_DIR, "outline.json")):
    print(
        f"[ERROR] Could not find a book folder to score. Looked in '{OUTPUT_ROOT}' for the most "
        f"recently modified subfolder containing an outline.json, or pass one explicitly with "
        f"--book-dir \"path\\to\\output_books\\your-book-folder\"."
    )
    sys.exit(1)

SCORE_DIR = args.score_dir or None
if SCORE_DIR and not os.path.isdir(SCORE_DIR):
    print(f"[ERROR] --score-dir '{SCORE_DIR}' does not exist. It should be a folder of "
          f"chapter_NN.txt files (e.g. a repolish/ folder), created before scoring_agent.py runs.")
    sys.exit(1)

print(f"[Config] Scoring book at: {BOOK_DIR}")
if SCORE_DIR:
    print(f"[Config] Chapter source + output redirected to --score-dir: {SCORE_DIR}")
print(f"[Config] Scorer model: {SCORER_MODEL_NAME}")

# =====================================================================
# 2. VRAM UTILITIES (same pattern as local-book-generator.py / editorial_agent.py)
# =====================================================================


def get_free_vram_mb() -> "int | None":
    """Query free VRAM in MB via nvidia-smi. Returns None (not an exception) if
    nvidia-smi isn't available, so callers can treat 'unknown' as 'proceed
    cautiously' instead of crashing on machines without it on PATH."""
    try:
        result = subprocess.run(
            ["nvidia-smi", "--query-gpu=memory.free", "--format=csv,noheader,nounits"],
            capture_output=True, text=True, timeout=10,
        )
        if result.returncode == 0 and result.stdout.strip():
            return int(result.stdout.strip().splitlines()[0])
    except (subprocess.SubprocessError, FileNotFoundError, ValueError, OSError):
        pass
    return None


def unload_a1111_checkpoint(a1111_url: str = "http://127.0.0.1:7860",
                            max_wait_seconds: int = 25) -> None:
    """Free VRAM held by AUTOMATIC1111 before this script loads its own model.

    Corrects the same wrong assumption fixed in editorial_agent.py on
    2026-08-13 — that a script which only talks to Ollama has no VRAM handoff
    to manage. The handoff that matters is the inbound one: cover generation
    leaves SDXL (~6GB) resident with no idle timeout, so llama3.1-16k can't fit
    on an 8GB card and Ollama silently runs it on the CPU instead of failing.
    See editorial_agent.py's copy for the full failure chain.

    Safe to call unconditionally — a no-op if A1111 isn't running, and A1111
    reloads its checkpoint by itself on the next txt2img request."""
    print("\n[VRAM] Unloading AUTOMATIC1111's checkpoint (if loaded) before the scorer model...")
    before = get_free_vram_mb()
    try:
        response = requests.post(f"{a1111_url.rstrip('/')}/sdapi/v1/unload-checkpoint", timeout=60)
        if response.status_code not in (200, 204):
            print(f"[WARN] A1111 unload-checkpoint returned HTTP {response.status_code}. "
                  f"Continuing anyway.")
    except requests.exceptions.RequestException as e:
        print(f"[VRAM] AUTOMATIC1111 not reachable ({e.__class__.__name__}) — nothing to unload.")
        return

    if before is None:
        time.sleep(3)
        return
    for _ in range(max_wait_seconds):
        time.sleep(1)
        now = get_free_vram_mb()
        if now is not None and now >= before + 500:
            print(f"[VRAM] Freed up to {now}MB available (was {before}MB).")
            return
    print(f"[WARN] VRAM didn't visibly increase after {max_wait_seconds}s "
          f"(still ~{before}MB free) — continuing anyway.")


def unload_ollama_model(model_name: str, max_wait_seconds: int = 20) -> None:
    """Drop a model from VRAM immediately (keep_alive: 0). Called on the way out
    of this script — it's the last stage in the chain, so leaving the model
    resident for Ollama's default ~5 minute idle timeout would collide with the
    NEXT BOOK's writer model in a batch run. Added 2026-08-13."""
    print(f"\n[VRAM] Unloading Ollama model '{model_name}'...")
    try:
        requests.post(
            "http://localhost:11434/api/generate",
            json={"model": model_name, "keep_alive": 0},
            timeout=15,
        )
    except requests.exceptions.RequestException as e:
        print(f"[WARN] Could not reach Ollama to unload the model ({e}). Continuing anyway.")
        return
    for _ in range(max_wait_seconds):
        time.sleep(1)
        if (get_free_vram_mb() or 0) >= 4000:
            return


_free_vram = get_free_vram_mb()
if _free_vram is not None:
    print(f"[VRAM] Free before starting: {_free_vram}MB")

# Inbound handoff — see unload_a1111_checkpoint() above.
unload_a1111_checkpoint()

_free_vram = get_free_vram_mb()
if _free_vram is not None:
    print(f"[VRAM] Free after A1111 unload: {_free_vram}MB")
    if _free_vram < 6000:
        print(f"[WARN] Only {_free_vram}MB VRAM free. '{SCORER_MODEL_NAME}' needs roughly 6.5GB "
              f"to stay fully on the GPU; below that Ollama quietly offloads to the CPU rather "
              f"than erroring. Check for a leftover A1111 or ollama process holding VRAM.")

# =====================================================================
# 3. JSON PARSING HELPERS (same retry/repair pattern as editorial_agent.py)
# =====================================================================


def _repair_common_json_mistakes(text: str) -> str:
    """Local models occasionally leave a trailing comma before a closing brace/
    bracket, which breaks json.loads even though the rest of the document is
    fine. Strip those before giving up."""
    return re.sub(r",(\s*[}\]])", r"\1", text)


def _parse_json_response(raw_text: str) -> dict:
    """Pull the JSON object out of the model's response, tolerating stray text/
    fences, same approach as local-book-generator.py's parse_outline()."""
    match = re.search(r"\{.*\}", raw_text, re.DOTALL)
    if not match:
        raise ValueError("No JSON object found in response.")
    json_text = match.group(0)
    try:
        return json.loads(json_text)
    except json.JSONDecodeError:
        return json.loads(_repair_common_json_mistakes(json_text))


def run_json_pass(description: str, expected_output: str, required_keys: tuple) -> "dict | None":
    """Run one scorer task expecting a JSON response, retrying up to
    MAX_SCORE_ATTEMPTS times on parse failure. Returns None (not an exception) on
    total failure so the caller can substitute neutral defaults instead of
    crashing the whole run over one bad response — same degrade-gracefully
    philosophy as editorial_agent.py's run_json_pass."""
    last_error = None
    for attempt in range(1, MAX_SCORE_ATTEMPTS + 1):
        task = Task(description=description, expected_output=expected_output, agent=scorer)
        crew = Crew(agents=[scorer], tasks=[task], process=Process.sequential, verbose=True)
        raw = str(crew.kickoff())
        try:
            data = _parse_json_response(raw)
            if not all(k in data for k in required_keys):
                raise ValueError(f"Response missing required key(s) {required_keys}.")
            return data
        except (ValueError, json.JSONDecodeError) as e:
            last_error = e
            print(f"[WARN] Scoring pass attempt {attempt}/{MAX_SCORE_ATTEMPTS} failed to parse ({e}).")
            if attempt < MAX_SCORE_ATTEMPTS:
                print("[WARN] Retrying...")
    print(f"[WARN] All {MAX_SCORE_ATTEMPTS} attempts failed to parse a valid response ({last_error}). "
          f"Substituting neutral (50) scores for this pass.")
    return None


def clamp_score(value, default: int = 50) -> int:
    """Local models don't always comply with 'return a number 0-100' as strictly
    as asked (we saw the same thing with the editor's category/severity enums) —
    coerce whatever comes back to a valid int in [0, 100], or fall back to a
    neutral default rather than letting a bad value (a string, a negative number,
    a value >100) corrupt an average."""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return default
    if math.isnan(v):
        return default
    return max(0, min(100, round(v)))


# =====================================================================
# 4. LOADING THE BOOK: outline, chapters (prefers the editor's edited/ version
#    if present), story bible / style sheet (read-only — this script never
#    writes to it)
# =====================================================================


def load_outline(book_dir: str) -> dict:
    with open(os.path.join(book_dir, "outline.json"), "r", encoding="utf-8") as f:
        return json.load(f)


def parse_main_characters(main_characters: list) -> list:
    """outline.json stores main_characters as ["Name: description", ...] strings —
    split each into a {"name", "description"} dict, tolerating entries with no
    colon (kept whole as the name, empty description) rather than raising."""
    parsed = []
    for entry in main_characters or []:
        if ":" in entry:
            name, desc = entry.split(":", 1)
            parsed.append({"name": name.strip(), "description": desc.strip()})
        else:
            parsed.append({"name": entry.strip(), "description": ""})
    return parsed


CHAPTER_FILE_PATTERN = re.compile(r"chapter_(\d+)\.txt$")


def discover_chapters(chapters_dir: str) -> list:
    """Find chapter_NN.txt files in the given directory, sorted numerically by
    chapter number. Works against either the book's root folder (raw draft) or
    its edited/ subfolder (post-editor revision) — caller decides which."""
    found = []
    for path in glob.glob(os.path.join(chapters_dir, "chapter_*.txt")):
        match = CHAPTER_FILE_PATTERN.search(os.path.basename(path))
        if match:
            found.append((int(match.group(1)), path))
    found.sort(key=lambda pair: pair[0])
    return found


def choose_chapter_source(book_dir: str) -> "tuple[list, bool]":
    """Prefer the editor's revised chapters in edited/ if they exist (that's the
    version that would actually get published, per the stated pipeline order:
    writer -> cover -> editor -> scorer), falling back to the raw draft if the
    editor hasn't been run yet (or was run in suggest-only mode)."""
    edited_dir = os.path.join(book_dir, "edited")
    if os.path.isdir(edited_dir):
        edited_chapters = discover_chapters(edited_dir)
        if edited_chapters:
            return edited_chapters, True
    return discover_chapters(book_dir), False


def split_chapter_file(path: str) -> tuple:
    """Chapter files are written as '{title}\\n\\n{body}' — split back into
    (title, body) so scoring reviews prose only, not the title line repeated as
    if it were the first sentence."""
    with open(path, "r", encoding="utf-8") as f:
        raw = f.read()
    parts = raw.split("\n\n", 1)
    if len(parts) == 2:
        return parts[0].strip(), parts[1]
    return "", raw


STYLE_SHEET_FILENAME = "style_sheet.json"


def load_style_sheet_readonly(book_dir: str, outline: dict) -> dict:
    """Read the editor's story bible if it exists (richer — has facts accumulated
    chapter by chapter); otherwise build an equivalent structure straight from the
    outline. Read-only: this script never writes to style_sheet.json, that's the
    editor's file to own."""
    path = os.path.join(book_dir, STYLE_SHEET_FILENAME)
    if os.path.isfile(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "title": outline.get("title", ""),
        "genre": outline.get("genre", ""),
        "premise": outline.get("premise", ""),
        "characters": parse_main_characters(outline.get("main_characters", [])),
        "established_facts": [],
        "style_terms": [],
    }


# =====================================================================
# 5. DETERMINISTIC METRICS (computed in Python, not left to the local model —
#    same philosophy as editorial_agent.py's typography fixes: this is
#    precise, reproducible math, and a local 8B model tends to eyeball this
#    kind of thing inconsistently run to run)
# =====================================================================

WORD_PATTERN = re.compile(r"[A-Za-z']+")
VOWEL_GROUP_PATTERN = re.compile(r"[aeiouy]+")


def _count_words(text: str) -> int:
    return len(WORD_PATTERN.findall(text))


def _count_syllables(word: str) -> int:
    """Heuristic vowel-group syllable counter. Not phonetically perfect (no
    dictionary lookups, no exceptions list) but consistent and good enough for a
    manuscript-length readability estimate — the same tradeoff Flesch-Kincaid
    calculators have made for decades."""
    word = word.lower()
    groups = VOWEL_GROUP_PATTERN.findall(word)
    count = len(groups)
    if word.endswith("e") and count > 1:
        count -= 1
    return max(count, 1)


def split_sentences(text: str) -> list:
    """Simple sentence splitter: break on ./!/? followed by whitespace. Not
    dialogue-punctuation-aware (a line like 'Stop!' she said. counts as one
    sentence break at the '!'), which is an acceptable approximation for a
    manuscript-wide statistical measure rather than a per-sentence audit."""
    clean = re.sub(r"\s+", " ", text).strip()
    if not clean:
        return []
    pieces = re.split(r"(?<=[.!?])\s+", clean)
    return [p.strip() for p in pieces if p.strip()]


GENRE_READABILITY_TARGETS = {
    # (min grade, max grade) — Flesch-Kincaid Grade Level. Rough genre-convention
    # bands; tune freely, these are a reasonable starting point not a hard rule.
    "cozy mystery": (5, 8),
    "romantic suspense": (6, 9),
    "sci-fi thriller": (7, 9),
    "dystopian survival": (6, 9),
    "noir detective": (7, 9),
    "space opera": (7, 10),
    "psychological horror": (7, 10),
    "epic fantasy": (8, 11),
}
DEFAULT_READABILITY_TARGET = (7, 10)


def flesch_kincaid_grade(text: str) -> float:
    sentences = split_sentences(text)
    words = WORD_PATTERN.findall(text)
    if not sentences or not words:
        return 0.0
    syllable_count = sum(_count_syllables(w) for w in words)
    grade = 0.39 * (len(words) / len(sentences)) + 11.8 * (syllable_count / len(words)) - 15.59
    return round(grade, 1)


def score_readability(text: str, genre: str) -> dict:
    grade = flesch_kincaid_grade(text)
    lo, hi = GENRE_READABILITY_TARGETS.get(genre, DEFAULT_READABILITY_TARGET)
    if lo <= grade <= hi:
        score = 100
    else:
        deviation = (lo - grade) if grade < lo else (grade - hi)
        score = max(0, 100 - round(deviation * 12))
    return {"score": score, "grade_level": grade, "target_band": f"{lo}-{hi}"}


def score_sentence_variety(text: str) -> dict:
    sentences = split_sentences(text)
    lengths = [len(WORD_PATTERN.findall(s)) for s in sentences]
    lengths = [l for l in lengths if l > 0]
    if len(lengths) < 2:
        return {"score": 50, "stdev_words": 0.0}
    mean = sum(lengths) / len(lengths)
    variance = sum((l - mean) ** 2 for l in lengths) / len(lengths)
    stdev = variance ** 0.5
    # Band recalibrated 2026-08-13. The original band (4-10, then -8/word above)
    # was set for prose in the abstract, but every real book this pipeline has
    # produced measures 15.2-22.0 words of standard deviation, scoring 4-59 —
    # i.e. the metric flagged all seven books as defective, which is the
    # signature of a miscalibrated threshold rather than seven bad books.
    #
    # The band was too low because it doesn't account for what a full manuscript
    # actually contains: dialogue exchanges ("Don't." / "Why not?") sit at 1-3
    # words while descriptive narration runs 25-40, and mixing the two in one
    # population produces a standard deviation well above 10 — in commercial
    # genre fiction, typically 12-20. High variance there is the deliberate
    # rhythm the style pass is explicitly asking the writer for; penalising it
    # scored against the thing the rest of the pipeline works to produce.
    #
    # New band: 6-18 is the healthy range, with gentler slopes on both sides.
    # Genuinely monotonous prose (everything ~15 words, stdev under 5) and
    # genuinely erratic prose (stdev over 25) both still score poorly.
    if 6 <= stdev <= 18:
        score = 100
    elif stdev < 6:
        score = max(0, 100 - round((6 - stdev) * 12))
    else:
        score = max(0, 100 - round((stdev - 18) * 5))
    return {"score": score, "stdev_words": round(stdev, 1)}


FILLER_WORDS = {
    "just", "really", "very", "suddenly", "actually", "literally", "basically",
    "somewhat", "quite", "rather", "extremely", "totally", "definitely", "simply",
}
FILTER_PHRASE_PATTERNS = [
    re.compile(r"\b(he|she|they)\s+(saw|felt|heard|noticed|realized|wondered|watched)\b", re.IGNORECASE),
]


def score_filler_words(text: str) -> dict:
    words = [w.lower() for w in WORD_PATTERN.findall(text)]
    n_words = len(words) or 1
    filler_count = sum(1 for w in words if w in FILLER_WORDS)
    filter_count = sum(len(p.findall(text)) for p in FILTER_PHRASE_PATTERNS)
    rate_per_1000 = (filler_count + filter_count) / n_words * 1000
    score = max(0, 100 - round(max(0.0, rate_per_1000 - 3) * 8))
    return {"score": score, "rate_per_1000_words": round(rate_per_1000, 1)}


DIALOGUE_PATTERN = re.compile(r"[“\"]([^”\"]*)[”\"]")


def score_dialogue_ratio(text: str) -> dict:
    dialogue_spans = DIALOGUE_PATTERN.findall(text)
    dialogue_words = sum(_count_words(d) for d in dialogue_spans)
    total_words = _count_words(text) or 1
    ratio = dialogue_words / total_words
    lo, hi = 0.20, 0.55
    if lo <= ratio <= hi:
        score = 100
    else:
        deviation = (lo - ratio) if ratio < lo else (ratio - hi)
        score = max(0, 100 - round(deviation * 300))
    return {"score": score, "dialogue_pct": round(ratio * 100, 1)}


def score_repetitive_starters(text: str) -> dict:
    sentences = split_sentences(text)
    starters = []
    for s in sentences:
        words = WORD_PATTERN.findall(s)
        if words:
            starters.append(words[0].lower())
    if len(starters) < 2:
        return {"score": 100, "repeat_rate_pct": 0.0}
    repeats = sum(1 for i in range(1, len(starters)) if starters[i] == starters[i - 1])
    rate = repeats / (len(starters) - 1)
    score = max(0, 100 - round(rate * 400))
    return {"score": score, "repeat_rate_pct": round(rate * 100, 1)}


CLICHE_PHRASES = [
    r"in the blink of an eye", r"at the end of the day", r"time stood still",
    r"heart skipped a beat", r"sent shivers down", r"piercing gaze",
    r"little did (he|she|they) know", r"all hell broke loose", r"dead silence",
    r"cut like a knife", r"a chill ran down", r"heart pounded in (his|her|their) chest",
    r"eyes widened in shock", r"blood ran cold", r"gut feeling", r"sinking feeling",
    r"the calm before the storm", r"every fiber of (his|her|their) being",
    r"without warning", r"as if on cue", r"against all odds", r"in the nick of time",
]
CLICHE_PATTERN = re.compile("|".join(CLICHE_PHRASES), re.IGNORECASE)


def score_cliche_density(text: str) -> dict:
    n_words = _count_words(text) or 1
    matches = len(CLICHE_PATTERN.findall(text))
    rate_per_1000 = matches / n_words * 1000
    score = max(0, 100 - round(rate_per_1000 * 20))
    return {"score": score, "matches": matches, "rate_per_1000_words": round(rate_per_1000, 2)}


SENSORY_KEYWORDS = {
    "visual": [r"\bsaw\b", r"\bglimpse", r"\bgleam", r"\bshadow", r"\bbright", r"\bdark",
               r"\bcolou?r", r"\bglow", r"\bflicker", r"\bsilhouette", r"\bglint"],
    "auditory": [r"\bheard\b", r"\bsound", r"\bwhisper", r"\becho", r"\broar", r"\bsilence",
                 r"\bhum\b", r"\bcrash", r"\bmurmur", r"\bclatter"],
    "tactile": [r"\btouch", r"\brough", r"\bsmooth", r"\bcold\b", r"\bwarm", r"\btexture",
                r"\bgrip", r"\bbrush", r"\bache"],
    "olfactory": [r"\bsmell", r"\bscent", r"\baroma", r"\bstench", r"\breek", r"\bfragrance"],
}


def score_sensory_engagement(text: str) -> dict:
    counts = {}
    for category, patterns in SENSORY_KEYWORDS.items():
        combined = re.compile("|".join(patterns), re.IGNORECASE)
        counts[category] = len(combined.findall(text))
    present = sum(1 for c in counts.values() if c > 0)
    # Reward all 4 senses appearing at all; penalize any completely absent
    # category (usually olfactory/tactile get neglected first).
    score = round((present / 4) * 100)
    return {"score": score, "counts": counts}


TELLING_PATTERN = re.compile(
    r"\b(was|were|felt|seemed)\s+(furious|angry|sad|happy|scared|terrified|nervous|anxious|"
    r"excited|confused|worried|surprised|shocked|devastated|heartbroken|relieved|exhausted|"
    r"frustrated)\b",
    re.IGNORECASE,
)


def score_show_vs_tell(text: str) -> dict:
    n_words = _count_words(text) or 1
    matches = len(TELLING_PATTERN.findall(text))
    rate_per_1000 = matches / n_words * 1000
    score = max(0, 100 - round(rate_per_1000 * 15))
    return {"score": score, "telling_instances": matches, "rate_per_1000_words": round(rate_per_1000, 2)}


def score_scene_balance(chapter_word_counts: list) -> dict:
    """Proxy for scene-length balance using chapter length variance — the
    generated manuscripts don't mark explicit scene breaks (no '***' or similar),
    so chapter-to-chapter swings are the closest signal available without adding
    scene-break detection. Flagged as an approximation, not exact."""
    lengths = [c for c in chapter_word_counts if c > 0]
    if len(lengths) < 2:
        return {"score": 50, "coefficient_of_variation": 0.0}
    mean = sum(lengths) / len(lengths)
    variance = sum((l - mean) ** 2 for l in lengths) / len(lengths)
    cv = (variance ** 0.5) / mean if mean else 0.0
    if 0.15 <= cv <= 0.45:
        score = 100
    elif cv < 0.15:
        score = max(0, 100 - round((0.15 - cv) * 300))
    else:
        score = max(0, 100 - round((cv - 0.45) * 150))
    return {"score": score, "coefficient_of_variation": round(cv, 2)}


def score_emotional_curve(intensities: list) -> dict:
    """RETIRED 2026-08-14 — kept only so the LLM's ratings can still be reported
    as a footnote. Do not score on this; see score_pacing_variation() below.

    This took the LLM's per-chapter emotional_intensity ratings and scored their
    spread. The ratings turned out to be unusable, and the evidence is stark. On
    "Rainy Night Requiem" the outline planned pitches of 15, 30, 20, 80, 95, 50,
    100, 5 — and the rater returned 85, 95, 98, 95, 98, 96, 98, 92. The chapter
    deliberately written as the quietest in the book was rated 92/100 for
    intensity.

    The rater is saturated. Asked to judge one chapter in isolation, an 8B model
    scores any noir chapter in the 90s because it contains rain, guns and grief;
    it has no reference frame for "compared with the rest of THIS book". So the
    spread it produces is noise. The previous run's perfect 100/100 on this
    metric rested entirely on one chapter happening to come back 40 instead of
    90 — nothing about the book earned it."""
    values = [v for v in intensities if v is not None]
    if len(values) < 2:
        return {"score": None, "stdev": 0.0, "retired": True}
    mean = sum(values) / len(values)
    stdev = (sum((v - mean) ** 2 for v in values) / len(values)) ** 0.5
    return {"score": None, "stdev": round(stdev, 1), "retired": True,
            "ratings": values}


# =====================================================================
# PACING VARIATION — replaces emotional_resonance, 2026-08-14
# =====================================================================
# Measures the same underlying quality the retired metric was aiming at — does
# this book have highs and lows, or is every chapter the same register — but
# from the text rather than from a model's opinion of the text.
#
# Three signals per chapter, all things a reader physically experiences as pace:
#   * average sentence length  — short sentences read fast
#   * average paragraph length — short paragraphs read fast
#   * dialogue proportion      — dialogue reads faster than narration
# Combined into a 0-100 "pace index" per chapter; the metric is the spread of
# that index across the book.
#
# VALIDATION, on this pipeline's own output:
#   Rainy Night Requiem  stdev 24.7  <- reads with real range
#   Ashes & Ember        stdev 12.6
#   Murder at Willowbrook stdev 11.9  } pre-pitch-instruction books,
#   Golden Heirloom      stdev  8.7  } written before any of this existed
#   Stars Without End    stdev  9.3
#   Scorched Earth       stdev  6.9  <- measurably the flattest
# That ordering matches reading them. Including books written before the
# emotional-pitch instruction existed matters: they land in a sensible middle
# rather than at zero, so the metric is measuring pacing rather than merely
# detecting whether the new prompt was used.
#
# HONEST LIMITATION: two of the three signals (sentence length, dialogue) are
# things editorial_agent.py now explicitly instructs by chapter pitch, so this
# partly measures compliance with our own instruction. That's defensible —
# short sentences in a climax and longer ones in a quiet chapter is genuinely
# how commercial fiction reads, so complying IS the improvement — but it is a
# weaker claim than "this book is emotionally resonant", and the metric is named
# for what it actually measures rather than for what we'd like it to mean.
#
# Cross-checked against intent: on Rainy Night Requiem the pace index correlates
# with the outline's planned pitch at r = +0.65. The retired LLM rater managed
# roughly zero on the same book.
PACING_EXCELLENT_STDEV = 18.0   # at or above this, full marks


def _chapter_pace_index(text: str) -> float:
    """0-100, higher = faster-reading chapter."""
    words = WORD_PATTERN.findall(text)
    n = max(1, len(words))
    sentences = [s for s in split_sentences(text) if s.strip()]
    paragraphs = [p for p in text.split("\n\n") if p.strip()]
    avg_sentence = n / max(1, len(sentences))
    avg_paragraph = n / max(1, len(paragraphs))
    dialogue_share = sum(_count_words(d) for d in DIALOGUE_PATTERN.findall(text)) / n
    # Each term normalised to 0-1, clamped. Bounds chosen to span the range real
    # chapters actually occupy rather than theoretical extremes.
    fast_sentences = max(0.0, min(1.0, (35 - avg_sentence) / 25))
    fast_paragraphs = max(0.0, min(1.0, (120 - avg_paragraph) / 100))
    fast_dialogue = max(0.0, min(1.0, dialogue_share / 0.40))
    return (fast_sentences + fast_paragraphs + fast_dialogue) / 3 * 100


def score_pacing_variation(chapter_texts_in_order: list) -> dict:
    """Reward a book whose chapters vary in pace; penalise a monotone one."""
    if len(chapter_texts_in_order) < 2:
        return {"score": 50, "stdev": 0.0, "per_chapter": [],
                "detail": "too few chapters to assess pacing variation"}
    paces = [_chapter_pace_index(t) for t in chapter_texts_in_order]
    mean = sum(paces) / len(paces)
    stdev = (sum((p - mean) ** 2 for p in paces) / len(paces)) ** 0.5
    score = 100 if stdev >= PACING_EXCELLENT_STDEV else max(
        0, round(stdev / PACING_EXCELLENT_STDEV * 100))
    return {
        "score": score,
        "stdev": round(stdev, 1),
        "range": round(max(paces) - min(paces), 1),
        "per_chapter": [round(p) for p in paces],
        "detail": (f"pace index runs {min(paces):.0f}-{max(paces):.0f} across "
                   f"{len(paces)} chapters (spread {stdev:.1f}; "
                   f"{PACING_EXCELLENT_STDEV:.0f}+ is full marks)"),
    }


# =====================================================================
# 6. PLOT-HOLE / CONTINUITY SIGNAL — reuse the editor's already-computed
#    continuity findings instead of re-deriving them with a second expensive
#    LLM pass. Falls back to an LLM check only if editorial_review.txt doesn't
#    exist yet (i.e. the editor hasn't been run on this book).
# =====================================================================


# =====================================================================
# DEGENERATION DETECTION — added 2026-08-13
# =====================================================================
# Catches the repetition spiral that produced "Rashomon's Rainy Nights" — a book
# whose every chapter collapsed into looping prose (one 400-word stretch was 12.4%
# the single word "hell", type-token ratio 0.055) and which this scorer nonetheless
# rated 80/100. None of the existing 18 metrics can see that failure: readability
# scores WELL on short looping sentences, cliche density scores 100 because word
# salad contains no cliches, and show-vs-tell scores 100 because there are no
# telling verbs left. A coherence floor has to be measured directly. Duplicated verbatim from
# local-book-generator.py (same "each script runs standalone" convention as the rest
# of this codebase) so this script can refuse to publish-rate a broken
# book even if it was written before the writer-side guard existed.
#
# Two signals, measured over a sliding 400-word window:
#
#   type-token ratio  — unique words / total words in the window. Collapses
#                       when the model starts cycling the same vocabulary.
#   runaway content word — the most frequent NON-function word's share of the
#                       window. Function words ("the", "a", "and") are
#                       legitimately 4-6% of any healthy English text, so
#                       they're excluded; a content word above 5% means the
#                       model is stuck on it.
#
# Thresholds were calibrated against 19 real chapters from this pipeline's own
# output_books/ — 5 known-bad (Rashomon) and 14 known-good (Murder at
# Willowbrook, Golden Heirloom, Aurora Initiative, Raining Ashes, Stardust
# Rebellion, Echoes of Eternity). Worst-window TTR measured 0.245-0.34 on the
# bad chapters and 0.365-0.505 on the good ones, so the 0.36 floor sits in a
# clean gap: it caught all 5 bad chapters and flagged not one window in any of
# the 14 good ones. Requiring two CONSECUTIVE bad windows keeps a single
# dialogue-heavy passage (naturally low TTR — short lines, repeated names)
# from tripping it.
DEGEN_WINDOW_WORDS = 400
DEGEN_STEP_WORDS = 200
DEGEN_TTR_FLOOR = 0.36
DEGEN_WORD_CEILING = 0.05
DEGEN_MIN_CONSECUTIVE = 2

_DEGEN_WORD_RE = re.compile(r"[A-Za-z']+")
_DEGEN_STOPWORDS = {
    "the", "a", "an", "and", "or", "but", "if", "of", "to", "in", "on", "at", "for", "with",
    "from", "by", "as", "is", "was", "were", "be", "been", "are", "am", "he", "she", "it",
    "they", "them", "his", "her", "him", "their", "its", "i", "you", "we", "me", "my", "that",
    "this", "these", "those", "not", "no", "had", "has", "have", "did", "do", "does", "would",
    "could", "should", "will", "can", "there", "then", "so", "up", "out", "down", "into",
    "over", "about", "just", "like", "what", "when", "who", "which", "all", "one", "more",
    "than", "too", "very", "said", "being", "own", "only", "other", "new",
}


def detect_degeneration(text: str) -> dict:
    """Return a verdict on whether `text` collapses into repetition.

    Keys: degenerate (bool), first_bad_word_index (int|None — where the
    collapse starts, for salvaging the healthy prefix), worst_ttr,
    runaway_word, total_words, longest_run, detail (human-readable)."""
    from collections import Counter as _Counter

    words = _DEGEN_WORD_RE.findall(text.lower())
    total = len(words)
    if total < DEGEN_WINDOW_WORDS:
        # Too short to measure meaningfully — don't guess. A chapter this short
        # has other problems the length check will catch anyway.
        return {"degenerate": False, "first_bad_word_index": None, "worst_ttr": None,
                "runaway_word": None, "total_words": total, "longest_run": 0,
                "detail": f"only {total} words — too short to assess"}

    # Scan on SEVERAL window phases and keep the worst verdict, rather than one
    # grid starting at word 0.
    #
    # Why (found 2026-08-14, "Rainy Night, Lonely Soul"): a single grid makes the
    # result phase-dependent, and on a chapter sitting right at the boundary that
    # decides the verdict. Chapter 9 of that book tested CLEAN as the writer saw
    # it (the chapter body, longest_run 1) and DEGENERATE as the scorer saw it
    # (the same body with a 6-word title line prepended, longest_run 2). Shifting
    # the start by 4 words flipped it. The chapter was in fact badly collapsed —
    # 20 repetitions of "murder" in one stretch — so the writer's miss let it
    # through to the editor, and only the scorer caught it.
    #
    # Four phases at quarter-step intervals means a bad stretch is measured
    # wherever it happens to begin, so the same text gets the same answer no
    # matter what precedes it.
    phases = [0, DEGEN_STEP_WORDS // 4, DEGEN_STEP_WORDS // 2, 3 * DEGEN_STEP_WORDS // 4]
    worst_ttr = 1.0
    runaway = None
    longest_run = 0
    first_bad = None
    for phase in phases:
        bad_offsets = []
        for i in range(phase, total - DEGEN_WINDOW_WORDS + 1, DEGEN_STEP_WORDS):
            window = words[i:i + DEGEN_WINDOW_WORDS]
            ttr = len(set(window)) / len(window)
            worst_ttr = min(worst_ttr, ttr)
            content = [w for w in window if w not in _DEGEN_STOPWORDS]
            share, top = 0.0, ""
            if content:
                top, count = _Counter(content).most_common(1)[0]
                share = count / len(window)
            if ttr < DEGEN_TTR_FLOOR or share > DEGEN_WORD_CEILING:
                bad_offsets.append(i)
                if share > DEGEN_WORD_CEILING and (runaway is None or share > runaway[1]):
                    runaway = (top, round(share, 4))
        run, prev = 0, None
        for i in bad_offsets:
            if prev is not None and i - prev == DEGEN_STEP_WORDS:
                run += 1
            else:
                run = 1
            longest_run = max(longest_run, run)
            prev = i
        if bad_offsets and (first_bad is None or bad_offsets[0] < first_bad):
            first_bad = bad_offsets[0]

    degenerate = longest_run >= DEGEN_MIN_CONSECUTIVE

    # first_bad_word_index is the FIRST bad window anywhere in the text, not the
    # start of the longest bad run — and that distinction matters for salvage.
    # Truncating at the longest run's start looked right but left earlier bad
    # patches in place: on Rashomon chapter 2 it produced a 2,596-word "salvaged"
    # prefix that was still degenerate, because that chapter was already
    # collapsing well before its worst stretch. Cutting at the first bad window
    # instead makes the surviving prefix clean by construction — every window
    # entirely inside it was measured and passed.
    detail = ""
    if degenerate:
        detail = (f"repetition collapse from around word {first_bad} of {total} "
                  f"(lowest vocabulary richness {round(worst_ttr, 3)}, floor {DEGEN_TTR_FLOOR})")
        if runaway:
            detail += f"; the word '{runaway[0]}' reaches {runaway[1] * 100:.1f}% of a 400-word stretch"
    return {"degenerate": degenerate, "first_bad_word_index": first_bad,
            "worst_ttr": round(worst_ttr, 3), "runaway_word": runaway, "total_words": total,
            "longest_run": longest_run, "detail": detail}



def score_plot_holes_from_editorial_review(book_dir: str, scoring_edited: bool = False) -> "dict | None":
    """Derive a plot-integrity score from the editor's issue counts.

    IMPORTANT (bug fixed 2026-08-13): editorial_review.txt describes the RAW
    draft. It is the list of problems the editor FOUND — and, on an --auto-apply
    run, subsequently FIXED. Feeding those counts into the score of the edited
    version was measuring the wrong artifact twice over:

      1. It penalised the edited book for defects that no longer exist in the
         text being scored.
      2. It inverted the incentive — a more thorough editor finds more issues,
         which drove the score DOWN. The penalty is steep (-6 per continuity
         issue, -3 per high-severity issue), so a normal review of 11 continuity
         + 19 high-severity issues clamps this metric to 0.

    That is not hypothetical: across all seven scored books in output_books/ as
    of 2026-08-13, this metric read exactly 0 every time. A metric that is
    constant carries no information — it was functioning as a flat ~-33 penalty
    on category 4 (Plot Integrity & World Logic), which is why that category sat
    at 37-60 on every book regardless of quality.

    The fix: only use this file when scoring the RAW draft, which is the version
    it actually describes. When scoring edited/ or repolish/, return None, and
    the caller measures residual continuity on the text being scored instead
    (see residual_continuity_score()) — which is the honest question anyway:
    not "how many problems did the draft have," but "how many are left."
    """
    if scoring_edited:
        return None
    path = os.path.join(book_dir, "editorial_review.txt")
    if not os.path.isfile(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        text = f.read()

    continuity_count = 0
    category_line = re.search(r"^By category:\s*(.+)$", text, re.MULTILINE)
    if category_line:
        for part in category_line.group(1).split(","):
            part = part.strip()
            if "continuity" in part.lower():
                m = re.search(r"=(\d+)", part)
                if m:
                    continuity_count += int(m.group(1))

    high_count = 0
    severity_line = re.search(r"^By severity:\s*(.+)$", text, re.MULTILINE)
    if severity_line:
        for part in severity_line.group(1).split(","):
            part = part.strip()
            if part.lower().startswith("high"):
                m = re.search(r"=(\d+)", part)
                if m:
                    high_count += int(m.group(1))

    score = max(0, 100 - continuity_count * 6 - high_count * 3)
    return {"score": score, "continuity_issues": continuity_count, "high_severity_issues": high_count,
            "source": "editorial_review.txt"}


# =====================================================================
# 7. LLM PASSES — per-chapter narrative craft pass, and a whole-book synthesis
#    pass for the judgments that need the full story in view (arc tracking,
#    subplot resolution, thematic cohesion). The synthesis pass reasons over
#    compact per-chapter synopses + a story-bible instead of the full ~40k-word
#    manuscript, since that's well beyond what an 8GB-VRAM local 8B model can
#    reliably reason over in one pass (the same degradation we saw in the
#    writer model's own later chapters on long content).
# =====================================================================


def narrative_pass(chapter_text: str, ch_num: int, ch_title: str, style_sheet: dict) -> dict:
    characters = "\n".join(f"- {c['name']}: {c['description']}" for c in style_sheet["characters"]) or "(none listed)"
    description = (
        f"{FICTION_FRAMING}"
        f"You are scoring Chapter {ch_num} ('{ch_title}') of the novel '{style_sheet['title']}' "
        f"({style_sheet['genre']}) on narrative craft. Score each dimension 0-100 "
        f"(100 = excellent professional-quality craft, 50 = average/serviceable, 0 = badly broken).\n\n"
        f"Main characters:\n{characters}\n\n"
        f"Chapter text:\n\"\"\"\n{chapter_text}\n\"\"\"\n\n"
        "Score:\n"
        "- pacing_score: does this chapter build tension/momentum appropriately for its place in the "
        "story, rather than dragging or rushing?\n"
        "- hook_strength: how strong is the chapter's opening hook and closing beat at pulling a "
        "reader forward?\n"
        "- dialogue_distinctiveness: do different characters' dialogue sound like different people "
        "(word choice, rhythm, sentence length) rather than one uniform voice? (Score 50 if there is "
        "little/no dialogue to judge.)\n"
        "- motivation_agency: does the point-of-view character actively drive events here, rather than "
        "just reacting to things that happen to them?\n"
        "- emotional_intensity: how emotionally charged is this chapter overall (0 = flat/neutral, "
        "100 = intense high point). This is descriptive, not a quality judgment — a quiet chapter can "
        "correctly score low here.\n\n"
        "Respond with ONLY a valid JSON object (no commentary, no markdown fences) in exactly this "
        "shape:\n"
        "{\n"
        '  "pacing_score": 0-100,\n'
        '  "hook_strength": 0-100,\n'
        '  "dialogue_distinctiveness": 0-100,\n'
        '  "motivation_agency": 0-100,\n'
        '  "emotional_intensity": 0-100,\n'
        '  "chapter_synopsis": "1-2 sentence summary of what happens in this chapter, for whole-book '
        'context later"\n'
        "}"
    )
    result = run_json_pass(
        description,
        expected_output="A JSON object with the 6 fields described, matching the schema exactly.",
        required_keys=("pacing_score", "hook_strength", "dialogue_distinctiveness",
                        "motivation_agency", "emotional_intensity", "chapter_synopsis"),
    )
    if result is None:
        return {
            "pacing_score": 50, "hook_strength": 50, "dialogue_distinctiveness": 50,
            "motivation_agency": 50, "emotional_intensity": 50,
            "chapter_synopsis": "(no synopsis available — scoring pass failed for this chapter)",
            "_failed": True,
        }
    return {
        "pacing_score": clamp_score(result.get("pacing_score")),
        "hook_strength": clamp_score(result.get("hook_strength")),
        "dialogue_distinctiveness": clamp_score(result.get("dialogue_distinctiveness")),
        "motivation_agency": clamp_score(result.get("motivation_agency")),
        "emotional_intensity": clamp_score(result.get("emotional_intensity")),
        "chapter_synopsis": str(result.get("chapter_synopsis", "")).strip() or "(no synopsis provided)",
        "_failed": False,
    }


def whole_book_synthesis(style_sheet: dict, chapter_synopses: list, outline_chapters: list,
                          opening_excerpt: str, closing_excerpt: str, need_plot_hole_fallback: bool,
                          midpoint_excerpt: str = "") -> dict:
    characters = "\n".join(f"- {c['name']}: {c['description']}" for c in style_sheet["characters"]) or "(none listed)"
    facts = "\n".join(f"- {f}" for f in style_sheet.get("established_facts", [])) or "(none tracked)"
    synopses_text = "\n".join(f"Ch{n}: {s}" for n, s in chapter_synopses) or "(none available)"
    subplot_beats = "\n".join(
        f"- Ch{c.get('chapter_number', '?')}: {c.get('summary', '')}" for c in outline_chapters
    ) or "(no outline chapter summaries available)"

    plot_hole_field = ""
    plot_hole_instruction = ""
    if need_plot_hole_fallback:
        plot_hole_field = '  "plot_hole_score": 0-100,\n'
        plot_hole_instruction = (
            "- plot_hole_score: 100 = no contradictions found against the established facts list "
            "above; deduct heavily for characters/facts that contradict it (e.g. a character stated "
            "as dead reappearing alive with no explanation).\n"
        )

    # Calibration block added 2026-08-13. These three metrics were returning
    # 20-48 on all seven scored books in output_books/ — arc_tracking never once
    # cleared 48, subplot_resolution sat at 10-42 on six of seven. Two causes,
    # both in the prompt rather than in the manuscripts:
    #
    #   1. The prompt told the model it was working from summaries, then asked
    #      it to judge things a 1-2 sentence summary structurally cannot show
    #      (interiority, incremental change, thematic echo). The model correctly
    #      reported it couldn't see those — and expressed "I can't see it" as a
    #      low score, so it was grading the synopsis, not the novel.
    #   2. The 0/50/100 anchors were too thin to hold an 8B model steady. Absent
    #      concrete reference points, it defaults to harsh — "not literary
    #      fiction" collapses toward 30 rather than landing near the "average
    #      commercial genre novel" the 50 anchor is supposed to mean.
    #
    # Fixed by stating the evidentiary rule explicitly and giving each band a
    # concrete description. This is a measurement fix, not a thumb on the scale:
    # a book that genuinely drops its subplots still scores low, because the
    # bands below still say so.
    scale_anchors = (
        "Scoring scale — calibrate against COMMERCIALLY PUBLISHED genre fiction, not against "
        "literary prize winners:\n"
        "  90-100 = exceptional; the best examples in this genre.\n"
        "  75-89  = strong. Clearly publishable. Does the job well with minor soft spots.\n"
        "  60-74  = solid and serviceable. A competent, readable commercial genre novel — this "
        "is where a large share of successfully published books actually sit.\n"
        "  40-59  = noticeably weak. Real structural problems a reader would feel.\n"
        "  20-39  = badly broken. Arcs or subplots genuinely abandoned mid-book.\n"
        "  0-19   = incoherent.\n\n"
        "EVIDENCE RULE — this matters, and previous runs got it wrong: you are scoring THE NOVEL, "
        "not this summary of it. The synopses below are compressed to 1-2 sentences per chapter, "
        "so they necessarily omit interiority, incremental character change, and thematic detail "
        "that are present in the manuscript. Do NOT treat that compression as evidence of a "
        "defect. Score down only for problems you can positively identify — a subplot introduced "
        "and then never mentioned again, an ending that contradicts the setup, a protagonist in "
        "the same position at the end as the start. If the summary is simply silent on something, "
        "assume the manuscript handles it at a serviceable level and score in the 60-74 band.\n\n"
    )

    description = (
        f"{FICTION_FRAMING}"
        f"You are scoring the novel '{style_sheet['title']}' ({style_sheet['genre']}) as a whole, "
        f"working from a chapter-by-chapter synopsis plus real excerpts from the manuscript "
        f"below.\n\n"
        f"{scale_anchors}"
        f"Main characters:\n{characters}\n\n"
        f"Established facts tracked across the book:\n{facts}\n\n"
        f"Chapter-by-chapter synopsis:\n{synopses_text}\n\n"
        f"Subplots/beats the outline intended to deliver:\n{subplot_beats}\n\n"
        f"Opening excerpt (start of Chapter 1):\n\"\"\"\n{opening_excerpt}\n\"\"\"\n\n"
        f"Midpoint excerpt (from the middle chapter):\n\"\"\"\n{midpoint_excerpt}\n\"\"\"\n\n"
        f"Closing excerpt (end of the final chapter):\n\"\"\"\n{closing_excerpt}\n\"\"\"\n\n"
        "Score 0-100 each, using the bands above:\n"
        "- arc_tracking: do the main characters undergo a measurable internal/external change "
        "across the book? Compare the opening excerpt against the midpoint and closing excerpts "
        "for concrete evidence of change in the protagonist's situation, behaviour, or stance.\n"
        "- subplot_resolution: do the subplots/beats listed above actually reach a payoff by the "
        "end, rather than being dropped? Name a specific dropped thread if you score below 60.\n"
        "- thematic_cohesion: does this read as one consistent story with a throughline, rather "
        "than disconnected episodes?\n"
        f"{plot_hole_instruction}\n"
        "Respond with ONLY a valid JSON object (no commentary, no markdown fences) in exactly this "
        "shape:\n"
        "{\n"
        '  "arc_tracking": 0-100,\n'
        '  "arc_note": "one sentence justification",\n'
        '  "subplot_resolution": 0-100,\n'
        '  "subplot_note": "one sentence justification",\n'
        '  "thematic_cohesion": 0-100,\n'
        '  "thematic_note": "one sentence justification",\n'
        f"{plot_hole_field}"
        '  "_end": true\n'
        "}"
    )
    required = ["arc_tracking", "subplot_resolution", "thematic_cohesion"]
    if need_plot_hole_fallback:
        required.append("plot_hole_score")
    result = run_json_pass(
        description,
        expected_output="A JSON object with the fields described, matching the schema exactly.",
        required_keys=tuple(required),
    )
    if result is None:
        out = {
            "arc_tracking": 50, "arc_note": "(synthesis pass failed — neutral default used)",
            "subplot_resolution": 50, "subplot_note": "(synthesis pass failed — neutral default used)",
            "thematic_cohesion": 50, "thematic_note": "(synthesis pass failed — neutral default used)",
            "_failed": True,
        }
        if need_plot_hole_fallback:
            out["plot_hole_score"] = 50
            out["plot_hole_note"] = "(synthesis pass failed — neutral default used)"
        return out
    out = {
        "arc_tracking": clamp_score(result.get("arc_tracking")),
        "arc_note": str(result.get("arc_note", "")).strip(),
        "subplot_resolution": clamp_score(result.get("subplot_resolution")),
        "subplot_note": str(result.get("subplot_note", "")).strip(),
        "thematic_cohesion": clamp_score(result.get("thematic_cohesion")),
        "thematic_note": str(result.get("thematic_note", "")).strip(),
        "_failed": False,
    }
    if need_plot_hole_fallback:
        out["plot_hole_score"] = clamp_score(result.get("plot_hole_score"))
        out["plot_hole_note"] = "LLM whole-book estimate (no editorial_review.txt found to derive this from)."
    return out


# =====================================================================
# 8b. FINISHED PRODUCT NOTES — the last job, run once scoring is complete.
# Produces a plain-text summary of the finished book: the descriptive fields a
# catalogue entry generally asks for, filled with generated content where that's
# useful (description/blurb, search keywords) and left to a human where it isn't
# (categories, which vary by destination and need a person's eye).
#
# Ebook-only scope — no print fields (trim size, bleed, ISBN, print cost) since
# nothing print-specific exists in this pipeline.
#
# Low-score handling (confirmed 2026-08-11): always generate this file
# — never silently skip a book — but open with a prominent warning if the
# score is below LOW_SCORE_THRESHOLD, so a bad book doesn't quietly get the
# same-looking notes file as a good one. You still make the actual
# keep/drop call.
# =====================================================================

LOW_SCORE_THRESHOLD = 70  # below this, the book is worth a second look before
                           # it gets treated as finished.

def extract_author_name(book_dir: str) -> str:
    """local-book-generator.py never writes the pen name into outline.json or
    style_sheet.json as structured data (it's computed purely from genre via
    GENRE_PEN_NAMES at generation time) — duplicating that mapping here would
    drift out of sync if you ever edit it in one place and not the other.
    Instead, parse it straight out of manuscript_raw.txt, which is the ground
    truth for what pen name THIS specific book was actually credited to
    ('by {AUTHOR_NAME}' on line 2, and again in the copyright line)."""
    path = os.path.join(book_dir, "manuscript_raw.txt")
    if not os.path.isfile(path):
        return "(author name not found — manuscript_raw.txt missing; check Stage D output)"
    with open(path, "r", encoding="utf-8") as f:
        text = f.read()
    m = re.search(r"^Copyright\s*©?\s*\d{4}\s+(.+)$", text, re.MULTILINE)
    if m:
        return m.group(1).strip()
    m = re.search(r"^by\s+(.+)$", text, re.MULTILINE)
    if m:
        return m.group(1).strip()
    return "(author name not found in manuscript_raw.txt — check the file manually)"


def _fallback_keywords(full_text: str, character_names: list, top_n: int = 7) -> list:
    """Simple frequency-based fallback if the LLM metadata pass fails —
    same technique as editorial_agent.py's candidate-keyword extraction.
    Lower quality than genuine reader-search-intent phrases (single frequent
    words aren't how people actually search for a book), but keeps the file usable
    rather than leaving the keywords field empty."""
    stopwords = set((
        "the a an and or but if of to in on at for with as by from is are was were be been being "
        "this that these those it its his her their they he she him them we you your our i me my "
        "not no so than then there here when where what who which how why all any both each few"
    ).split())
    exclude = {n.lower() for full in character_names for n in full.split()}
    words = re.findall(r"[A-Za-z']{4,}", full_text.lower())
    counts = Counter(w for w in words if w not in stopwords and w not in exclude)
    return [word for word, _ in counts.most_common(top_n)]


def metadata_pass(style_sheet: dict, outline: dict, chapter_synopses: list, opening_excerpt: str) -> dict:
    """One combined LLM call for the two fields that genuinely need generated
    prose: the book description/blurb, and search keywords. (Categories are
    NOT LLM-generated — they are left for a human to pick.)"""
    synopses_text = "\n".join(f"Ch{n}: {s}" for n, s in chapter_synopses) or "(none available)"
    characters = "\n".join(f"- {c['name']}: {c['description']}" for c in style_sheet["characters"]) or "(none listed)"
    description = (
        f"{FICTION_FRAMING}"
        f"You are a commercial fiction copywriter writing back-cover sales copy for the "
        f"novel '{style_sheet['title']}' ({style_sheet['genre']}).\n\n"
        f"Premise: {style_sheet.get('premise', outline.get('premise', ''))}\n\n"
        f"Main characters:\n{characters}\n\n"
        f"Chapter-by-chapter synopsis:\n{synopses_text}\n\n"
        f"Opening excerpt:\n\"\"\"\n{opening_excerpt}\n\"\"\"\n\n"
        "Write:\n"
        "- book_description: a genuinely compelling, human-sounding back-cover-style book description, "
        "150-250 words. Hook first line, build stakes, end on a hook — do NOT summarize the ending. "
        "Avoid generic AI-sounding openers like 'In a world where...' or 'Little did she know' or "
        "'This is a story about...'. Write it the way an actual publisher's marketing copy reads for "
        "this genre.\n"
        "- search_keywords: exactly 7 phrases a real reader would type into a book search to find a "
        "book like this (not single generic words) — each 50 characters or fewer, genre/trope/theme "
        "specific (e.g. 'enemies to lovers space opera', 'found family sci-fi crew', not 'space' or "
        "'adventure' alone).\n\n"
        "Respond with ONLY a valid JSON object (no commentary, no markdown fences) in exactly this "
        "shape:\n"
        "{\n"
        '  "book_description": "...",\n'
        '  "search_keywords": ["...", "...", "...", "...", "...", "...", "..."]\n'
        "}"
    )
    result = run_json_pass(
        description,
        expected_output="A JSON object with book_description and exactly 7 search_keywords.",
        required_keys=("book_description", "search_keywords"),
    )
    if result is None:
        return {
            "book_description": (
                f"{style_sheet['title']} — {style_sheet.get('premise', outline.get('premise', ''))} "
                "(auto-generated fallback description — the metadata pass failed to produce one; "
                "write your own or re-run scoring_agent.py)."
            ),
            "search_keywords": None,  # filled in by caller with the frequency-based fallback
            "_failed": True,
        }
    keywords = result.get("search_keywords")
    if not isinstance(keywords, list) or not keywords:
        keywords = None  # caller substitutes the fallback
    else:
        keywords = [str(k).strip()[:50] for k in keywords][:7]
    return {
        "book_description": str(result.get("book_description", "")).strip(),
        "search_keywords": keywords,
        "_failed": False,
    }


def split_author_name(full_name: str) -> tuple:
    """The author field is usually two separate inputs (First name / Last name), with
    the form's own guidance to put a middle name or prefix in the first-name
    field and a suffix in the last-name field. Simple split: everything but the
    last word -> first name, last word -> last name. A single-word name (or an
    extraction-failure message) is passed through unsplit rather than guessed at."""
    if not full_name or full_name.startswith("("):
        return "", full_name  # extraction failed — surface the message as-is in last-name slot
    parts = full_name.strip().split()
    if len(parts) == 1:
        return "", parts[0]
    return " ".join(parts[:-1]), parts[-1]


# Genres where the "Sexually Explicit Images or Title" question is worth a
# specific double-check rather than a blind "No" — the default assumption for
# every other genre in GENRE_POOL is that the answer is No.
ADULT_CONTENT_CAUTION_GENRES = {"romantic suspense", "psychological horror"}


def build_finished_product_notes(style_sheet: dict, outline: dict, author_name: str, metadata: dict,
                                  overall_score: int, category_scores: dict, using_edited: bool,
                                  degenerate_chapters: list = None) -> str:
    lines = []
    # This file is the last thing read before a book is called finished, so a
    # coherence failure has to be unmissable HERE, above the low-score notice
    # — it's a "this book is broken" verdict, not a "consider re-editing" nudge.
    if degenerate_chapters:
        lines.append("!" * 70)
        lines.append("! STOP — THIS BOOK IS NOT FINISHED")
        lines.append(f"! {len(degenerate_chapters)} chapter(s) collapse into looping, incoherent "
                     f"prose:")
        for d in degenerate_chapters:
            lines.append(f"!   Chapter {d['chapter']}: {d['detail']}")
        lines.append("!")
        lines.append("! The metadata below is still filled in so nothing is lost, but this book")
        lines.append("! is not usable until those chapters are rerolled and it is rescored.")
        lines.append("!" * 70)
        lines.append("")
    if overall_score < LOW_SCORE_THRESHOLD:
        lines.append("!" * 70)
        lines.append(f"! LOW SCORE ({overall_score}/100) — REVIEW BEFORE CALLING THIS DONE")
        lines.append(f"! This book scored below the {LOW_SCORE_THRESHOLD} threshold. The fields below")
        lines.append("! were still generated so you have the option, but consider re-editing,")
        lines.append("! rescoring, or just dropping this one.")
        lines.append("!" * 70)
        lines.append("")

    lines.append(f"FINISHED PRODUCT NOTES — {style_sheet['title']}")
    lines.append(f"Scored version: {'edited/ (post-editor)' if using_edited else 'raw draft (editor has not run yet)'}")
    lines.append("")
    lines.append("Generated listing metadata for the finished book. These are the fields a")
    lines.append("storefront or catalogue entry generally asks for, filled in from the outline,")
    lines.append("the finished text and the score. Everything here is usable as written except")
    lines.append("'Categories', which is left blank — category trees differ between destinations")
    lines.append("and need a human eye.")
    lines.append("")
    lines.append("=" * 70)
    lines.append("LISTING METADATA")
    lines.append("=" * 70)
    lines.append("")
    lines.append("Language: English")
    lines.append("")
    lines.append(f"Book Title: {style_sheet['title']}")
    subtitle = outline.get("subtitle", "")
    lines.append(f"Subtitle (optional): {subtitle if subtitle else '(none)'}")
    lines.append("")
    lines.append("Series: (not applicable — standalone novel)")
    lines.append("")
    lines.append("Edition Number (optional): (leave blank — first edition)")
    lines.append("")
    first_name, last_name = split_author_name(author_name)
    lines.append("Primary Author or Contributor:")
    lines.append(f"  First name: {first_name}")
    lines.append(f"  Last name: {last_name}")
    lines.append("Contributors (optional): (none)")
    lines.append("")
    lines.append("Description (plain text is fine; most storefronts cap this around 4000 chars):")
    lines.append("-" * 70)
    lines.append(metadata["book_description"])
    lines.append("-" * 70)
    lines.append("")
    lines.append("Publishing rights: confirm you hold the necessary rights for this specific")
    lines.append("  book before submitting it anywhere.")
    lines.append("")
    genre_key_notes = (style_sheet.get("genre") or "").strip().lower()
    adult_note = (
        "No — but double-check this one specifically, this genre can run steamy/dark"
        if genre_key_notes in ADULT_CONTENT_CAUTION_GENRES
        else "No (default assumption for general-audience genre fiction)"
    )
    lines.append(f"Sexually explicit content: {adult_note}")
    lines.append("Reading age (optional): (leave blank — adult/general-audience fiction, not children's/YA)")
    lines.append("")
    lines.append("Categories (not guessed — pick these manually against whatever category")
    lines.append(f"  tree the destination uses; this book's genre is '{style_sheet['genre']}'):")
    lines.append("  - (pick manually)")
    lines.append("")
    lines.append("Keywords (up to 7, each 50 characters or fewer):")
    for i, kw in enumerate(metadata["search_keywords"], 1):
        lines.append(f"  {i}. {kw}")
    lines.append("")
    # AI disclosure. Most destinations ask for this for content generated by a
    # tool — which these books are, on both text and cover art.
    lines.append("-" * 70)
    lines.append("AI CONTENT DISCLOSURE — REQUIRED, DO NOT SKIP")
    lines.append("-" * 70)
    lines.append("This book is machine-generated on both counts:")
    lines.append("  * Text  — written by a local language model")
    lines.append("  * Images — cover art generated with Stable Diffusion")
    lines.append("The usual distinction is between AI-GENERATED (created by the tool — disclose)")
    lines.append("and AI-ASSISTED (a human wrote it and the tool helped edit — generally no")
    lines.append("disclosure needed). This book is the first category. Disclose it wherever the")
    lines.append("destination platform asks, and check that platform's current policy before")
    lines.append("submitting — the rules move.")
    lines.append("")
    lines.append("#" * 70)
    lines.append("# INTERNAL RECORD — for your own tracking, not part of the listing")
    lines.append("#" * 70)
    lines.append(f"# Overall score: {overall_score}/100")
    for name, score in category_scores.items():
        lines.append(f"#   {name}: {score}/100")
    lines.append(f"# Scored on: (fill in manually, or check book_score.json's file timestamp)")
    lines.append("# Use this section to track how the book actually performed against its score.")
    lines.append("#" * 70)
    return "\n".join(lines)


# =====================================================================
# 8b. TIMING INSTRUMENTATION (dashboard section 6.14 — ETA/countdown timers)
# =====================================================================


def _emit_timing(event: str, stage: str, **fields) -> None:
    """Machine-readable timing marker — see local-book-generator.py's copy
    of this same function for the full rationale. Duplicated here per this
    codebase's "each script runs standalone" convention."""
    parts = [f"event={event}", f"stage={stage}"]
    for key, value in fields.items():
        # Sanitize spaces (e.g. a chapter title in label=...) so the
        # dashboard's whitespace-tokenized parser doesn't fragment it.
        safe_value = re.sub(r"\s+", "_", str(value)) if value is not None else ""
        parts.append(f"{key}={safe_value}")
    print("[TIMING] " + " ".join(parts))


# =====================================================================
# 9. MAIN
# =====================================================================

outline = load_outline(BOOK_DIR)
style_sheet = load_style_sheet_readonly(BOOK_DIR, outline)
genre = style_sheet.get("genre", "") or outline.get("genre", "")
chapters_meta = {c["chapter_number"]: c for c in outline.get("chapters", [])}

if SCORE_DIR:
    chapter_files = discover_chapters(SCORE_DIR)
    using_edited = False  # not meaningful in --score-dir mode; scored_version is set separately below
    if not chapter_files:
        print(f"[ERROR] No chapter_NN.txt files found in --score-dir '{SCORE_DIR}'.")
        sys.exit(1)
    print(f"[Config] Scoring the REPOLISH version in '{SCORE_DIR}' ({len(chapter_files)} chapter file(s)).")
else:
    chapter_files, using_edited = choose_chapter_source(BOOK_DIR)
    if not chapter_files:
        print(f"[ERROR] No chapter_NN.txt files found in '{BOOK_DIR}' or its edited/ subfolder.")
        sys.exit(1)

    print(f"[Config] Scoring the {'EDITED (post-editor)' if using_edited else 'RAW DRAFT'} version "
          f"({len(chapter_files)} chapter file(s)).")

# Every output file (Finished Product Notes.txt, book_score.json, book_score_report.txt)
# is written here. Normally that's BOOK_DIR (with a self-contained copy mirrored into
# edited/ when using_edited). In --score-dir mode it's the score-dir itself instead, and
# the edited/ mirroring step is skipped entirely so a repolish score never clobbers
# edited/'s own real score.
OUTPUT_DIR = SCORE_DIR or BOOK_DIR

chapter_texts = {}     # ch_num -> full text
chapter_titles = {}    # ch_num -> title
chapter_word_counts = {}  # ch_num -> word count
narrative_results = {}  # ch_num -> narrative_pass() result

print(f"\n--- Scoring pass starting: {len(chapter_files)} chapter(s) ---")

for ch_num, path in chapter_files:
    ch_title, chapter_text = split_chapter_file(path)
    chapter_texts[ch_num] = chapter_text
    chapter_titles[ch_num] = ch_title
    chapter_word_counts[ch_num] = _count_words(chapter_text)

    if "NEEDS MANUAL REGENERATION" in chapter_text:
        print(f"[SKIP] Chapter {ch_num} is a placeholder from a failed writer pass — skipping scoring "
              f"until it's regenerated.")
        narrative_results[ch_num] = {
            "pacing_score": None, "hook_strength": None, "dialogue_distinctiveness": None,
            "motivation_agency": None, "emotional_intensity": None,
            "chapter_synopsis": "(skipped — placeholder chapter)", "_skipped": True,
        }
        continue

    print(f"\n--- Scoring pass: Chapter {ch_num}/{len(chapter_files)}: '{ch_title}' ---")
    _chapter_t0 = time.time()
    _emit_timing("start", "chapter_score", ch=ch_num, total=len(chapter_files), label=ch_title)
    narrative_results[ch_num] = narrative_pass(chapter_text, ch_num, ch_title, style_sheet)
    narrative_results[ch_num]["_skipped"] = False
    print(f"[Scorer] Chapter {ch_num}: pacing={narrative_results[ch_num]['pacing_score']} "
          f"hook={narrative_results[ch_num]['hook_strength']} "
          f"dialogue_voice={narrative_results[ch_num]['dialogue_distinctiveness']} "
          f"agency={narrative_results[ch_num]['motivation_agency']} "
          f"emotion={narrative_results[ch_num]['emotional_intensity']}")
    _emit_timing("end", "chapter_score", ch=ch_num, total=len(chapter_files),
                 elapsed=f"{time.time() - _chapter_t0:.1f}")

print(f"\n--- Per-chapter scoring complete: {len(chapter_files)} chapter(s) processed ---")

# ---- Whole-book synthesis (arc tracking, subplot resolution, thematic cohesion) ----

ordered_nums = sorted(chapter_texts.keys())
chapter_synopses = [
    (n, narrative_results[n]["chapter_synopsis"]) for n in ordered_nums if not narrative_results[n].get("_skipped")
]
opening_excerpt = chapter_texts[ordered_nums[0]][:1200] if ordered_nums else ""
closing_excerpt = chapter_texts[ordered_nums[-1]][-1200:] if ordered_nums else ""
# Midpoint excerpt added 2026-08-13. arc_tracking asks whether characters change
# ACROSS the book, but the pass was only ever shown the first and last 1200
# characters — two endpoints with nothing between them. A middle sample gives it
# actual evidence of the trajectory instead of forcing it to infer one from the
# synopsis line, which is a large part of why that metric never cleared 48.
_mid_idx = len(ordered_nums) // 2 if ordered_nums else 0
midpoint_excerpt = chapter_texts[ordered_nums[_mid_idx]][:1200] if ordered_nums else ""
outline_chapters = outline.get("chapters", [])

# scoring_edited=True makes this return None for the edited/repolish versions,
# which routes plot integrity through the synthesis pass's own contradiction
# check against the text actually being scored. See the function's docstring for
# why reading the raw draft's review here was wrong (it scored 0 on 7/7 books).
plot_hole_result = score_plot_holes_from_editorial_review(
    BOOK_DIR, scoring_edited=bool(using_edited or SCORE_DIR)
)
need_plot_hole_fallback = plot_hole_result is None

print("\n--- Whole-book synthesis pass (arc tracking, subplot resolution, thematic cohesion"
      + (", plot-hole fallback" if need_plot_hole_fallback else "") + ") ---")
_synthesis_t0 = time.time()
_emit_timing("start", "whole_book_synthesis")
synthesis = whole_book_synthesis(
    style_sheet, chapter_synopses, outline_chapters, opening_excerpt, closing_excerpt,
    need_plot_hole_fallback, midpoint_excerpt=midpoint_excerpt
)
_emit_timing("end", "whole_book_synthesis", elapsed=f"{time.time() - _synthesis_t0:.1f}")

if plot_hole_result is None:
    plot_hole_result = {
        "score": synthesis.get("plot_hole_score", 50),
        "continuity_issues": None, "high_severity_issues": None,
        # Source string distinguishes the two ways this path is reached, so a
        # report is self-explaining about which text the number describes.
        "source": (
            "llm_residual_check_on_scored_text (edited/repolish — the raw draft's "
            "editorial_review.txt describes defects the editor already fixed, so it is "
            "deliberately not used here)"
            if (using_edited or SCORE_DIR)
            else "llm_whole_book_fallback (no editorial_review.txt found)"
        ),
    }

# ---- Deterministic whole-manuscript metrics ----

full_manuscript_text = "\n\n".join(chapter_texts[n] for n in ordered_nums)
total_words = sum(chapter_word_counts.values())
character_names = [c["name"] for c in style_sheet["characters"]]

readability = score_readability(full_manuscript_text, (genre or "").strip().lower())
sentence_variety = score_sentence_variety(full_manuscript_text)
filler_words = score_filler_words(full_manuscript_text)
dialogue_ratio = score_dialogue_ratio(full_manuscript_text)
repetitive_starters = score_repetitive_starters(full_manuscript_text)
cliche_density = score_cliche_density(full_manuscript_text)
sensory_engagement = score_sensory_engagement(full_manuscript_text)
show_vs_tell = score_show_vs_tell(full_manuscript_text)
scene_balance = score_scene_balance(list(chapter_word_counts.values()))
emotional_curve = score_emotional_curve(
    [narrative_results[n]["emotional_intensity"] for n in ordered_nums if not narrative_results[n].get("_skipped")]
)
# The metric that actually counts toward the score — see score_pacing_variation().
pacing_variation = score_pacing_variation([chapter_texts[n] for n in ordered_nums])

# ---- Aggregate per-chapter LLM scores ----


def _avg_field(field_name: str) -> int:
    values = [narrative_results[n][field_name] for n in ordered_nums if not narrative_results[n].get("_skipped")]
    values = [v for v in values if v is not None]
    return round(sum(values) / len(values)) if values else 50


avg_pacing = _avg_field("pacing_score")
avg_hook = _avg_field("hook_strength")
avg_dialogue_voice = _avg_field("dialogue_distinctiveness")
avg_motivation = _avg_field("motivation_agency")

# ---- Roll up into the 6 categories, then overall ----


def _avg(values: list) -> int:
    values = [v for v in values if v is not None]
    return round(sum(values) / len(values)) if values else 50


category_scores = {
    "1. Technical Execution & Readability": _avg([
        readability["score"], sentence_variety["score"], filler_words["score"], dialogue_ratio["score"],
    ]),
    "2. Narrative Pacing & Structure": _avg([avg_pacing, avg_hook, scene_balance["score"]]),
    "3. Character Consistency & Voice": _avg([avg_dialogue_voice, synthesis["arc_tracking"], avg_motivation]),
    "4. Plot Integrity & World Logic": _avg([
        plot_hole_result["score"], synthesis["subplot_resolution"], show_vs_tell["score"],
    ]),
    "5. Stylistic & Sensory Polish": _avg([
        sensory_engagement["score"], cliche_density["score"], repetitive_starters["score"],
    ]),
    "6. Reader Engagement & Cohesion": _avg([synthesis["thematic_cohesion"],
                                             pacing_variation["score"]]),
}

overall_score = _avg(list(category_scores.values()))

# =====================================================================
# COHERENCE FLOOR — added 2026-08-13, applied AFTER the rubric
# =====================================================================
# Deliberately a ceiling imposed on the finished score rather than another
# metric folded into the average. A 19th metric would have been diluted to
# nothing: the rubric averages 18 sub-metrics into 6 categories and then
# averages those, so even a 0 on one new metric could only move the overall
# score a few points — and "Rashomon's Rainy Nights" proved a book can be
# entirely unreadable while scoring 80. Publishability is not an average;
# a book with a chapter of looping gibberish is not 80% publishable.
#
# Product decision 2026-08-13: hard fail. Any degenerate chapter caps the
# overall score below the 70 publish threshold, and the report names which
# chapters and why.
DEGEN_SCORE_CEILING = 45  # comfortably below LOW_SCORE_THRESHOLD (70)

print("\n--- Coherence check (repetition-collapse guard) ---")
degenerate_chapters = []
for n in ordered_nums:
    verdict = detect_degeneration(chapter_texts[n])
    if verdict["degenerate"]:
        degenerate_chapters.append({"chapter": n, **verdict})
        print(f"[!] Chapter {n}: {verdict['detail']}")
    else:
        ttr = verdict.get("worst_ttr")
        print(f"    Chapter {n}: OK"
              + (f" (lowest vocabulary richness {ttr})" if ttr is not None else ""))

if degenerate_chapters:
    pre_ceiling_score = overall_score
    overall_score = min(overall_score, DEGEN_SCORE_CEILING)
    print(f"\n[!] {len(degenerate_chapters)} of {len(ordered_nums)} chapter(s) collapse into "
          f"repetition. The rubric scored this book {pre_ceiling_score}/100, but that number "
          f"only describes measurable craft — none of the 18 sub-metrics can detect looping "
          f"prose (readability actually IMPROVES on it).")
    print(f"[!] Overall score capped at {overall_score}/100 — below the {LOW_SCORE_THRESHOLD} "
          f"publish threshold. Do not publish this book as-is; reroll the flagged chapters.")
else:
    print("    All chapters passed the coherence check.")

print(f"\n--- Scoring complete. Overall score: {overall_score}/100 ---")

# ---- Final job, run last: Finished Product Notes.txt ----

print("\n--- Generating Finished Product Notes ---")
author_name = extract_author_name(BOOK_DIR)
metadata_result = metadata_pass(style_sheet, outline, chapter_synopses, opening_excerpt)
if metadata_result["search_keywords"] is None:
    metadata_result["search_keywords"] = _fallback_keywords(full_manuscript_text, character_names)
    print("[WARN] Metadata pass didn't return usable keywords — used frequency-based fallback instead "
          "(lower quality than reader-search-intent phrases; consider re-running).")
if metadata_result.get("_failed"):
    print("[WARN] Metadata pass failed — book_description is a basic auto-generated fallback, not "
          "genuine sales copy. Consider re-running scoring_agent.py or writing the description by hand.")

finished_notes_text = build_finished_product_notes(
    style_sheet, outline, author_name, metadata_result,
    overall_score, category_scores, using_edited,
    degenerate_chapters=degenerate_chapters,
)
notes_path = os.path.join(OUTPUT_DIR, "Finished Product Notes.txt")
with open(notes_path, "w", encoding="utf-8") as f:
    f.write(finished_notes_text)
print(f"[Saved] {notes_path}")

edited_dir_check = os.path.join(BOOK_DIR, "edited")
if not SCORE_DIR and os.path.isdir(edited_dir_check):
    shutil.copyfile(notes_path, os.path.join(edited_dir_check, "Finished Product Notes.txt"))
    print(f"[Saved] {os.path.join(edited_dir_check, 'Finished Product Notes.txt')} (copy, so edited/ stays self-contained)")

if overall_score < LOW_SCORE_THRESHOLD:
    print(f"[!] Overall score {overall_score}/100 is below {LOW_SCORE_THRESHOLD} — Finished Product "
          f"Notes.txt opens with a warning. Worth a second look before publishing.")

# =====================================================================
# 10. REPORT (JSON + human-readable text)
# =====================================================================

score_data = {
    "title": style_sheet.get("title", outline.get("title", "")),
    "genre": genre,
    "scored_version": "repolish" if SCORE_DIR else ("edited" if using_edited else "raw_draft"),
    "overall_score": overall_score,
    "category_scores": category_scores,
    # Coherence guard result, recorded in the JSON so the dashboard and any
    # later tooling can tell a capped score from a genuinely low one.
    "coherence": {
        "passed": not degenerate_chapters,
        "degenerate_chapters": degenerate_chapters,
        "rubric_score_before_cap": pre_ceiling_score if degenerate_chapters else overall_score,
        "score_ceiling_applied": DEGEN_SCORE_CEILING if degenerate_chapters else None,
    },
    "sub_metrics": {
        "readability_index": readability,
        "sentence_variety": sentence_variety,
        "filler_filter_words": filler_words,
        "dialogue_exposition_ratio": dialogue_ratio,
        "chapter_progression_pacing": {"score": avg_pacing},
        "hook_strength": {"score": avg_hook},
        "scene_length_balance": scene_balance,
        "dialogue_distinctiveness": {"score": avg_dialogue_voice},
        "arc_tracking": {"score": synthesis["arc_tracking"], "note": synthesis["arc_note"]},
        "motivation_agency": {"score": avg_motivation},
        "plot_hole_score": plot_hole_result,
        "subplot_resolution": {"score": synthesis["subplot_resolution"], "note": synthesis["subplot_note"]},
        "show_vs_tell_ratio": show_vs_tell,
        "sensory_engagement": sensory_engagement,
        "cliche_density": cliche_density,
        "repetitive_starters": repetitive_starters,
        "thematic_cohesion": {"score": synthesis["thematic_cohesion"], "note": synthesis["thematic_note"]},
        "pacing_variation": pacing_variation,
        # Retired 2026-08-14, recorded but NOT scored — see score_emotional_curve().
        "emotional_resonance_retired": emotional_curve,
    },
}

score_json_path = os.path.join(OUTPUT_DIR, "book_score.json")
with open(score_json_path, "w", encoding="utf-8") as f:
    json.dump(score_data, f, indent=2)

report_lines = []
report_lines.append(f"QUALITY SCORE — {score_data['title']}")
report_lines.append(f"Genre: {genre}")
if SCORE_DIR:
    report_lines.append(f"Scored version: repolish ({SCORE_DIR})")
else:
    report_lines.append(f"Scored version: {'edited/ (post-editor)' if using_edited else 'raw draft (editor has not run yet)'}")
report_lines.append(f"Overall score: {overall_score}/100")
if degenerate_chapters:
    report_lines.append("")
    report_lines.append("!" * 70)
    report_lines.append("! DO NOT PUBLISH — REPETITION COLLAPSE DETECTED")
    report_lines.append(f"! {len(degenerate_chapters)} of {len(ordered_nums)} chapters break down "
                        f"into looping, incoherent prose.")
    report_lines.append(f"! The craft rubric below scored this book {pre_ceiling_score}/100, but "
                        f"none of its 18")
    report_lines.append("! sub-metrics can detect this failure — readability and cliche-density "
                        "scores actually")
    report_lines.append("! IMPROVE on repetitive text. The score above is capped instead.")
    report_lines.append("!")
    for d in degenerate_chapters:
        report_lines.append(f"!   Chapter {d['chapter']}: {d['detail']}")
    report_lines.append("!")
    report_lines.append("! Fix: reroll the chapters listed above, or lower their target_words in")
    report_lines.append("! outline.json — an over-long target is what triggers the spiral.")
    report_lines.append("!" * 70)
report_lines.append("")
report_lines.append("Category scores:")
for name, score in category_scores.items():
    report_lines.append(f"  {name}: {score}/100")
report_lines.append("")
report_lines.append("=" * 70)
report_lines.append("1. TECHNICAL EXECUTION & READABILITY")
report_lines.append(f"  Readability index: {readability['score']}/100 "
                     f"(Flesch-Kincaid grade {readability['grade_level']}, target {readability['target_band']})")
report_lines.append(f"  Sentence variety: {sentence_variety['score']}/100 "
                     f"(stdev {sentence_variety['stdev_words']} words)")
report_lines.append(f"  Overused/filter words: {filler_words['score']}/100 "
                     f"({filler_words['rate_per_1000_words']} per 1000 words)")
report_lines.append(f"  Dialogue-to-exposition ratio: {dialogue_ratio['score']}/100 "
                     f"({dialogue_ratio['dialogue_pct']}% dialogue)")
report_lines.append("")
report_lines.append("2. NARRATIVE PACING & STRUCTURE")
report_lines.append(f"  Chapter progression/pacing: {avg_pacing}/100 (avg across chapters)")
report_lines.append(f"  Hook strength: {avg_hook}/100 (avg across chapters)")
report_lines.append(f"  Scene length balance: {scene_balance['score']}/100 "
                     f"(chapter-length coefficient of variation {scene_balance['coefficient_of_variation']} "
                     f"— approximated from chapter length, no explicit scene-break markers in the text)")
report_lines.append("")
report_lines.append("3. CHARACTER CONSISTENCY & VOICE")
report_lines.append(f"  Dialogue distinctiveness: {avg_dialogue_voice}/100 (avg across chapters)")
report_lines.append(f"  Arc tracking: {synthesis['arc_tracking']}/100 — {synthesis['arc_note']}")
report_lines.append(f"  Motivation & agency: {avg_motivation}/100 (avg across chapters)")
report_lines.append("")
report_lines.append("4. PLOT INTEGRITY & WORLD LOGIC")
report_lines.append(f"  Contradiction/plot-hole score: {plot_hole_result['score']}/100 "
                     f"(source: {plot_hole_result['source']})")
report_lines.append(f"  Subplot resolution: {synthesis['subplot_resolution']}/100 — {synthesis['subplot_note']}")
report_lines.append(f"  Show vs. tell ratio: {show_vs_tell['score']}/100 "
                     f"({show_vs_tell['telling_instances']} telling instances, "
                     f"{show_vs_tell['rate_per_1000_words']} per 1000 words)")
report_lines.append("")
report_lines.append("5. STYLISTIC & SENSORY POLISH")
report_lines.append(f"  Sensory engagement: {sensory_engagement['score']}/100 "
                     f"(visual={sensory_engagement['counts']['visual']}, "
                     f"auditory={sensory_engagement['counts']['auditory']}, "
                     f"tactile={sensory_engagement['counts']['tactile']}, "
                     f"olfactory={sensory_engagement['counts']['olfactory']})")
report_lines.append(f"  Cliche density: {cliche_density['score']}/100 "
                     f"({cliche_density['matches']} matches, {cliche_density['rate_per_1000_words']} per 1000 words)")
report_lines.append(f"  Repetitive sentence starters: {repetitive_starters['score']}/100 "
                     f"({repetitive_starters['repeat_rate_pct']}% of consecutive sentences)")
report_lines.append("")
report_lines.append("6. READER ENGAGEMENT & COHESION")
report_lines.append(f"  Thematic cohesion: {synthesis['thematic_cohesion']}/100 — {synthesis['thematic_note']}")
report_lines.append(f"  Pacing variation: {pacing_variation['score']}/100 — "
                     f"{pacing_variation['detail']}")
report_lines.append(f"    per-chapter pace index: "
                     f"{', '.join(str(p) for p in pacing_variation['per_chapter'])}")
report_lines.append(f"    (measured from sentence length, paragraph length and dialogue share. "
                     f"Replaces the old 'emotional resonance' metric, which asked the model to "
                     f"rate each chapter's intensity and got {', '.join(str(v) for v in emotional_curve.get('ratings', []))} "
                     f"— a saturated rater that scored even deliberately quiet chapters in the 90s.)")

report_path = os.path.join(OUTPUT_DIR, "book_score_report.txt")
with open(report_path, "w", encoding="utf-8") as f:
    f.write("\n".join(report_lines))

# Mirror both the JSON and the human-readable report into edited/ when that's
# the version actually scored — same "keep edited/ self-contained" reasoning
# already applied to cover images and Finished Product Notes.txt. Only do this
# when using_edited is True: if the editor hasn't run yet (or was suggest-only)
# the score reflects the raw draft, and copying it into edited/ would falsely
# imply the edited version was what got scored. Never do this in --score-dir
# mode either — that would clobber edited/'s own real score with a repolish score.
edited_score_json_path = None
edited_report_path = None
if using_edited and not SCORE_DIR:
    edited_score_json_path = os.path.join(edited_dir_check, "book_score.json")
    with open(edited_score_json_path, "w", encoding="utf-8") as f:
        json.dump(score_data, f, indent=2)
    edited_report_path = os.path.join(edited_dir_check, "book_score_report.txt")
    with open(edited_report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))
    print(f"[Saved] {edited_score_json_path} (copy — strictly the edited-version score)")
    print(f"[Saved] {edited_report_path} (copy — strictly the edited-version score)")

print(f"\n=======================================================")
print("SCORING PASS COMPLETE")
print(f"=======================================================")
print(f"Book: {score_data['title']}")
print(f"Overall score: {overall_score}/100")
if SCORE_DIR:
    print(f"Scored version: repolish ({SCORE_DIR})")
else:
    print(f"Scored version: {'edited/ (post-editor)' if using_edited else 'raw draft (editor has not run yet)'}")
print(f"JSON saved to: {score_json_path}")
print(f"Report saved to: {report_path}")
if using_edited and not SCORE_DIR:
    print(f"JSON copy saved to: {edited_score_json_path}")
    print(f"Report copy saved to: {edited_report_path}")
print(f"Finished Product Notes saved to: {notes_path}")

# Last stage in the chain — drop the model rather than leaving it resident for
# Ollama's ~5 minute idle timeout, which in a batch run would otherwise overlap
# the NEXT book's writer model and push it onto the CPU. Added 2026-08-13.
unload_ollama_model(SCORER_MODEL_NAME)
