"""
Local Book Forge — Project Cleanup / Fresh Start
------------------------------------------------
Clears out generated books, logs, title history and timing history so the
project can start clean before a real publishing run, while preserving the
handful of things that are genuinely worth keeping.

WHY THIS SCRIPT EXISTS RATHER THAN A HANDFUL OF DELETE COMMANDS
Two of the twenty books in output_books/ are not disposable output — they are
the calibration data for the repetition-collapse detector added 2026-08-13/14
(see detect_degeneration() in local-book-generator.py / scoring_agent.py). Those
thresholds were fitted against 5 known-bad chapters from one book and 14
known-good chapters from six others. If the detector ever needs retuning, that
labelled set is how it gets done, and it cannot be regenerated on demand — a
model that has been fixed no longer produces the failure on request.

So this archives the CHAPTER TEXT of those books (a few hundred KB) rather than
the whole folders (~240MB of cover PNGs that carry no diagnostic value), then
clears everything else.

SAFETY
Runs as a DRY RUN by default: it prints exactly what it would archive and
delete, and changes nothing. Nothing is touched until you pass --apply.

    python project_cleanup.py                 # show the plan, change nothing
    python project_cleanup.py --apply         # actually do it

Everything deleted is first copied into _archive/ if it has any value, and
title_history.json / job_timing_history.json are backed up there before being
reset, so an --apply run is recoverable from within the project folder itself.
"""

import argparse
import json
import os
import shutil
import sys
from datetime import datetime

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
    except AttributeError:
        pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_ROOT = os.path.join(SCRIPT_DIR, "output_books")
LOGS_DIR = os.path.join(SCRIPT_DIR, "logs")
ARCHIVE_ROOT = os.path.join(SCRIPT_DIR, "_archive")

# ---------------------------------------------------------------------
# Books whose chapter text is the detector's calibration set. Keeping the
# labels here rather than in a loose note means a future retune knows which
# side of the line each book was on without having to re-read all of them.
# ---------------------------------------------------------------------
REFERENCE_BOOKS = {
    # The failure case. Every chapter collapses in its second half; chapter 2's
    # back half is 12.4% the single word "hell", type-token ratio 0.055. This
    # book scored 80/100 before the coherence guard existed, which is the whole
    # reason the guard exists.
    "rashomon-s-rainy-nights-20260813-143556": "known-bad — all 5 chapters collapse",
    # The healthy controls. Worst-window type-token ratio 0.365-0.505 across
    # these, versus 0.245-0.34 for the book above — that gap is where the 0.36
    # floor sits.
    "murder-at-willowbrook-manor-20260811-220813": "known-good control",
    "mystery-of-the-golden-heirloom-20260812-113753": "known-good control",
    "rainy-night-redemption-20260812-211832": "known-good control",
    "the-aurora-initiative-20260810-191957": "known-good control",
    "raining-ashes-20260811-214039": "known-good control",
    "stardust-rebellion-20260811-151041": "known-good control",
    "echoes-of-eternity-part-3-20260812-192858": "known-good control",
}

# Only these are worth archiving from a reference book. Cover PNGs are ~5MB
# each and tell us nothing about prose collapse.
REFERENCE_KEEP_SUFFIXES = (".txt", ".json")
REFERENCE_SKIP_NAMES = {"Finished Product Notes.txt"}

# Logs worth keeping: the three runs that diagnosed the bugs. Everything else
# is noise, and there is a lot of it.
KEEP_LOG_SUBSTRINGS = (
    "20260813-143527_Full_Pipeline",   # the "scored 80 but unpublishable" run
    "20260813-171118_Full_Pipeline",   # the chapter-5 collapse-reroll stall
)

STATE_FILES_TO_RESET = {
    # filename -> what an empty version must contain.
    #
    # These shapes are NOT interchangeable and getting one wrong crashes the next
    # run. title_history.json is an OBJECT keyed by normalised title, not an
    # array — dedupe_title() calls history.get(key) on it, so an empty "[]" here
    # produces "AttributeError: 'list' object has no attribute 'get'" the moment
    # the outline stage finishes. That is exactly what an earlier version of this
    # script shipped, and it took out a full run. Both shapes are now verified
    # against their consumers rather than assumed.
    "title_history.json": "{}",
    "job_timing_history.json": json.dumps({"stage_durations": {}, "job_durations": {}}, indent=2),
}


def human(n_bytes):
    for unit in ("B", "KB", "MB", "GB"):
        if n_bytes < 1024 or unit == "GB":
            return f"{n_bytes:.1f} {unit}" if unit != "B" else f"{n_bytes} B"
        n_bytes /= 1024


def dir_size(path):
    total = 0
    for root, _dirs, files in os.walk(path):
        for name in files:
            try:
                total += os.path.getsize(os.path.join(root, name))
            except OSError:
                pass
    return total


def main():
    parser = argparse.ArgumentParser(
        description="Clear generated books/logs/history for a fresh start, preserving the "
                    "coherence-detector reference set."
    )
    parser.add_argument("--apply", action="store_true",
                        help="Actually make the changes. Without this flag nothing is touched.")
    parser.add_argument("--keep-logs", action="store_true",
                        help="Leave logs/ alone entirely.")
    args = parser.parse_args()
    apply = args.apply

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    archive_dir = os.path.join(ARCHIVE_ROOT, f"cleanup-{stamp}")
    reference_dir = os.path.join(ARCHIVE_ROOT, "detector-reference-books")

    print("=" * 74)
    print("LOCAL BOOK FORGE — PROJECT CLEANUP" + ("" if apply else "   [DRY RUN — nothing will change]"))
    print("=" * 74)
    print(f"Project folder: {SCRIPT_DIR}")
    print()

    freed = 0
    plan_delete_books, plan_keep_books, missing_refs = [], [], []

    # ---- 1. output_books ----
    if os.path.isdir(OUTPUT_ROOT):
        for name in sorted(os.listdir(OUTPUT_ROOT)):
            full = os.path.join(OUTPUT_ROOT, name)
            if not os.path.isdir(full) or name == "deleted_book_backups":
                continue
            size = dir_size(full)
            if name in REFERENCE_BOOKS:
                plan_keep_books.append((name, size))
            else:
                plan_delete_books.append((name, size))
                freed += size
        for ref in REFERENCE_BOOKS:
            if not os.path.isdir(os.path.join(OUTPUT_ROOT, ref)):
                missing_refs.append(ref)

    print("-- output_books --")
    if plan_keep_books:
        print(f"  ARCHIVE chapter text of {len(plan_keep_books)} reference book(s) "
              f"-> _archive/detector-reference-books/, then remove the folders:")
        for name, size in plan_keep_books:
            print(f"    {name}")
            print(f"        ({REFERENCE_BOOKS[name]}, {human(size)} on disk, only .txt/.json kept)")
    if missing_refs:
        print(f"  NOTE: {len(missing_refs)} reference book(s) already gone — nothing to archive for:")
        for name in missing_refs:
            print(f"    {name}")
    if plan_delete_books:
        print(f"  DELETE {len(plan_delete_books)} generated book(s), {human(sum(s for _, s in plan_delete_books))}:")
        for name, size in plan_delete_books:
            print(f"    {name}  ({human(size)})")
    if not plan_keep_books and not plan_delete_books:
        print("  (already empty)")
    print()

    # ---- 2. logs ----
    plan_keep_logs, plan_delete_logs = [], []
    if os.path.isdir(LOGS_DIR) and not args.keep_logs:
        for name in sorted(os.listdir(LOGS_DIR)):
            full = os.path.join(LOGS_DIR, name)
            if not os.path.isfile(full):
                continue
            size = os.path.getsize(full)
            if any(k in name for k in KEEP_LOG_SUBSTRINGS):
                plan_keep_logs.append((name, size))
            else:
                plan_delete_logs.append((name, size))
                freed += size

    print("-- logs --")
    if args.keep_logs:
        print("  (skipped, --keep-logs)")
    else:
        if plan_keep_logs:
            print(f"  ARCHIVE {len(plan_keep_logs)} diagnostic log(s) -> _archive/cleanup-{stamp}/logs/:")
            for name, size in plan_keep_logs:
                print(f"    {name}  ({human(size)})")
        if plan_delete_logs:
            print(f"  DELETE {len(plan_delete_logs)} other log(s), {human(sum(s for _, s in plan_delete_logs))}")
        if not plan_keep_logs and not plan_delete_logs:
            print("  (already empty)")
    print()

    # ---- 3. state files ----
    print("-- state files (backed up, then reset to empty) --")
    state_present = []
    for name in STATE_FILES_TO_RESET:
        full = os.path.join(SCRIPT_DIR, name)
        if os.path.isfile(full):
            state_present.append(name)
            detail = ""
            if name == "title_history.json":
                try:
                    with open(full, "r", encoding="utf-8") as f:
                        detail = f"  ({len(json.load(f))} tracked titles)"
                except Exception:
                    pass
            print(f"  {name}{detail}")
        else:
            print(f"  {name}  (not present — will be created empty)")
    print("  Timing history is reset deliberately: every sample in it came from runs that were")
    print("  CPU-bound or spent minutes on discarded chapters, so keeping it would make the")
    print("  dashboard's ETAs wrong for a healthy pipeline.")
    print()

    # ---- 4. junk ----
    junk = []
    for root, dirs, files in os.walk(SCRIPT_DIR):
        if "_archive" in root:
            continue
        for d in list(dirs):
            if d == "__pycache__":
                p = os.path.join(root, d)
                junk.append(p)
                freed += dir_size(p)
                dirs.remove(d)
        for fn in files:
            if fn.endswith((".pyc", ".pyo")) or fn.startswith("~$"):
                p = os.path.join(root, fn)
                junk.append(p)
                try:
                    freed += os.path.getsize(p)
                except OSError:
                    pass
    print("-- caches / temp --")
    print(f"  DELETE {len(junk)} item(s) (__pycache__, .pyc, Office lock files)" if junk
          else "  (nothing to clean)")
    print()

    print("=" * 74)
    print(f"Approximate space freed: {human(freed)}")
    print("=" * 74)

    if not apply:
        print()
        print("DRY RUN — nothing above has been done.")
        print("Re-run with --apply to carry it out:")
        print(f"    python \"{os.path.join(SCRIPT_DIR, 'project_cleanup.py')}\" --apply")
        return 0

    # =================================================================
    # Execute
    # =================================================================
    print()
    print("Applying...")
    os.makedirs(archive_dir, exist_ok=True)

    # Reference books: copy chapter text out, then remove the folder.
    for name, _size in plan_keep_books:
        src = os.path.join(OUTPUT_ROOT, name)
        dst = os.path.join(reference_dir, name)
        os.makedirs(dst, exist_ok=True)
        copied = 0
        for root, _dirs, files in os.walk(src):
            rel = os.path.relpath(root, src)
            for fn in files:
                if not fn.endswith(REFERENCE_KEEP_SUFFIXES) or fn in REFERENCE_SKIP_NAMES:
                    continue
                target_dir = dst if rel == "." else os.path.join(dst, rel)
                os.makedirs(target_dir, exist_ok=True)
                shutil.copy2(os.path.join(root, fn), os.path.join(target_dir, fn))
                copied += 1
        print(f"  archived {copied} text file(s) from {name}")
        shutil.rmtree(src, ignore_errors=True)
        print(f"  removed  {name}")

    if plan_keep_books or os.path.isdir(reference_dir):
        readme = os.path.join(reference_dir, "README.txt")
        with open(readme, "w", encoding="utf-8") as f:
            f.write(
                "COHERENCE-DETECTOR REFERENCE SET\n"
                "================================\n\n"
                "Chapter text from books used to calibrate detect_degeneration() in\n"
                "local-book-generator.py and scoring_agent.py. Do not delete: this is the\n"
                "labelled data the thresholds were fitted against, and it cannot be\n"
                "regenerated once the writer stops producing the failure on demand.\n\n"
                "Measured with a sliding 400-word window:\n"
                "  known-bad book  : worst-window type-token ratio 0.245 - 0.34\n"
                "  known-good books: worst-window type-token ratio 0.365 - 0.505\n"
                "The DEGEN_TTR_FLOOR of 0.36 sits in that gap. At those settings the\n"
                "detector flagged 5/5 bad chapters and 0/14 good ones.\n\n"
                "Books:\n"
            )
            for name, label in REFERENCE_BOOKS.items():
                f.write(f"  {name}\n      {label}\n")
        print(f"  wrote {os.path.relpath(readme, SCRIPT_DIR)}")

    # Non-reference books: delete outright.
    for name, _size in plan_delete_books:
        shutil.rmtree(os.path.join(OUTPUT_ROOT, name), ignore_errors=True)
    if plan_delete_books:
        print(f"  deleted {len(plan_delete_books)} generated book folder(s)")

    # Old delete-backups from the Library panel are pure ballast now.
    backups = os.path.join(OUTPUT_ROOT, "deleted_book_backups")
    if os.path.isdir(backups):
        shutil.rmtree(backups, ignore_errors=True)
        print("  deleted output_books/deleted_book_backups/")
    os.makedirs(OUTPUT_ROOT, exist_ok=True)

    # Logs.
    if not args.keep_logs and os.path.isdir(LOGS_DIR):
        if plan_keep_logs:
            log_archive = os.path.join(archive_dir, "logs")
            os.makedirs(log_archive, exist_ok=True)
            for name, _size in plan_keep_logs:
                shutil.copy2(os.path.join(LOGS_DIR, name), os.path.join(log_archive, name))
            print(f"  archived {len(plan_keep_logs)} diagnostic log(s)")
        for name, _size in plan_delete_logs + plan_keep_logs:
            try:
                os.remove(os.path.join(LOGS_DIR, name))
            except OSError:
                pass
        print(f"  cleared logs/ ({len(plan_delete_logs) + len(plan_keep_logs)} file(s))")

    # State files: back up, then reset.
    for name, empty in STATE_FILES_TO_RESET.items():
        full = os.path.join(SCRIPT_DIR, name)
        if os.path.isfile(full):
            shutil.copy2(full, os.path.join(archive_dir, name))
        with open(full, "w", encoding="utf-8") as f:
            f.write(empty)
        print(f"  reset {name} (backup in {os.path.relpath(archive_dir, SCRIPT_DIR)})")

    # Junk.
    for p in junk:
        if os.path.isdir(p):
            shutil.rmtree(p, ignore_errors=True)
        else:
            try:
                os.remove(p)
            except OSError:
                pass
    if junk:
        print(f"  deleted {len(junk)} cache/temp item(s)")

    print()
    print("=" * 74)
    print("CLEANUP COMPLETE")
    print("=" * 74)
    print(f"Kept, and safe to leave alone:  {os.path.relpath(reference_dir, SCRIPT_DIR)}")
    print(f"This run's backups:             {os.path.relpath(archive_dir, SCRIPT_DIR)}")
    print("output_books/ is empty, title history is cleared, ETA history will rebuild")
    print("from your next runs.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
