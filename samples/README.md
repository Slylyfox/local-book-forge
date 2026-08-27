# Sample output

One complete run of the pipeline, kept as a reference artifact so the output
format is inspectable without running the system.

**`rainy-night-requiem/`** — noir private investigator, 19,056 words, scored
82/100. Generated end to end on local hardware in roughly 34 minutes.

| file | produced by | what it is |
|---|---|---|
| `outline.json` | Stage A | the generated outline — chapter beats, characters, per-chapter word targets and emotional pitch |
| `chapter_01_raw.txt` | Stage B | the writer's first-pass draft |
| `chapter_01_edited.txt` | editorial agent | the same chapter after review and revision |
| `editorial_review.txt` | editorial agent | every issue found, by category and severity, with the fix applied |
| `style_sheet.json` | editorial agent | accumulated continuity facts, characters and style terms |
| `book_score.json` / `book_score_report.txt` | scoring agent | all 18 sub-metrics, the six category rollups, and the derived length band |
| `finished_product_notes.txt` | scoring agent | generated listing metadata — blurb, keywords, categories, AI-content disclosure |
| `cover.jpg` | Stage C | SDXL cover art with Pillow title/author overlay (downscaled here from the 2.5MB print original) |
| `model_info.json` | Stage B | which local models produced this book |

## Why only chapter 1

Full manuscripts are generated output, not source, and a complete book is
several hundred kilobytes of prose that nobody is going to read in a diff — so
the repository keeps one chapter rather than all nine. One chapter in both raw
and edited form shows what the writer produces and what the editorial pass does
to it, which is the part that's actually interesting here.

Comparing the two files is the fastest way to see the editorial agent working:
the raw draft runs long-sentenced and narration-heavy, and the edited version
breaks the sentences up and converts reported speech into played-out dialogue.

## Reading the score

The report is deliberately verbose about *why* each number came out as it did,
including the calibration history for metrics that were wrong at some point.
`Reader Engagement & Cohesion: 53/100` on this book is a real weakness the
scorer caught — thematic cohesion and pacing variation both scored low — and it
is visible in the report rather than averaged away.
