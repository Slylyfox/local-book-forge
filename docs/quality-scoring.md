# Quality scoring: calibrating a rubric against real output

The scoring agent produces 18 sub-metrics across 6 categories, plus a derived
length band. Most are computed rather than asked of a model, and the ones that
were asked have mostly been retired.

This document is mainly a record of the metrics that were *wrong*, because
every one of them was wrong in a way that looked plausible until it was
measured against real books.

## The failure that motivated everything else

A book scored **80/100** while being unreadable — one word made up 12.4% of the
text, type-token ratio 0.055 over a 400-word window.

None of the 18 sub-metrics caught it. Readability *improves* on looping prose;
short repeated clauses produce excellent Flesch-Kincaid grades. Sentence
variety, cliché density and show-vs-tell all read as normal or better.

The lesson is not "add a repetition metric." It is that **a rubric of averaged
quality signals cannot represent catastrophic failure**, because averaging is
precisely the operation that hides it. So the collapse detector was added as a
**score ceiling**, not as a 19th averaged metric: a book that trips it caps at
45 regardless of everything else.

### The detector

Sliding-window type-token ratio plus a single-word frequency ceiling, requiring
consecutive bad windows to fire. Calibrated against 19 real chapters: flags 5/5
known-bad, 0/14 known-good.

Later hardened after a phase-dependence bug — the writer scanned the chapter
body and got `longest_run=1` (clean); the scorer scanned the same text with a
title line prepended and got `longest_run=2` (degenerate). A four-word offset
changed the verdict because the sliding window landed differently. Now scanned
at four independent phase offsets and the worst result taken.

## Metrics that were miscalibrated

### Plot-hole score: reading the wrong file

Scored 0/100 on 7 of 7 books. It was reading the raw draft's
`editorial_review.txt` while scoring the `edited/` version — penalising the book
for defects the editor had already fixed.

Worse than a wrong number: it **inverted the incentive**. The harder the editor
worked, the more issues appeared in the review, the lower the edited book
scored. Fixed by passing `scoring_edited` and using a residual check on the
scored text instead.

### Sentence variety: a band set for prose in the abstract

The original healthy band was a standard deviation of 4–10 words. Every real
book measured 15.2–22.0, scoring 4–59 — i.e. the metric flagged all seven books
as defective, which is the signature of a miscalibrated threshold rather than
seven bad books.

The band was too low because it ignored what a manuscript actually contains.
Dialogue exchanges ("Don't." / "Why not?") sit at 1–3 words while descriptive
narration runs 25–40. Mixing both in one population produces a standard
deviation well above 10 — in commercial genre fiction, typically 12–20. High
variance there is the deliberate rhythm the style pass explicitly asks for, so
the metric was scoring against the thing the rest of the pipeline works to
produce. Rebanded to 6–18.

### Whole-book synthesis: grading the synopsis

The LLM synthesis metrics were reading the outline and the book's summary rather
than its prose, so they were effectively grading the synopsis. Fixed by feeding
a midpoint excerpt of actual chapter text, adding calibration anchors, and an
explicit rule requiring evidence from the text for any claim.

### Emotional resonance: a saturated rater

An LLM was asked to rate each chapter's emotional intensity 0–100. On one book
it returned 60, 62, 98, 80, 85, 87, 85, 99, 80 — including 90s for chapters the
outline had deliberately planned as quiet. It correlated near zero with the
outline's own intent.

Retired and replaced with a **pacing index** computed from sentence length,
paragraph length and dialogue share. On the same book the computed index
correlates **+0.65** with the outline's planned emotional pitch, where the LLM
rater managed roughly zero.

This is the clearest case for the general principle: *if you can compute it,
don't ask a model to rate it.* The model isn't lying, it's just an unreliable
instrument on a subjective 0–100 scale with no anchors, and it saturates high.

## Length banding

The scorer also classifies the finished book into a length band, which is what
downstream packaging keys off:

| band | words |
|---|---|
| Full-length | 32,000+ |
| Standard | 20,000–32,000 |
| Short | under 20,000 |

One subtlety here took a real miss to find: **the editorial pass shrinks a
book, so the *writer's* word target has to account for it.** Two books landed
one band lower than intended, by 944 and 1,687 words, because the preset floors
were set against pre-edit length. `EDITOR_RETENTION = 0.87` now inflates the
writer's budget so the post-edit book lands in the band that was actually
requested.

That constant is measured, not guessed — it is the observed retention ratio of
the editorial pass across real runs, and it is commented with the runs it came
from.

## What the rubric still can't do

It measures prose mechanics well and story quality poorly. Thematic cohesion,
subplot resolution and arc tracking are LLM-judged and are the weakest, noisiest
metrics in the set — consistently the lowest scores, and the ones least likely
to mean what they say.

That is a known limitation rather than a solved problem. The honest framing:
this rubric reliably catches a book that is *badly written* and only weakly
detects a book that is *boring*.
