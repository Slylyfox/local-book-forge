# The continuity seed is a style ratchet

*Investigation, 2026-08-15*

## Symptom

Readability scored 27/100 on a recent book — Flesch-Kincaid grade 15.1 against a
6–9 target for the genre — with chapters ranging from 25 to 50 words per
sentence. An earlier book ran 27 to 55.

This was puzzling because the writer prompt already contains an explicit PROSE
STYLE block asking for short-to-medium sentences and an 8th-to-10th grade
reading level. The question was not "why doesn't it follow instructions" but
"why does it follow them sometimes."

## The tell

Chapter 1 is the only chapter that receives no continuity seed.

Measured across the raw drafts of the 7 books with complete chapter sets:

| | mean words/sentence | stdev | range | n |
|---|---|---|---|---|
| Chapter 1 (unseeded) | 26.5 | 1.7 | 25.2–30.6 | 7 |
| Chapters 2+ (seeded) | 31.8 | 7.9 | 19.3–55.0 | 37 |

Seven books across seven genres, and unseeded the writer lands within a few
words of the same number every time — close to what the prompt asks for. Seeded,
the mean climbs 5 words and the spread widens 4.5×.

Chapter 1 is not written differently. It is the only chapter whose prompt ends
with an *instruction* rather than with an *example*.

## Isolating the mechanism

If the seed is the cause, then the specific text pasted into the prompt should
predict the next chapter better than the previous chapter's style in general.
That is a testable distinction, because the seed is only the last 1,500
characters — not the whole chapter.

Correlations are centred within book, so a book that simply runs long-sentenced
throughout can't manufacture the effect. Permutation test (50,000 shuffles of
the predictor within book) because n is small.

| predictor | r | p |
|---|---|---|
| previous chapter's **tail** (the 1,500 chars actually pasted in) | **+0.42** | **0.012** |
| previous chapter's **body** (its overall style) | +0.18 | 0.27 |

What gets pasted into the prompt predicts the next chapter. What doesn't get
pasted doesn't. That asymmetry is the causal signature — the model is imitating
the sample it was shown, and because the sample sits at the very end of the
prompt it outranks the abstract style instruction earlier.

Once a chapter runs long, the next one inherits it, and nothing pulls it back.
A ratchet with no restoring force. One book climbed 26.8 → 55.0 over six
chapters that way.

## Alternatives tested and rejected

All within-book centred, same corpus:

| hypothesis | r |
|---|---|
| chapter position (later chapters drift) | +0.24 |
| emotional pitch from the outline | +0.01 |
| chapter word target | −0.01 |

Emotional pitch is worth a note: the outline assigns each chapter a 0–100
intensity, and the writer prompt varies its instructions accordingly ("let it
breathe" for quiet chapters). It has essentially zero relationship to sentence
length in the output — consistent with the earlier finding that this writer
ignores style instructions the editor acts on.

## Fix

Three changes at the seed site in `local-book-generator.py`:

1. **`CONTINUITY_MAX_TAIL_SENTENCE = 32.0`.** The tail is now vetted for
   sentence length as well as for repetition collapse. Over the ceiling, it
   falls back to the previous chapter's outline summary — carrying the events
   without carrying the style. This reuses the fallback path already built for
   the collapse check.

   32.0 sits above every tail the best-scoring book produced (its worst was
   29.6) and below the ones that ran away, so the book the pipeline got most
   right is left untouched.

2. **The seed is labelled as reference material** — "for continuity of events
   only… NOT a style sample."

3. **The length constraint is restated after the seed**, so the last thing in
   the prompt is the constraint rather than the example. Until this change the
   prompt literally ended with a raw block of prose and no instruction after it,
   which is close to the canonical way of asking a model to imitate something.

Also added `average_sentence_words()`, using the same splitter and word pattern
as the scoring agent — verified identical to four decimal places across 53 real
chapters, so the writer optimises for exactly what the scorer measures.

## Verification

Replaying the rule over the real corpus: fires on 12 of 39 seeds (31%), catching
the two runaway books before the chapters where they climbed, and leaving the
best-scoring book and nine others untouched.

**This is a replay, not a simulation.** It shows where the trigger fires on the
chapters that were actually generated. It cannot show what the following
chapters would have become, because changing chapter N's seed changes chapter
N+1's text. The predicted improvement is unproven until a full run.

## Known cost

The fallback replaces prose continuity with summary continuity on roughly a
third of chapters. That is a real tradeoff — weaker scene-to-scene continuity in
exchange for stopping the ratchet. If continuity or plot-integrity scores drop
in subsequent books, the correct response is to raise the ceiling, not to
abandon the check.
