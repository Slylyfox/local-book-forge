# Revision selection: the guard that caused the defect

*Investigation, 2026-08-14 / 15*

The editorial agent's revision step has been wrong three times, in three
different ways. The sequence is worth recording because each fix was reasonable
and each one was wrong for a reason the previous one couldn't have predicted.

## Round 1: a threshold sitting on top of the distribution

The original acceptance test was `len(revised) > 0.5 * len(original)`.

Measured across one book's 15 chapters, accepted revisions came back at **51–65%
of the original** — this editor legitimately condenses. The threshold sat
exactly on top of where the distribution lives, making acceptance close to a
coin flip.

Two chapters fell just under it, were rejected three times each, and shipped
**unedited**. They were then the worst prose in the book: one had 47% of its
sentences over 40 words and contained a single 112-word sentence.

The guard meant to protect quality was the reason the two chapters that most
needed editing received none.

## Round 2: overcorrecting

Lowering the floor to 0.35 stopped the silent reverts and the metrics jumped —
readability 46 → 81 in a single run. But it also licensed the editor to cut far
harder than "tightening". On the next run the writer delivered 14,554 words,
correctly in range, and the editor shipped 8,331 — a 43% cut that dropped the
book below the word floor its own length preset had asked for.

**A threshold alone cannot separate "condensed well" from "cut too much",
because both look identical to a length check.** So the third version does not
guess at a fourth number. The floor moved to where real tightening actually
lives (10–25% reduction, i.e. 0.75) and the work of hitting it moved into the
prompt: every attempt states the required word count outright, and a short
return is retried *with the shortfall quoted back*, rather than re-sending an
identical prompt and hoping for a different sample.

A separate salvage floor at 0.35 keeps a short-but-complete revision when every
attempt fails, because the ordering that matters is: full-length real edit >
short-but-complete edit > unedited original.

That salvage floor was nearly set to 0.45. A mock test showed a stubborn model
returning 41.6% — and 41% is exactly what a subsequent real run produced. At
0.45 those chapters would have shipped unedited, reinstating the original bug
through a side door. Caught in testing, not in production.

## Round 3: the length floor fights the style directives

The most counter-intuitive result in the project. On a 9-chapter book, the
feedback retry rescued 6 chapters; 3 exhausted their attempts and fell to
salvage. Measuring the two groups:

| | readability (FK) | dialogue | avg sentence |
|---|---|---|---|
| salvaged (3 chapters) | 10.9 | 14.1% | 23.1 words |
| fully revised (6 chapters) | 13.1 | 3.8% | 28.1 words |

**The chapters the length check rejected read better on every measure than the
ones it accepted.**

The mechanism: told its revision came back short, the model complies the safe
way. It restores the original dense narration rather than rewriting it into
scenes. So "passes the length check" turned out to be nearly uncorrelated with
"reads well" — and the loop was returning on exactly that evidence, keeping the
first attempt that fit and discarding the rest.

### Fix

**Score every attempt, keep the best.** `prose_quality()` is a deterministic
0–100 composite of the three signals a revision actually moves and the scorer
later grades: Flesch-Kincaid against the genre band (35%), dialogue share
against 25–40% (35%), and sentence discipline — average length plus share of
sentences over 35 words (30%). Plot, coherence and continuity are deliberately
excluded; revision isn't supposed to change them, so grading attempts on them
would mostly measure noise.

Salvage ranking changed the same way. Length is now the eligibility test; among
attempts that clear it, the better-reading one wins. Ranking salvage by raw
length was selecting whichever attempt had restored the most original narration
— precisely the attempt the table above says to avoid.

**Bounded cost.** A good first attempt still costs exactly one generation.
`MAX_QUALITY_RETRIES = 1` caps purely quality-motivated extra generations at one
per chapter; retries caused by an attempt actually failing are unchanged. Worst
case per chapter is identical to before.

**Reworded the feedback.** The old text — "you cut roughly N words of material
that needs to stay" — is what produced the restoration behaviour. It now asks
the model to reach the word count by *dramatising*: find passages that report an
exchange between characters and play them out as dialogue with action beats.
Explicitly forbids pasting the original narration back in.

## Round 3.5: measuring the wrong text

The first live run of the above exposed a bug in it.

The editor logged 2–10% dialogue on every chapter. The scorer reported 24.6% for
the same book — using an identical regex.

The writer emits **mixed quote characters**. One chapter came out of the model
with 89 straight quotes, 12 curly-open and 23 curly-close. With straight quotes
a single character serves as both delimiters, so one unbalanced quote inverts
the alternation and every subsequent match captures *the narration between* the
dialogue rather than the dialogue itself:

| chapter 3, same text, same regex | dialogue |
|---|---|
| after the pipeline's typography normalisation | 39.5% |
| as the model returned it | 4.8% |

The scorer measures post-normalisation text. The editor was measuring pre-.

Consequence: the dialogue gate could never be cleared, so every revised chapter
burned its extra generation for nothing, and selection ran partly on a corrupted
signal. Fixed by normalising through the pipeline's own shared typography
function before measuring — verified that editor and scorer now agree to the decimal on
all nine chapters of the run that exposed it, where before they disagreed on
seven.

## The generalisable lesson

Two components measuring "the same thing" on slightly different inputs is a bug
class, not an incident. It appeared twice in this project:

- the collapse detector was **phase-dependent** — the writer saw the chapter
  body and scored it clean, the scorer saw the same text with a title line
  prepended and scored it degenerate. A four-word offset flipped the verdict.
- the dialogue measure was **normalisation-dependent**, above.

Both were found by noticing two components disagreeing about the same book,
not by a test. Both fixes were the same shape: make the measurement invariant to
the difference, and verify the two components agree on real data.
