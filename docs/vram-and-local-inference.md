# VRAM contention and silent CPU offload

*Reference hardware: RTX 4060 Laptop GPU, 8GB VRAM*

## Three symptoms, one cause

Reported separately, and they looked unrelated:

1. the machine crashed partway through a 5-book batch
2. the editorial pass hung indefinitely
3. a "rogue" Ollama process consumed ~50% of CPU and RAM, and respawned after
   being killed from Task Manager

All three were the same bug. Nothing unloaded AUTOMATIC1111's SDXL checkpoint
before the language model loaded. With both resident in 8GB, **Ollama silently
fell back to CPU inference rather than erroring.**

That is the failure mode worth internalising. It does not announce itself. There
is no error, no warning, no degraded-mode log line. Inference just becomes 5–10×
slower and starts consuming system RAM and every core — which is what the "rogue
process" was. It respawned after being killed because Ollama runs under a
supervisor that restarts it.

## Why it isn't obvious

Two things conspire:

- **Ollama holds a model in VRAM for ~5 minutes after its last call** by default
  (`keep_alive`). A stage that has finished can still be occupying the card when
  the next stage starts.
- **Failure is graceful by design.** Falling back to CPU is the right behaviour
  for a general-purpose server. It is the wrong behaviour for a pipeline where a
  chapter takes 3 minutes on GPU and 20 on CPU, because the pipeline has no way
  to know which one it got.

A related dead end worth recording: switching Windows' per-app "Graphics
performance preference" to the integrated GPU does nothing here. That setting
affects DirectX/OpenGL apps. PyTorch/CUDA — which both Ollama and A1111 use —
targets the NVIDIA device directly, and Intel UHD cannot run CUDA at all. There
was never a way to move this workload off the discrete card.

## Fix

Explicit unload and verified-free wait at every stage transition:

```
unload_a1111_checkpoint()   # POST to A1111's API
unload_ollama_model(name)   # POST keep_alive: 0
wait_for_free_vram(mb)      # poll nvidia-smi until it actually frees
```

The polling step is the one that matters. Issuing the unload and proceeding
immediately does not work — the VRAM takes time to release, and the next stage
will load into whatever is left.

### Measured result

| | before | after |
|---|---|---|
| chapter write time | 128–899s, erratic | 108–183s, tight |
| full pipeline | 45–63 min | 33.8 min |

The variance is the more diagnostic number. Erratic timings that swing 7× are
the signature of intermittent CPU fallback; tight timings mean the GPU path is
being taken consistently.

## A second-order lesson

An early version of the VRAM warning cried wolf. It measured free VRAM before
the writing stage and reported ~2,670MB free, which looked alarming — but the
outline model was still resident from the previous stage. The warning was
correct about the number and wrong about the conclusion.

Fixed by unloading the outline model *before* taking the reading. Diagnostics
that measure at the wrong moment produce false alarms that train you to ignore
them, which is worse than having no diagnostic.

## Cover generation fallback ladder

The same constraint shapes image generation. Each cover seed runs a three-tier
ladder:

1. hi-res fix at the target size
2. plain base resolution
3. reduced resolution

A1111's hi-res second pass was the specific thing failing (`Need: 5.6GB free,
Have: 4.6GB free`), and with Ollama still loaded even the base generation
returned HTTP 500. Resolution strategy is also selected by querying A1111 for
the *actually loaded* checkpoint rather than assuming SDXL — SDXL and SD1.5 want
different native buckets, and hardcoding either one breaks when the other is
loaded.
