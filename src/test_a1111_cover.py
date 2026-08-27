"""
Quick standalone test: hits AUTOMATIC1111's txt2img API with the same
settings the book generator uses (including hires-fix, the part that was
throwing HTTP 500), so you can confirm the fix works in under a minute
instead of waiting through a full chapter-writing run.

Run it with:
    python test_a1111_cover.py

On success, saves test_cover_raw.png in the current folder so you can
open it and see the actual generated art.
"""

import base64
import time

import requests

URL = "http://127.0.0.1:7860/sdapi/v1/txt2img"

PROMPT = (
    "A lone lighthouse on a rocky cliff at sunset, dramatic orange and "
    "purple sky, ocean waves crashing below, cinematic lighting"
)
NEGATIVE_PROMPT = (
    "text, watermark, signature, letters, words, blurry, low quality, "
    "deformed, extra limbs, bad proportions"
)

payload = {
    "prompt": PROMPT,
    "negative_prompt": NEGATIVE_PROMPT,
    "steps": 30,
    "cfg_scale": 7.0,
    "width": 832,
    "height": 1216,
    "sampler_name": "DPM++ 2M Karras",
    "enable_hr": True,
    "hr_scale": 1.5,  # matches local-book-generator.py — was 2.0, which needs more
                      # VRAM than the real pipeline ever asks for and made this test
                      # fail in cases the actual run would have handled fine.
    "hr_upscaler": "R-ESRGAN 4x+",
    "denoising_strength": 0.4,
}

print(f"Sending request to {URL} ...")
print("This uses hires-fix (the part that was previously failing with HTTP 500),")
print("so it may take 30-90 seconds depending on your GPU. Please wait...\n")

start = time.time()
try:
    response = requests.post(URL, json=payload, timeout=300)
except requests.exceptions.ConnectionError as e:
    print(f"CONNECTION FAILED: {e}")
    print("\nIs AUTOMATIC1111 actually running and reachable at 127.0.0.1:7860?")
    raise SystemExit(1)

elapsed = time.time() - start
print(f"Response received in {elapsed:.1f}s — HTTP {response.status_code}")

if response.status_code != 200:
    print("\n--- FAILURE ---")
    print(f"HTTP {response.status_code}")
    print(response.text[:2000])
    raise SystemExit(1)

data = response.json()
image_bytes = base64.b64decode(data["images"][0])
with open("test_cover_raw.png", "wb") as f:
    f.write(image_bytes)

print("\n--- SUCCESS ---")
print(f"Saved test_cover_raw.png ({len(image_bytes)} bytes)")
print("Open that file to see the generated image.")
