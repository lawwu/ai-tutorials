# Download & Transcribe YouTube Videos (Locally)

A step-by-step guide to download videos (or audio) from YouTube and transcribe
them to text — entirely on your own machine, with no API keys and no upload.

Everything here runs offline once the tools are installed. Tested on macOS
(Apple Silicon).

---

## What you need

| Tool | Purpose | Install |
|------|---------|---------|
| `yt-dlp` | Download from YouTube | `pip install -U yt-dlp` |
| `ffmpeg` | Convert/extract audio | `brew install ffmpeg` |
| `deno` | Runs YouTube's JS challenge solver | `brew install deno` |
| `faster-whisper` | Local speech-to-text | `pip install faster-whisper` |

Check what you already have:

```bash
yt-dlp --version
ffmpeg -version | head -1
deno --version | head -1
pip show faster-whisper | head -2
```

---

## Part 1 — Download from YouTube

### Keep yt-dlp current (important)

YouTube changes things constantly. If `yt-dlp` is more than a few weeks old,
downloads start failing. Always update first:

```bash
pip install -U yt-dlp
```

### Download the most recent video as MP3 (audio only)

Audio-only is much smaller and ideal for talks, sermons, podcasts, lectures.

```bash
yt-dlp --remote-components ejs:github \
  --playlist-end 1 \
  -x --audio-format mp3 --audio-quality 0 \
  --embed-thumbnail --add-metadata \
  -o "downloads/%(title)s.%(ext)s" \
  "https://www.youtube.com/channel/CHANNEL_ID_HERE"
```

What the flags do:

- `--remote-components ejs:github` — **required as of mid-2026.** YouTube now
  requires solving a JavaScript challenge. The solver bundled with yt-dlp is
  usually outdated, so this fetches the current one from GitHub (needs `deno`).
  If downloads fail with *"This video is not available"* or *"n challenge
  solving failed"*, this flag is the fix.
- `--playlist-end 1` — how many videos, newest first. Change to `25`, `50`, etc.
  Remove it to download the **entire** channel.
- `-x --audio-format mp3` — extract audio, convert to MP3.
- `--audio-quality 0` — best audio quality.
- `--embed-thumbnail --add-metadata` — embed cover art + title/artist tags.
- `-o "downloads/%(title)s.%(ext)s"` — save into a `downloads/` folder, named by
  the video title.

### Variations

```bash
# Most recent 25 videos as MP3
yt-dlp --remote-components ejs:github --playlist-end 25 \
  -x --audio-format mp3 --audio-quality 0 --embed-thumbnail --add-metadata \
  -o "downloads/%(title)s.%(ext)s" "CHANNEL_URL"

# Full video, best quality (MP4)
yt-dlp --remote-components ejs:github --playlist-end 1 \
  -f "bv*+ba/b" --merge-output-format mp4 \
  -o "downloads/%(title)s.%(ext)s" "CHANNEL_URL"

# Video capped at 720p (smaller files)
yt-dlp --remote-components ejs:github --playlist-end 1 \
  -f "bv*[height<=720]+ba/b[height<=720]" --merge-output-format mp4 \
  -o "downloads/%(title)s.%(ext)s" "CHANNEL_URL"

# A single video by URL (instead of a channel)
yt-dlp --remote-components ejs:github -x --audio-format mp3 \
  -o "downloads/%(title)s.%(ext)s" "https://www.youtube.com/watch?v=VIDEO_ID"
```

### List what's on a channel before downloading

```bash
yt-dlp --flat-playlist \
  --print "%(title)s | %(duration_string)s | %(id)s" \
  "CHANNEL_URL"

# Just count the videos
yt-dlp --flat-playlist --print "%(id)s" "CHANNEL_URL" | wc -l
```

---

## Part 2 — Transcribe to text

This uses [`faster-whisper`](https://github.com/SYSTRAN/faster-whisper), which
runs OpenAI's Whisper model locally — no API, no cost, nothing leaves your
machine. It also accepts video files directly (`.mp4`, `.webm`), extracting the
audio automatically, so you don't have to convert to MP3 first.

### The script

Save this as `transcribe.py`:

```python
#!/usr/bin/env python
"""Transcribe an audio/video file locally with faster-whisper.

Usage: python transcribe.py "path/to/file.mp3" [model]
  model defaults to medium.en (try small.en for speed, large-v3 for accuracy)
"""
import sys
from pathlib import Path
from faster_whisper import WhisperModel

src = Path(sys.argv[1])
model_name = sys.argv[2] if len(sys.argv) > 2 else "medium.en"

# int8 runs efficiently on Apple Silicon CPU
model = WhisperModel(model_name, device="cpu", compute_type="int8")

segments, info = model.transcribe(str(src), beam_size=5)
print(f"Detected language: {info.language} | duration: {info.duration:.0f}s", file=sys.stderr)


def ts(seconds: float, sep: str) -> str:
    h, rem = divmod(int(seconds), 3600)
    m, s = divmod(rem, 60)
    ms = int((seconds - int(seconds)) * 1000)
    return f"{h:02d}:{m:02d}:{s:02d}{sep}{ms:03d}"


stem = src.with_suffix("")
txt_f = open(f"{stem}.txt", "w")
srt_f = open(f"{stem}.srt", "w")
vtt_f = open(f"{stem}.vtt", "w")
vtt_f.write("WEBVTT\n\n")

for i, seg in enumerate(segments, 1):
    text = seg.text.strip()
    txt_f.write(text + "\n")
    srt_f.write(f"{i}\n{ts(seg.start, ',')} --> {ts(seg.end, ',')}\n{text}\n\n")
    vtt_f.write(f"{ts(seg.start, '.')} --> {ts(seg.end, '.')}\n{text}\n\n")
    print(f"[{ts(seg.start, '.')}] {text}", file=sys.stderr)  # live progress

for f in (txt_f, srt_f, vtt_f):
    f.close()

print(f"\nWrote: {stem}.txt / .srt / .vtt", file=sys.stderr)
```

### Run it

```bash
python transcribe.py "downloads/your-file.mp3"            # default: medium.en
python transcribe.py "downloads/your-file.mp3" small.en   # ~3x faster, less accurate
python transcribe.py "downloads/your-file.mp3" large-v3   # most accurate, slowest
```

You get three files next to the source, named to match it:

- **`.txt`** — plain transcript
- **`.srt`** — subtitles (for video players or uploading to YouTube)
- **`.vtt`** — timestamped web captions

The first run downloads the model (`medium.en` is ~1.5 GB); after that it's
cached. On a CPU, expect transcription to take roughly **20–40% of the audio's
length** with `medium.en` (a 40-minute file ≈ 8–15 minutes).

### Choosing a model

| Model | Size | Speed | Accuracy | Good for |
|-------|------|-------|----------|----------|
| `tiny.en` / `base.en` | small | fastest | rough | quick drafts |
| `small.en` | medium | fast | good | most uses |
| `medium.en` | ~1.5 GB | moderate | great | **recommended default** |
| `large-v3` | ~3 GB | slowest | best | accuracy-critical, multilingual |

The `.en` models are English-only and a bit faster/more accurate for English
audio. Drop the `.en` (e.g. `large-v3`) for other languages.

---

## Troubleshooting

**`This video is not available` / `n challenge solving failed`**
yt-dlp is outdated or missing the JS-challenge solver. Run `pip install -U
yt-dlp` and make sure every command includes `--remote-components ejs:github`
(and that `deno` is installed).

**`Applying the VAD filter requires the onnxruntime package`**
The optional voice-activity (silence-trimming) filter needs `onnxruntime`,
which can break against NumPy 2.x. The script above already omits the VAD
filter, so this won't occur. If you re-add `vad_filter=True` and hit this,
either `pip install -U onnxruntime` or just leave VAD off — core transcription
is unaffected.

**`A module compiled using NumPy 1.x cannot be run in NumPy 2.x`**
A non-fatal warning; transcription still runs. To silence it, you can pin
`pip install "numpy<2"` in this environment.

**Filenames have weird `｜` characters**
yt-dlp replaces characters illegal in filenames (like `|`) with look-alikes.
That's expected and harmless.

---

## A note on usage

Download content you have the right to use — your own videos, content you have
permission for, or where the license/terms allow it. Respect copyright and
YouTube's Terms of Service.
