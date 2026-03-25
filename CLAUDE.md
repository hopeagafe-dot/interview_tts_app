# Project: Interview MP3 + Subtitle Generator

## Goal
A Windows desktop Python GUI app that converts interview Q/A scripts into:
- per-item mp3 files
- full subtitle files (.srt, .lrc)
- optionally merged full practice audio later

## Current stack
- Python 3.x
- tkinter GUI
- edge-tts or similar TTS engine
- pydub for audio merging if needed
- python-docx / csv parsing

## Important rules
- Do not change features unrelated to the user’s request.
- Keep the app simple and stable.
- Prioritize Windows compatibility.
- Prefer minimal, maintainable code over fancy abstractions.
- Always explain which file and function was changed.
- Before editing, inspect the current file structure and summarize it briefly.
- After editing, run a syntax check or execution test when possible.

## UI rules
- The window must be resizable.
- Main controls must remain visible even on smaller screens.
- Use scrollable layout if needed.
- Avoid hard-coded geometry that causes clipping.

## Output rules
- Preserve current working features.
- Do not break txt/docx/csv parsing.
- Keep output filenames clean and predictable.

## When fixing bugs
- First identify exact cause.
- Then propose minimal fix.
- Then implement.
- Then verify with command output if possible.