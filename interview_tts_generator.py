import asyncio
import csv
import os
import re
import sys
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import List, Tuple, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Optional imports handled gracefully
try:
    import edge_tts
except Exception:
    edge_tts = None

try:
    from docx import Document
except Exception:
    Document = None


APP_TITLE = "Interview MP3 + Subtitle Generator"
DEFAULT_PAUSE_MS = 1200
DEFAULT_INTERVIEWER_VOICE = "en-US-AndrewMultilingualNeural"
DEFAULT_CANDIDATE_VOICE = "en-US-BrianMultilingualNeural"

VOICE_PRESETS = {
    "International Male 1": "en-US-BrianMultilingualNeural",
    "International Male 2": "en-US-AndrewMultilingualNeural",
    "International Female 1": "en-US-AvaMultilingualNeural",
    "International Female 2": "en-GB-SoniaNeural",
    "Interviewer Male": "en-GB-RyanNeural",
    "Interviewer Female": "en-GB-SoniaNeural",
}


@dataclass
class Segment:
    role: str  # Q or A
    text: str
    title: str


def sanitize_filename(name: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|]", "_", name)
    name = re.sub(r"\s+", "_", name.strip())
    return name[:120] if name else "output"


def format_ms(ms: int) -> str:
    hours = ms // 3600000
    ms %= 3600000
    minutes = ms // 60000
    ms %= 60000
    seconds = ms // 1000
    millis = ms % 1000
    return f"{hours:02}:{minutes:02}:{seconds:02},{millis:03}"


def estimate_duration_ms(text: str, rate: float = 1.0) -> int:
    words = max(1, len(text.split()))
    base_wpm = 145  # natural interview speaking rate
    effective_wpm = max(80, min(230, base_wpm * rate))
    minutes = words / effective_wpm
    duration_ms = int(minutes * 60 * 1000)
    return max(1200, duration_ms)


def split_sentences(text: str) -> List[str]:
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return []
    parts = re.split(r"(?<=[.!?])\s+(?=[A-Z0-9])", text)
    cleaned = [p.strip() for p in parts if p.strip()]
    return cleaned if cleaned else [text]


def parse_dialogue_text(raw: str) -> List[Segment]:
    raw = raw.replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.strip() for line in raw.split("\n")]

    segments: List[Segment] = []
    current_role = None
    buffer: List[str] = []
    question_no = 0

    def flush_buffer():
        nonlocal buffer, current_role, question_no
        if current_role and buffer:
            text = " ".join(buffer).strip()
            if text:
                title = f"Q{question_no:02d}" if current_role == "Q" else f"Q{question_no:02d}_Answer"
                segments.append(Segment(role=current_role, text=text, title=title))
        buffer = []

    for line in lines:
        if not line:
            continue
        m = re.match(r"^(Q|Question)\s*[:.-]\s*(.*)$", line, re.I)
        if m:
            flush_buffer()
            question_no += 1
            current_role = "Q"
            content = m.group(2).strip()
            buffer = [content] if content else []
            continue
        m = re.match(r"^(A|Answer)\s*[:.-]\s*(.*)$", line, re.I)
        if m:
            flush_buffer()
            if question_no == 0:
                question_no = 1
            current_role = "A"
            content = m.group(2).strip()
            buffer = [content] if content else []
            continue

        # numbered question patterns like 1. Tell me about yourself
        m = re.match(r"^(\d+)\s*[.)-]\s*(.*)$", line)
        if m and len(m.group(2).split()) > 2:
            flush_buffer()
            question_no += 1
            current_role = "Q"
            buffer = [m.group(2).strip()]
            continue

        buffer.append(line)

    flush_buffer()
    return segments


def read_txt(path: Path) -> str:
    for enc in ("utf-8", "utf-8-sig", "cp949", "euc-kr"):
        try:
            return path.read_text(encoding=enc)
        except Exception:
            continue
    raise UnicodeDecodeError("read_txt", b"", 0, 1, f"Unable to decode file: {path}")


def read_docx(path: Path) -> str:
    if Document is None:
        raise RuntimeError("python-docx is not installed. Please run: pip install python-docx")
    doc = Document(str(path))
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return "\n".join(paras)


def load_script(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".txt":
        return read_txt(path)
    if ext == ".docx":
        return read_docx(path)
    if ext == ".csv":
        rows = []
        with open(path, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                q = (row.get("Q") or row.get("Question") or "").strip()
                a = (row.get("A") or row.get("Answer") or "").strip()
                if q:
                    rows.append(f"Q: {q}")
                if a:
                    rows.append(f"A: {a}")
        return "\n".join(rows)
    raise ValueError("Supported input: .txt, .docx, .csv")


async def tts_edge_save(text: str, voice: str, rate_pct: int, pitch_hz: int, out_path: Path):
    if edge_tts is None:
        raise RuntimeError("edge-tts is not installed. Please run: pip install edge-tts")
    rate = f"{rate_pct:+d}%"
    pitch = f"{pitch_hz:+d}Hz"
    communicate = edge_tts.Communicate(text=text, voice=voice, rate=rate, pitch=pitch)
    await communicate.save(str(out_path))


def build_srt_for_segments(segments: List[Segment], pause_ms: int, speaking_rate: float) -> str:
    entries = []
    current_ms = 0
    idx = 1

    for seg in segments:
        sentences = split_sentences(seg.text)
        for sentence in sentences:
            dur = estimate_duration_ms(sentence, rate=speaking_rate)
            start_ms = current_ms
            end_ms = current_ms + dur
            speaker = "Interviewer" if seg.role == "Q" else "Candidate"
            text = f"[{speaker}] {sentence}"
            entries.append(f"{idx}\n{format_ms(start_ms)} --> {format_ms(end_ms)}\n{text}\n")
            idx += 1
            current_ms = end_ms
        current_ms += pause_ms

    return "\n".join(entries)


def build_lrc_for_segments(segments: List[Segment], pause_ms: int, speaking_rate: float) -> str:
    lines = []
    current_ms = 0
    for seg in segments:
        speaker = "Interviewer" if seg.role == "Q" else "Candidate"
        for sentence in split_sentences(seg.text):
            mm = current_ms // 60000
            ss = (current_ms % 60000) / 1000
            lines.append(f"[{mm:02}:{ss:05.2f}][{speaker}] {sentence}")
            current_ms += estimate_duration_ms(sentence, rate=speaking_rate)
        current_ms += pause_ms
    return "\n".join(lines)


async def generate_all(
    segments: List[Segment],
    output_dir: Path,
    interviewer_voice: str,
    candidate_voice: str,
    rate_pct: int,
    pitch_hz: int,
    pause_ms: int,
    status_cb=None,
):
    output_dir.mkdir(parents=True, exist_ok=True)

    audio_files: List[Path] = []
    speaking_rate = 1.0 + (rate_pct / 100.0)
    speaking_rate = max(0.7, min(1.8, speaking_rate))

    # Individual files
    for i, seg in enumerate(segments, start=1):
        prefix = f"{i:02d}_{sanitize_filename(seg.title)}"
        mp3_path = output_dir / f"{prefix}.mp3"
        voice = interviewer_voice if seg.role == "Q" else candidate_voice
        if status_cb:
            status_cb(f"Generating {mp3_path.name} ...")
        await tts_edge_save(seg.text, voice, rate_pct, pitch_hz, mp3_path)
        audio_files.append(mp3_path)

    # Full combined script text for user to merge externally if desired
    full_script_path = output_dir / "full_script_for_batch_reading.txt"
    with open(full_script_path, "w", encoding="utf-8") as f:
        for seg in segments:
            speaker = "Interviewer" if seg.role == "Q" else "Candidate"
            f.write(f"[{speaker}] {seg.text}\n\n")

    # Subtitle files
    srt = build_srt_for_segments(segments, pause_ms, speaking_rate)
    lrc = build_lrc_for_segments(segments, pause_ms, speaking_rate)
    (output_dir / "full_interview_practice.srt").write_text(srt, encoding="utf-8")
    (output_dir / "full_interview_practice.lrc").write_text(lrc, encoding="utf-8")

    # M3U playlist
    playlist = "#EXTM3U\n" + "\n".join(str(p.name) for p in audio_files)
    (output_dir / "playlist.m3u").write_text(playlist, encoding="utf-8")

    # Export parsed script CSV
    with open(output_dir / "parsed_segments.csv", "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["No", "Role", "Text"])
        for i, seg in enumerate(segments, start=1):
            writer.writerow([i, seg.role, seg.text])

    if status_cb:
        status_cb("Completed successfully.")


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("900x700")
        self.root.minsize(680, 400)

        self.input_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.cwd() / "tts_output"))
        self.interviewer_voice = tk.StringVar(value=DEFAULT_INTERVIEWER_VOICE)
        self.candidate_voice = tk.StringVar(value=DEFAULT_CANDIDATE_VOICE)
        self.rate_pct = tk.IntVar(value=0)
        self.pitch_hz = tk.IntVar(value=0)
        self.pause_ms = tk.IntVar(value=DEFAULT_PAUSE_MS)
        self.status = tk.StringVar(value="Ready")

        self._build_ui()

    def _build_ui(self):
        # Pin the action bar to the bottom of root FIRST so it is always visible
        # regardless of window height. tkinter allocates space in pack order.
        bottom = ttk.Frame(self.root, padding=(12, 6))
        bottom.pack(side="bottom", fill="x")
        ttk.Button(bottom, text="Generate MP3 + SRT", command=self.start_generation).pack(side="left")
        ttk.Label(bottom, textvariable=self.status).pack(side="left", padx=12)

        ttk.Separator(self.root, orient="horizontal").pack(side="bottom", fill="x")

        # Scrollable area for form sections
        canvas = tk.Canvas(self.root, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        frm = ttk.Frame(canvas, padding=12)
        canvas_win = canvas.create_window((0, 0), window=frm, anchor="nw")

        def _on_frame_resize(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_resize(event):
            canvas.itemconfig(canvas_win, width=event.width)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        frm.bind("<Configure>", _on_frame_resize)
        canvas.bind("<Configure>", _on_canvas_resize)
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # 1) Input Script
        top = ttk.LabelFrame(frm, text="1) Input Script", padding=10)
        top.pack(fill="x", pady=6)
        ttk.Label(top, text="Input file (.txt / .docx / .csv)").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.input_path, width=80).grid(row=1, column=0, padx=(0, 8), sticky="ew")
        ttk.Button(top, text="Browse", command=self.browse_input).grid(row=1, column=1)
        top.columnconfigure(0, weight=1)

        # 2) Output
        mid = ttk.LabelFrame(frm, text="2) Output", padding=10)
        mid.pack(fill="x", pady=6)
        ttk.Label(mid, text="Output folder").grid(row=0, column=0, sticky="w")
        ttk.Entry(mid, textvariable=self.output_dir, width=80).grid(row=1, column=0, padx=(0, 8), sticky="ew")
        ttk.Button(mid, text="Browse", command=self.browse_output).grid(row=1, column=1)
        mid.columnconfigure(0, weight=1)

        # 3) Voice & Timing
        voice = ttk.LabelFrame(frm, text="3) Voice & Timing", padding=10)
        voice.pack(fill="x", pady=6)
        ttk.Label(voice, text="Interviewer voice").grid(row=0, column=0, sticky="w")
        interviewer_combo = ttk.Combobox(voice, values=list(VOICE_PRESETS.values()), textvariable=self.interviewer_voice, width=38)
        interviewer_combo.grid(row=1, column=0, sticky="w", padx=(0, 12))
        ttk.Label(voice, text="Candidate voice").grid(row=0, column=1, sticky="w")
        candidate_combo = ttk.Combobox(voice, values=list(VOICE_PRESETS.values()), textvariable=self.candidate_voice, width=38)
        candidate_combo.grid(row=1, column=1, sticky="w")
        ttk.Label(voice, text="Speed (%)").grid(row=2, column=0, sticky="w", pady=(10, 0))
        ttk.Spinbox(voice, from_=-50, to=50, textvariable=self.rate_pct, width=10).grid(row=3, column=0, sticky="w")
        ttk.Label(voice, text="Pitch (Hz)").grid(row=2, column=1, sticky="w", pady=(10, 0))
        ttk.Spinbox(voice, from_=-50, to=50, textvariable=self.pitch_hz, width=10).grid(row=3, column=1, sticky="w")
        ttk.Label(voice, text="Pause after each Q/A (ms)").grid(row=4, column=0, sticky="w", pady=(10, 0))
        ttk.Spinbox(voice, from_=0, to=5000, increment=100, textvariable=self.pause_ms, width=10).grid(row=5, column=0, sticky="w")

        # 4) Script Format Example
        preview = ttk.LabelFrame(frm, text="4) Script Format Example", padding=10)
        preview.pack(fill="x", pady=6)
        sample = (
            "Q: Please introduce yourself.\n"
            "A: My name is Lee, and I have worked in marine engineering for more than twenty years...\n\n"
            "Q: Why do you want to join Seapeak?\n"
            "A: I am interested in Seapeak because the company has a strong safety culture...\n"
        )
        txt = tk.Text(preview, height=8, wrap="word")
        txt.insert("1.0", sample)
        txt.config(state="disabled")
        txt.pack(fill="x")

    def browse_input(self):
        path = filedialog.askopenfilename(filetypes=[("Supported", "*.txt *.docx *.csv"), ("All files", "*.*")])
        if path:
            self.input_path.set(path)

    def browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def set_status(self, text: str):
        self.status.set(text)
        self.root.update_idletasks()

    def start_generation(self):
        if not self.input_path.get().strip():
            messagebox.showerror(APP_TITLE, "Please choose an input file first.")
            return

        thread = threading.Thread(target=self.run_generation, daemon=True)
        thread.start()

    def run_generation(self):
        try:
            input_path = Path(self.input_path.get())
            output_dir = Path(self.output_dir.get())
            self.set_status("Loading script...")
            raw = load_script(input_path)
            segments = parse_dialogue_text(raw)
            if not segments:
                raise RuntimeError("No valid Q/A segments found. Please check your file format.")
            if edge_tts is None:
                raise RuntimeError("Missing package: edge-tts. Run: pip install edge-tts")

            asyncio.run(generate_all(
                segments=segments,
                output_dir=output_dir,
                interviewer_voice=self.interviewer_voice.get(),
                candidate_voice=self.candidate_voice.get(),
                rate_pct=self.rate_pct.get(),
                pitch_hz=self.pitch_hz.get(),
                pause_ms=self.pause_ms.get(),
                status_cb=self.set_status,
            ))
            self.set_status(f"Done. Output: {output_dir}")
            messagebox.showinfo(APP_TITLE, f"Completed successfully.\n\nSaved to:\n{output_dir}")
        except Exception as e:
            self.set_status("Error")
            messagebox.showerror(APP_TITLE, str(e))


if __name__ == "__main__":
    root = tk.Tk()
    try:
        ttk.Style().theme_use("clam")
    except Exception:
        pass
    app = App(root)
    root.mainloop()
