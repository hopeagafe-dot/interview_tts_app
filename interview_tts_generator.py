import asyncio
import csv
import json
import os
import re
import sys
import threading
from dataclasses import dataclass
from datetime import date as _date
from pathlib import Path
from typing import List, Tuple, Optional


def resource_path(relative: str) -> Path:
    """PyInstaller --onefile exe와 일반 실행 모두에서 리소스 파일 경로를 반환.
    exe 실행 시: sys._MEIPASS (임시 압축 해제 폴더)
    일반 실행 시: 스크립트 디렉터리
    """
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    return base / relative

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Optional imports handled gracefully
try:
    import edge_tts
except Exception:
    edge_tts = None

try:
    from docx import Document
    _docx_err = None
except Exception as _e:
    Document = None
    _docx_err = repr(_e)   # 실제 원인을 보존해 에러 메시지에 표시


APP_TITLE = "Interview MP3 + Subtitle Generator"
DEFAULT_PAUSE_MS = 1200


class GenerationCancelled(Exception):
    """Raised inside generate_all() when the user clicks Cancel."""

# Voice presets: {display label shown in UI : edge-tts voice ID}
# The StringVar stores the display label; the voice ID is resolved at generation time.
VOICE_PRESETS = {
    # ── US Male ──────────────────────────────────────────────────────────
    "[US Male]  Andrew  (Multilingual)": "en-US-AndrewMultilingualNeural",
    "[US Male]  Brian   (Multilingual)": "en-US-BrianMultilingualNeural",
    "[US Male]  Guy":                    "en-US-GuyNeural",
    "[US Male]  Christopher":            "en-US-ChristopherNeural",
    "[US Male]  Eric":                   "en-US-EricNeural",
    "[US Male]  Roger":                  "en-US-RogerNeural",
    # ── US Female ────────────────────────────────────────────────────────
    "[US Female] Ava    (Multilingual)": "en-US-AvaMultilingualNeural",
    "[US Female] Emma   (Multilingual)": "en-US-EmmaMultilingualNeural",
    "[US Female] Jenny":                 "en-US-JennyNeural",
    "[US Female] Aria":                  "en-US-AriaNeural",
    "[US Female] Michelle":              "en-US-MichelleNeural",
    "[US Female] Sara":                  "en-US-SaraNeural",
    # ── UK Male ──────────────────────────────────────────────────────────
    "[UK Male]  Ryan":                   "en-GB-RyanNeural",
    "[UK Male]  Thomas":                 "en-GB-ThomasNeural",
    # ── UK Female ────────────────────────────────────────────────────────
    "[UK Female] Sonia":                 "en-GB-SoniaNeural",
    "[UK Female] Libby":                 "en-GB-LibbyNeural",
    "[UK Female] Maisie":                "en-GB-MaisieNeural",
}

DEFAULT_INTERVIEWER_VOICE = "[UK Male]  Ryan"
DEFAULT_CANDIDATE_VOICE   = "[US Male]  Andrew  (Multilingual)"
DEFAULT_PAUSE_QNA_MS      = 2000   # inter-pair gap in merge mode


# ── MP3 무음 프레임 (MPEG1 Layer3 128kbps 44100Hz Stereo) ─────────────────────
# 각 프레임 = 417 bytes, 26 ms (1152 samples / 44100 Hz)
_SILENT_MP3_FRAME = b'\xff\xfb\x90\x00' + b'\x00' * 413
_SILENT_FRAME_MS  = 26


def _silence_bytes(ms: int) -> bytes:
    """Q-A 사이 또는 pair 사이에 삽입할 MP3 무음 바이트."""
    if ms <= 0:
        return b""
    return _SILENT_MP3_FRAME * max(1, round(ms / _SILENT_FRAME_MS))


# ── Config 저장/불러오기 ──────────────────────────────────────────────────────
def _config_path() -> Path:
    """실행 파일 옆 config.json (PyInstaller frozen exe 에서도 동작)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent / "config.json"
    return Path(__file__).parent / "config.json"


def load_config() -> dict:
    defaults = {
        "input_path":        "",
        "output_dir":        str(Path.cwd() / "tts_output"),
        "interviewer_voice": DEFAULT_INTERVIEWER_VOICE,
        "candidate_voice":   DEFAULT_CANDIDATE_VOICE,
        "rate_pct":          0,
        "pitch_hz":          0,
        "pause_ms":          DEFAULT_PAUSE_MS,
        "pause_qna_ms":      DEFAULT_PAUSE_QNA_MS,
        "output_mode":       "split",
    }
    try:
        data = json.loads(_config_path().read_text(encoding="utf-8"))
        return {**defaults, **data}
    except Exception:
        return defaults


def save_config(data: dict) -> None:
    try:
        _config_path().write_text(
            json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8"
        )
    except Exception:
        pass


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
            content = m.group(2).strip()
            # ALL-CAPS 섹션 헤더(예: "2. ENGINEERING / LNG")는 Q로 오인식하지 않음
            if not any(c.islower() for c in content):
                buffer.append(line)
                continue
            flush_buffer()
            question_no += 1
            current_role = "Q"
            buffer = [content]
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
        detail = f"\nDetail: {_docx_err}" if _docx_err else ""
        raise RuntimeError(
            f"python-docx could not be loaded.{detail}\n"
            "If running from source: pip install python-docx"
        )
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


async def tts_edge_stream(
    text: str, voice: str, rate_pct: int, pitch_hz: int, out_path: Path
) -> List[Tuple[int, int, str]]:
    """TTS 생성 + MP3 저장 + WordBoundary 이벤트로 실제 단어 타이밍 반환.

    Returns: [(start_ms, end_ms, word), ...] — 세그먼트 시작 기준 상대 시간(ms).
    edge-tts 스트리밍 실패 시 communicate.save() 로 폴백하며 빈 리스트 반환.
    """
    if edge_tts is None:
        raise RuntimeError("edge-tts is not installed. Please run: pip install edge-tts")
    rate = f"{rate_pct:+d}%"
    pitch = f"{pitch_hz:+d}Hz"

    audio_chunks: List[bytes] = []
    timings: List[Tuple[int, int, str]] = []
    try:
        communicate = edge_tts.Communicate(text=text, voice=voice, rate=rate, pitch=pitch)
        async for event in communicate.stream():
            if event["type"] == "audio":
                audio_chunks.append(event["data"])
            elif event["type"] == "WordBoundary":
                # edge-tts 단위: 100 나노초 → //10000 으로 ms 변환
                start_ms = event["offset"] // 10000
                dur_ms   = event["duration"] // 10000
                timings.append((start_ms, start_ms + dur_ms, event["text"]))
        out_path.write_bytes(b"".join(audio_chunks))
    except Exception:
        # 스트리밍 미지원 버전 폴백
        communicate = edge_tts.Communicate(text=text, voice=voice, rate=rate, pitch=pitch)
        await communicate.save(str(out_path))
    return timings


# ── 자막 엔트리 헬퍼 ──────────────────────────────────────────────────────────
# WordBoundary = 1 단어 = 1 엔트리 (음절 수준에 가장 근접한 단위)

def _seg_entries(
    seg: Segment,
    timings: List[Tuple[int, int, str]],
    speaking_rate: float,
    base_ms: int = 0,
) -> List[Tuple[int, int, str]]:
    """Return list of (start_ms, end_ms, word) — one entry per word.

    Uses real WordBoundary timings when available (1 word per entry).
    Fallback: distributes estimated segment duration evenly across words.
    Speaker labels are intentionally omitted.
    """
    result: List[Tuple[int, int, str]] = []
    if timings:
        # 1 단어 = 1 자막 엔트리 (WordBoundary 단위)
        for start_ms, end_ms, word in timings:
            result.append((base_ms + start_ms, base_ms + end_ms, word))
    else:
        # 폴백: 문장을 단어로 분리하고 시간을 균등 분배
        inner = 0
        for sentence in split_sentences(seg.text):
            words = sentence.split()
            if not words:
                continue
            dur = estimate_duration_ms(sentence, speaking_rate)
            word_dur = max(80, dur // len(words))
            for word in words:
                result.append((base_ms + inner, base_ms + inner + word_dur, word))
                inner += word_dur
    return result


def _entries_to_srt(entries: List[Tuple[int, int, str]], start_idx: int = 1) -> str:
    parts = []
    for i, (s, e, txt) in enumerate(entries, start=start_idx):
        parts.append(f"{i}\n{format_ms(s)} --> {format_ms(e)}\n{txt}\n")
    return "\n".join(parts)


def _entries_to_lrc(entries: List[Tuple[int, int, str]]) -> str:
    lines = []
    for s, _e, txt in entries:
        mm = s // 60000
        ss = (s % 60000) / 1000
        lines.append(f"[{mm:02}:{ss:05.2f}]{txt}")
    return "\n".join(lines)


async def generate_all(
    segments: List[Segment],
    output_dir: Path,
    interviewer_voice: str,
    candidate_voice: str,
    rate_pct: int,
    pitch_hz: int,
    pause_ms: int,          # Q → A 사이 무음 간격 (ms)
    pause_qna_ms: int,      # pair → pair 사이 무음 간격 (merge 모드 전용, ms)
    output_mode: str,       # "split" | "merge"
    status_cb=None,
    cancel_event: Optional[threading.Event] = None,
):
    output_dir.mkdir(parents=True, exist_ok=True)

    audio_files: List[Path] = []
    speaking_rate = max(0.7, min(1.8, 1.0 + rate_pct / 100.0))
    all_entries: List[Tuple[int, int, str]] = []
    full_offset_ms = 0

    # ── Q+A 세그먼트 → pair 목록 ──────────────────────────────────────────────
    pairs: List[Tuple[Optional[Segment], Optional[Segment]]] = []
    idx = 0
    while idx < len(segments):
        seg = segments[idx]
        if (seg.role == "Q"
                and idx + 1 < len(segments)
                and segments[idx + 1].role == "A"):
            pairs.append((seg, segments[idx + 1]))
            idx += 2
        else:
            pairs.append((seg, None) if seg.role == "Q" else (None, seg))
            idx += 1

    # ── 공통: pair 하나의 TTS 생성 ─────────────────────────────────────────────
    async def _gen_pair(pair_no, q_seg, a_seg):
        """Returns (q_bytes, q_timings, a_bytes, a_timings, q_dur_ms, a_base)."""
        q_bytes, q_timings = b"", []
        a_bytes, a_timings = b"", []
        if q_seg:
            q_tmp = output_dir / f"__tmp_q{pair_no:02d}.mp3"
            q_timings = await tts_edge_stream(
                q_seg.text, interviewer_voice, rate_pct, pitch_hz, q_tmp
            )
            q_bytes = q_tmp.read_bytes()
            q_tmp.unlink(missing_ok=True)
        if a_seg:
            a_tmp = output_dir / f"__tmp_a{pair_no:02d}.mp3"
            a_timings = await tts_edge_stream(
                a_seg.text, candidate_voice, rate_pct, pitch_hz, a_tmp
            )
            a_bytes = a_tmp.read_bytes()
            a_tmp.unlink(missing_ok=True)
        q_dur_ms = (q_timings[-1][1] if q_timings
                    else sum(estimate_duration_ms(s, speaking_rate)
                             for s in split_sentences(q_seg.text)) if q_seg else 0)
        a_base = q_dur_ms + pause_ms
        return q_bytes, q_timings, a_bytes, a_timings, q_dur_ms, a_base

    # ── MERGE 모드 ─────────────────────────────────────────────────────────────
    if output_mode == "merge":
        today_str   = _date.today().strftime("%Y-%m-%d")
        merged_stem = f"1~{len(pairs)} Q&A ({today_str})"
        merged_audio: bytes = b""

        for pair_no, (q_seg, a_seg) in enumerate(pairs, start=1):
            if cancel_event and cancel_event.is_set():
                raise GenerationCancelled()
            if status_cb:
                status_cb(f"[{pair_no}/{len(pairs)}] Generating Q&A pair ...")

            q_bytes, q_timings, a_bytes, a_timings, q_dur_ms, a_base = \
                await _gen_pair(pair_no, q_seg, a_seg)

            # Q + 무음(pause_ms) + A
            merged_audio += q_bytes + _silence_bytes(pause_ms) + a_bytes

            # pair 사이 무음 (마지막 pair 제외)
            if pair_no < len(pairs):
                merged_audio += _silence_bytes(pause_qna_ms)

            # 자막 엔트리 수집
            pair_entries: List[Tuple[int, int, str]] = []
            if q_seg:
                pair_entries.extend(_seg_entries(q_seg, q_timings, speaking_rate, base_ms=0))
            if a_seg:
                pair_entries.extend(_seg_entries(a_seg, a_timings, speaking_rate, base_ms=a_base))
            for s, e, txt in pair_entries:
                all_entries.append((full_offset_ms + s, full_offset_ms + e, txt))

            # 전체 오프셋 전진
            a_dur_ms = (a_timings[-1][1] if a_timings
                        else sum(estimate_duration_ms(s, speaking_rate)
                                 for s in split_sentences(a_seg.text)) if a_seg else 0)
            pair_dur_ms = a_base + a_dur_ms
            full_offset_ms += pair_dur_ms + (pause_qna_ms if pair_no < len(pairs) else 0)

        (output_dir / f"{merged_stem}.mp3").write_bytes(merged_audio)
        audio_files.append(output_dir / f"{merged_stem}.mp3")
        (output_dir / f"{merged_stem}.srt").write_text(
            _entries_to_srt(all_entries), encoding="utf-8"
        )
        (output_dir / f"{merged_stem}.lrc").write_text(
            _entries_to_lrc(all_entries), encoding="utf-8"
        )

    # ── SPLIT 모드 ─────────────────────────────────────────────────────────────
    else:
        for pair_no, (q_seg, a_seg) in enumerate(pairs, start=1):
            if cancel_event and cancel_event.is_set():
                raise GenerationCancelled()
            q_label  = q_seg.title if q_seg else a_seg.title.replace("_Answer", "")
            prefix   = f"{pair_no:02d}_{q_label}_Q&A"
            mp3_path = output_dir / f"{prefix}.mp3"
            if status_cb:
                status_cb(f"Generating {mp3_path.name} ...")

            q_bytes, q_timings, a_bytes, a_timings, q_dur_ms, a_base = \
                await _gen_pair(pair_no, q_seg, a_seg)

            # Q + 무음(pause_ms) + A → 단일 MP3
            mp3_path.write_bytes(q_bytes + _silence_bytes(pause_ms) + a_bytes)
            audio_files.append(mp3_path)

            pair_entries: List[Tuple[int, int, str]] = []
            if q_seg:
                pair_entries.extend(_seg_entries(q_seg, q_timings, speaking_rate, base_ms=0))
            if a_seg:
                pair_entries.extend(_seg_entries(a_seg, a_timings, speaking_rate, base_ms=a_base))

            (output_dir / f"{prefix}.srt").write_text(
                _entries_to_srt(pair_entries), encoding="utf-8"
            )
            (output_dir / f"{prefix}.lrc").write_text(
                _entries_to_lrc(pair_entries), encoding="utf-8"
            )
            for s, e, txt in pair_entries:
                all_entries.append((full_offset_ms + s, full_offset_ms + e, txt))

            a_dur_ms = (a_timings[-1][1] if a_timings
                        else sum(estimate_duration_ms(s, speaking_rate)
                                 for s in split_sentences(a_seg.text)) if a_seg else 0)
            full_offset_ms += (a_base + a_dur_ms) + pause_ms

        (output_dir / "full_interview_practice.srt").write_text(
            _entries_to_srt(all_entries), encoding="utf-8"
        )
        (output_dir / "full_interview_practice.lrc").write_text(
            _entries_to_lrc(all_entries), encoding="utf-8"
        )

    # ── 공통 출력 파일 ──────────────────────────────────────────────────────────
    full_script_path = output_dir / "full_script_for_batch_reading.txt"
    with open(full_script_path, "w", encoding="utf-8") as f:
        for seg in segments:
            speaker = "Interviewer" if seg.role == "Q" else "Candidate"
            f.write(f"[{speaker}] {seg.text}\n\n")

    playlist = "#EXTM3U\n" + "\n".join(str(p.name) for p in audio_files)
    (output_dir / "playlist.m3u").write_text(playlist, encoding="utf-8")

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
        self.root.geometry("860x680")
        self.root.minsize(700, 500)

        # --- 변수 정의 (config.json 에서 이전 값 복원) ---
        cfg = load_config()
        self.input_path        = tk.StringVar(value=cfg["input_path"])
        self.output_dir        = tk.StringVar(value=cfg["output_dir"])
        self.interviewer_voice = tk.StringVar(value=cfg["interviewer_voice"])
        self.candidate_voice   = tk.StringVar(value=cfg["candidate_voice"])
        self.rate_pct          = tk.IntVar(value=int(cfg["rate_pct"]))
        self.pitch_hz          = tk.IntVar(value=int(cfg["pitch_hz"]))
        self.pause_ms          = tk.IntVar(value=int(cfg["pause_ms"]))
        self.pause_qna_ms      = tk.IntVar(value=int(cfg["pause_qna_ms"]))
        self.output_mode       = tk.StringVar(value=cfg["output_mode"])
        self.status            = tk.StringVar(value="Ready")

        # --- 스타일 및 색상 테마 설정 ---
        self.style = ttk.Style()
        self.style.theme_use("clam")  # 모던한 테마 기반

        # 주요 색상 정의
        self.color_primary = "#2B5797"    # 로고의 블루 계열 (프로페셔널한 느낌)
        self.color_secondary = "#E1E1E1"  # 연한 회색 (배경)
        self.color_accent = "#D9534F"     # 로고의 레드 계열 (포인트)
        self.color_text = "#333333"       # 진한 회색 (텍스트)
        self.color_white = "#FFFFFF"

        # Ttk 스타일 — 기본값: 흰색 배경 + Segoe UI (Windows 선명 렌더링)
        self.style.configure(".",
                             background=self.color_white, foreground=self.color_text)
        self.style.configure("TLabel",
                             background=self.color_white, font=("Segoe UI", 10))
        self.style.configure("TEntry",
                             fieldbackground=self.color_white, borderwidth=1, relief="solid")
        self.style.configure("TCombobox",
                             fieldbackground=self.color_white, borderwidth=1, relief="solid")
        self.style.configure("TSpinbox",
                             fieldbackground=self.color_white, borderwidth=1, relief="solid")
        self.style.configure("TLabelframe",
                             background=self.color_white, borderwidth=1, relief="solid")
        self.style.configure("TLabelframe.Label",
                             background=self.color_white,
                             foreground=self.color_primary,
                             font=("Segoe UI", 11, "bold"))
        # 헤더 전용 스타일 — 회색 배경
        self.style.configure("Header.TFrame", background=self.color_secondary)
        self.style.configure("Header.TLabel", background=self.color_secondary,
                             font=("Segoe UI", 10))

        self._build_ui()

    def _build_ui(self):
        self.root.configure(background=self.color_secondary)

        # 상단 헤더 — 회색 배경 유지
        header_frame = ttk.Frame(self.root, style="Header.TFrame", padding=(16, 8))
        header_frame.pack(side="top", fill="x")

        # MCE 로고 (694×731 원본 → subsample(10,10) ≈ 69×73 px)
        logo_path = resource_path("MCE_logo.png")
        if logo_path.exists():
            self.logo_img = tk.PhotoImage(file=str(logo_path)).subsample(10, 10)
            ttk.Label(header_frame, image=self.logo_img,
                      style="Header.TLabel").pack(side="left", padx=(0, 12))

        ttk.Label(header_frame, text=APP_TITLE,
                  font=("Segoe UI", 18, "bold"),
                  foreground=self.color_primary,
                  style="Header.TLabel").pack(side="left")

        ttk.Separator(self.root, orient="horizontal").pack(side="top", fill="x")

        # --- 하단 액션 바: canvas보다 먼저 pack해야 항상 보임 ---
        # tkinter pack은 선언 순서대로 공간을 배분하므로,
        # bottom_bar를 canvas(expand=True)보다 먼저 선언해야 공간이 확보됨.
        ttk.Separator(self.root, orient="horizontal").pack(side="bottom", fill="x")
        bottom_bar = ttk.Frame(self.root, style="TFrame", padding=(20, 10))
        bottom_bar.pack(side="bottom", fill="x")

        self.progress = ttk.Progressbar(bottom_bar, mode="indeterminate", length=120)
        self.progress.pack(side="left", padx=(0, 12))
        ttk.Label(bottom_bar, textvariable=self.status,
                  font=("Segoe UI", 10, "italic")).pack(side="left")
        # Generate 버튼: tk.Button으로 relief="raised" 3D 효과 (self.gen_btn 에 저장)
        self.gen_btn = tk.Button(
            bottom_bar, text="  Generate MP3 + SRT  ",
            bg=self.color_accent, fg="white",
            activebackground="#A93226", activeforeground="white",
            relief="raised", bd=3, cursor="hand2",
            font=("Segoe UI", 10, "bold"),
            command=self.start_generation,
        )
        self.gen_btn.pack(side="right", pady=3)

        # 메인 콘텐츠 영역 (스크롤 가능) — canvas는 반드시 마지막에 pack
        canvas = tk.Canvas(self.root, borderwidth=0, highlightthickness=0, background=self.color_white)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        frm = ttk.Frame(canvas, padding=20, style="TFrame")
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


        # --- 각 섹션 카드 디자인 및 배치 ---

        # 1) Input Script
        input_frame = ttk.LabelFrame(
            frm, text="1) Input Script :  Input file (.txt / .docx / .csv)",
            padding=(10, 6), style="TLabelframe")
        input_frame.pack(fill="x", pady=(0, 6))
        input_entry = ttk.Entry(input_frame, textvariable=self.input_path)
        input_entry.grid(row=0, column=0, padx=(0, 8), sticky="ew")
        tk.Button(input_frame, text="Browse",
                  bg=self.color_primary, fg="white",
                  activebackground="#1A3F7A", activeforeground="white",
                  relief="raised", bd=2, cursor="hand2",
                  font=("Segoe UI", 9, "bold"), padx=10, pady=2,
                  command=self.browse_input).grid(row=0, column=1)
        input_frame.columnconfigure(0, weight=1)

        # 2) Output Settings
        output_frame = ttk.LabelFrame(
            frm, text="2) Output Settings :  Output folder",
            padding=(10, 6), style="TLabelframe")
        output_frame.pack(fill="x", pady=6)
        output_entry = ttk.Entry(output_frame, textvariable=self.output_dir)
        output_entry.grid(row=0, column=0, padx=(0, 8), sticky="ew")
        tk.Button(output_frame, text="Browse",
                  bg=self.color_primary, fg="white",
                  activebackground="#1A3F7A", activeforeground="white",
                  relief="raised", bd=2, cursor="hand2",
                  font=("Segoe UI", 9, "bold"), padx=10, pady=2,
                  command=self.browse_output).grid(row=0, column=1)
        output_frame.columnconfigure(0, weight=1)

        # 3) Voice & Timing
        settings_frame = ttk.LabelFrame(frm, text="3) Voice & Timing",
                                         padding=(10, 6), style="TLabelframe")
        settings_frame.pack(fill="x", pady=6)
        
        voice_labels = list(VOICE_PRESETS.keys())
        
        # 음성 선택
        ttk.Label(settings_frame, text="Interviewer voice", style="TLabel").grid(row=0, column=0, sticky="w", pady=(0, 5))
        interviewer_combo = ttk.Combobox(settings_frame, values=voice_labels, textvariable=self.interviewer_voice, style="TCombobox", state="readonly")
        interviewer_combo.grid(row=1, column=0, sticky="ew", padx=(0, 15))
        
        ttk.Label(settings_frame, text="Candidate voice", style="TLabel").grid(row=0, column=1, sticky="w", pady=(0, 5))
        candidate_combo = ttk.Combobox(settings_frame, values=voice_labels, textvariable=self.candidate_voice, style="TCombobox", state="readonly")
        candidate_combo.grid(row=1, column=1, sticky="ew")
        
        settings_frame.columnconfigure(0, weight=1)
        settings_frame.columnconfigure(1, weight=1)

        # 상세 설정 (속도, 피치, 포즈)
        details_frame = ttk.Frame(settings_frame, padding=(0, 8, 0, 0))
        details_frame.grid(row=2, column=0, columnspan=2, sticky="ew")

        ttk.Label(details_frame, text="Speed (%)", style="TLabel").grid(row=0, column=0, sticky="w", pady=(0, 5))
        ttk.Spinbox(details_frame, from_=-50, to=50, textvariable=self.rate_pct, style="TSpinbox", width=10).grid(row=1, column=0, sticky="w", padx=(0, 20))

        ttk.Label(details_frame, text="Pitch (Hz)", style="TLabel").grid(row=0, column=1, sticky="w", pady=(0, 5))
        ttk.Spinbox(details_frame, from_=-50, to=50, textvariable=self.pitch_hz, style="TSpinbox", width=10).grid(row=1, column=1, sticky="w", padx=(0, 20))

        ttk.Label(details_frame, text="Pause Q → A (ms)", style="TLabel").grid(row=0, column=2, sticky="w", pady=(0, 5))
        ttk.Spinbox(details_frame, from_=0, to=5000, increment=100, textvariable=self.pause_ms, style="TSpinbox", width=14).grid(row=1, column=2, sticky="w")

        # ── 출력 모드: 분리 / 통합 ──────────────────────────────────────────
        ttk.Separator(details_frame, orient="horizontal").grid(
            row=2, column=0, columnspan=4, sticky="ew", pady=(12, 8)
        )
        ttk.Label(details_frame, text="Output mode:", style="TLabel").grid(
            row=3, column=0, sticky="w", pady=(0, 6)
        )
        mode_btn_frame = ttk.Frame(details_frame)
        mode_btn_frame.grid(row=3, column=1, columnspan=3, sticky="w", pady=(0, 6))
        ttk.Radiobutton(
            mode_btn_frame, text="Split  — separate file per question",
            variable=self.output_mode, value="split",
            command=self._on_mode_change,
        ).pack(side="left", padx=(0, 24))
        ttk.Radiobutton(
            mode_btn_frame, text="Merge  — single combined file",
            variable=self.output_mode, value="merge",
            command=self._on_mode_change,
        ).pack(side="left")

        ttk.Label(details_frame, text="Pause Q&A → Q&A (ms)\n(Merge mode only)",
                  style="TLabel").grid(row=4, column=0, sticky="w", pady=(0, 5))
        self.pause_qna_spin = ttk.Spinbox(
            details_frame, from_=0, to=10000, increment=100,
            textvariable=self.pause_qna_ms, style="TSpinbox", width=14,
        )
        self.pause_qna_spin.grid(row=4, column=1, sticky="w")

        # 초기 활성 상태 설정
        self._on_mode_change()

        # 4) Script Format Example
        preview_frame = ttk.LabelFrame(frm, text="4) Script Format Example",
                                        padding=(10, 6), style="TLabelframe")
        preview_frame.pack(fill="x", pady=6)

        sample = (
            "Q: Please introduce yourself.\n"
            "A: My name is Lee, and I have worked in marine engineering for more than twenty years...\n\n"
            "Q: Why do you want to join Seapeak?\n"
            "A: I am interested in Seapeak because the company has a strong safety culture...\n"
        )
        txt = tk.Text(preview_frame, height=7, wrap="word", font=("Consolas", 10),
                      background=self.color_white, foreground=self.color_text,
                      borderwidth=1, relief="solid")
        txt.insert("1.0", sample)
        txt.config(state="disabled")
        txt.pack(fill="x")



    # -------------------------------------------------------------------------
    def _on_mode_change(self):
        """Merge 모드일 때만 'Pause Q&A → Q&A' 스핀박스를 활성화."""
        state = "normal" if self.output_mode.get() == "merge" else "disabled"
        if hasattr(self, "pause_qna_spin"):
            self.pause_qna_spin.config(state=state)

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

    def cancel_generation(self):
        """Cancel 버튼 클릭 → 진행 중인 생성을 즉시 중단."""
        if hasattr(self, "_cancel_event"):
            self._cancel_event.set()
        self.set_status("Cancelling...")

    def _restore_gen_btn(self):
        """Generate 버튼을 원래 상태로 복원 (main thread에서 호출)."""
        self.gen_btn.config(
            text="  Generate MP3 + SRT  ",
            bg=self.color_accent,
            activebackground="#A93226",
            command=self.start_generation,
        )

    def start_generation(self):
        if not self.input_path.get().strip():
            messagebox.showerror(APP_TITLE, "Please choose an input file first.")
            return

        # ── 기존 출력 파일 존재 여부 확인 ──────────────────────────────────
        output_dir = Path(self.output_dir.get())
        if output_dir.exists() and any(output_dir.glob("*Q&A*.mp3")):
            result = messagebox.askyesnocancel(
                "Output files already exist",
                f"Q&A MP3 files already exist in:\n{output_dir}\n\n"
                "  Yes    →  Overwrite existing files\n"
                "  No     →  Choose a different output folder\n"
                "  Cancel →  Abort",
            )
            if result is None:          # Cancel → 중단
                return
            elif result is False:       # No → 새 폴더 선택
                new_dir = filedialog.askdirectory(title="Choose output folder")
                if not new_dir:
                    return
                self.output_dir.set(new_dir)

        # ── Cancel 버튼으로 교체 ───────────────────────────────────────────
        self._cancel_event = threading.Event()
        self.gen_btn.config(
            text="  Cancel  ",
            bg="#888888",
            activebackground="#666666",
            command=self.cancel_generation,
        )

        self.progress.start(10)
        self.set_status("Starting generation...")
        threading.Thread(target=self.run_generation, daemon=True).start()

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

            interviewer_voice_id = VOICE_PRESETS.get(
                self.interviewer_voice.get(), self.interviewer_voice.get()
            )
            candidate_voice_id = VOICE_PRESETS.get(
                self.candidate_voice.get(), self.candidate_voice.get()
            )
            # 설정 저장
            save_config({
                "input_path":        str(input_path),
                "output_dir":        str(output_dir),
                "interviewer_voice": self.interviewer_voice.get(),
                "candidate_voice":   self.candidate_voice.get(),
                "rate_pct":          self.rate_pct.get(),
                "pitch_hz":          self.pitch_hz.get(),
                "pause_ms":          self.pause_ms.get(),
                "pause_qna_ms":      self.pause_qna_ms.get(),
                "output_mode":       self.output_mode.get(),
            })
            asyncio.run(generate_all(
                segments=segments,
                output_dir=output_dir,
                interviewer_voice=interviewer_voice_id,
                candidate_voice=candidate_voice_id,
                rate_pct=self.rate_pct.get(),
                pitch_hz=self.pitch_hz.get(),
                pause_ms=self.pause_ms.get(),
                pause_qna_ms=self.pause_qna_ms.get(),
                output_mode=self.output_mode.get(),
                status_cb=self.set_status,
                cancel_event=self._cancel_event,
            ))
            self.set_status(f"Done. Output: {output_dir}")
            messagebox.showinfo(APP_TITLE, f"Completed successfully.\n\nSaved to:\n{output_dir}")
        except GenerationCancelled:
            self.set_status("Generation cancelled.")
        except Exception as e:
            self.set_status("Error")
            messagebox.showerror(APP_TITLE, str(e))
        finally:
            self.progress.stop()
            # Generate 버튼 복원은 반드시 main thread에서 실행
            self.root.after(0, self._restore_gen_btn)

if __name__ == "__main__":
    root = tk.Tk()
    # 창 아이콘 설정 (.ico → 윈도우 탐색기 / 작업표시줄 아이콘)
    # PyInstaller 변환 시에도 --icon=MCE_logo.ico 옵션으로 .exe 아이콘 지정
    ico_path = resource_path("MCE_logo.ico")
    if ico_path.exists():
        root.iconbitmap(str(ico_path))
    try:
        ttk.Style().theme_use("clam")
    except Exception:
        pass
    app = App(root)
    root.mainloop()
