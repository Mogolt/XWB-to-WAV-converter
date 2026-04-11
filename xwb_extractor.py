"""
XWB to WAV Extractor - GUI Version
Built with tkinter (included with Python, no extra installs needed)

To build as a standalone .exe:
    pip install pyinstaller
    python -m PyInstaller --onefile --noconsole --name "XWB Extractor" xwb_extractor.py
"""

import os
import sys
import json
import struct
import threading
import tempfile
try:
    import winsound
    WINSOUND_OK = True
except ImportError:
    WINSOUND_OK = False
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ── XWB parsing constants ────────────────────────────────────────────────────

SIGN_LE = b"WBND"
SIGN_BE = b"DNBW"
CODEC_PCM   = 0
CODEC_XMA   = 1
CODEC_ADPCM = 2
CODEC_WMA   = 3
ADPCM_BLOCKALIGN_OFFSET = 22
CHUNK = 65536

CONFIG_FILENAME  = "config.json"
RECENT_FILENAME  = "recent_folders.json"
MAX_RECENT       = 5
DEFAULT_CONFIG = {
    "_readme": "Optional: map hex track names to friendly names. Example below.",
    "_example": {
        "bio4bgm": {
            "00000000": "main_menu_theme",
            "00000001": "village_ambience"
        }
    },
    "track_names": {}
}


# ── WAV helpers ─────────────────────────────────────────────────────────────

def ru32_le(f): return struct.unpack("<I", f.read(4))[0]
def ru32_be(f): return struct.unpack(">I", f.read(4))[0]


def make_wav_header(codec, channels, rate, bits, align, data_size):
    if channels <= 0:
        channels = 1
    if codec == CODEC_PCM:
        fmt_tag     = 0x0001
        bits_per    = 8 << bits
        block_align = (bits_per // 8) * channels
        avg_bytes   = rate * block_align
        extra       = b""
    elif codec == CODEC_XMA:
        fmt_tag     = 0x0069
        bits_per    = 4
        block_align = 36 * channels
        avg_bytes   = (689 * block_align) + 4
        extra       = b"\x02\x00\x40\x00"
    elif codec == CODEC_ADPCM:
        fmt_tag           = 0x0002
        bits_per          = 4
        block_align       = (align + ADPCM_BLOCKALIGN_OFFSET) * channels
        avg_bytes         = 21 * block_align
        samples_per_block = ((block_align // channels - 7) * 2) + 2
        extra             = struct.pack("<HH", 2, samples_per_block)
    else:
        return None

    fmt_size  = 16 + len(extra)
    riff_size = 4 + 8 + fmt_size + 8 + data_size
    h  = b"RIFF" + struct.pack("<I", riff_size) + b"WAVE"
    h += b"fmt " + struct.pack("<I", fmt_size)
    h += struct.pack("<hHIIHH", fmt_tag, channels, rate, avg_bytes, block_align, bits_per)
    h += extra
    h += b"data" + struct.pack("<I", data_size)
    return h


def copy_bytes(f_in, f_out, size):
    remaining = size
    while remaining > 0:
        buf = f_in.read(min(CHUNK, remaining))
        if not buf:
            break
        f_out.write(buf)
        remaining -= len(buf)


# ── XWB extraction ───────────────────────────────────────────────────────────

def extract_xwb(xwb_path, out_dir, track_names_map=None, stop_event=None):
    with open(xwb_path, "rb") as f:
        sig = f.read(4)
        if sig == SIGN_LE:
            ru32 = ru32_le
        elif sig == SIGN_BE:
            ru32 = ru32_be
        else:
            f.seek(0)
            scan = f.read(min(1024 * 1024, os.path.getsize(xwb_path)))
            found = -1
            for i in range(len(scan) - 8):
                s = scan[i:i+4]
                if (s == SIGN_LE and scan[i+7] == 0) or (s == SIGN_BE and scan[i+4] == 0):
                    found = i
                    break
            if found < 0:
                raise ValueError("Not a valid XWB file")
            f.seek(found + 4)
            sig  = scan[found:found+4]
            ru32 = ru32_le if sig == SIGN_LE else ru32_be

        version      = ru32(f)
        segments     = []
        last_segment = 3 if version <= 3 else 4
        if version >= 42:
            f.read(4)
        for _ in range(last_segment + 1):
            segments.append((ru32(f), ru32(f)))
        while len(segments) < 5:
            segments.append((0, 0))

        bank_data_off = f.tell() if version == 1 else segments[0][0]
        f.seek(bank_data_off)

        flags       = ru32(f)
        entry_count = ru32(f)
        is_compact  = bool(flags & 0x00020000)
        f.read(16 if version in (2, 3) else 64)

        if version == 1:
            wavebank_offset   = f.tell()
            meta_element_size = 20
            alignment         = 4
            compact_format    = 0
        else:
            meta_element_size = ru32(f)
            ru32(f)
            alignment         = ru32(f)
            wavebank_offset   = segments[1][0]
            compact_format    = ru32(f) if is_compact else 0

        playregion_offset = segments[last_segment][0]
        if not playregion_offset:
            playregion_offset = wavebank_offset + (entry_count * meta_element_size)

        os.makedirs(out_dir, exist_ok=True)
        output_files = []

        for entry_idx in range(entry_count):
            if stop_event and stop_event.is_set():
                break

            ep = wavebank_offset + entry_idx * meta_element_size
            f.seek(ep)

            if is_compact:
                raw_val  = ru32(f)
                fmt      = compact_format
                play_off = (raw_val & 0x1FFFFF) * alignment
                if entry_idx == entry_count - 1:
                    play_len = segments[last_segment][1] - play_off
                else:
                    f.seek(ep + meta_element_size)
                    play_len = (ru32(f) & 0x1FFFFF) * alignment - play_off
            else:
                fmt = play_off = play_len = 0
                if version == 1:
                    fmt      = ru32(f)
                    play_off = ru32(f)
                    play_len = ru32(f)
                else:
                    if meta_element_size >= 4:  f.read(4)
                    if meta_element_size >= 8:  fmt      = ru32(f)
                    if meta_element_size >= 12: play_off = ru32(f)
                    if meta_element_size >= 16: play_len = ru32(f)
                if meta_element_size < 24 and not play_len:
                    play_len = segments[last_segment][1]

            play_off += playregion_offset

            if version == 1:
                codec = fmt        & 0x01
                chans = (fmt >> 1) & 0x07
                rate  = (fmt >> 5) & 0x3FFFF
                align = (fmt >>23) & 0xFF
                bits  = (fmt >>31) & 0x01
            else:
                codec = fmt        & 0x03
                chans = (fmt >> 2) & 0x07
                rate  = (fmt >> 5) & 0x3FFFF
                align = (fmt >>23) & 0xFF
                bits  = (fmt >>31) & 0x01

            if play_len == 0:
                continue

            hex_name  = f"{entry_idx:08x}"
            base_name = track_names_map.get(hex_name, hex_name) if track_names_map else hex_name
            ext       = ".wma" if codec == CODEC_WMA else ".wav"
            out_path  = os.path.join(out_dir, base_name + ext)

            f.seek(play_off)
            with open(out_path, "wb") as fout:
                if codec != CODEC_WMA:
                    hdr = make_wav_header(codec, chans, rate, bits, align, play_len)
                    if hdr:
                        fout.write(hdr)
                copy_bytes(f, fout, play_len)

            output_files.append(out_path)

    return output_files




# ── XWB inject helpers ───────────────────────────────────────────────────────

def _parse_xwb_tracks(xwb_path):
    """Return a list of track dicts: index, offset, size, codec, duration."""
    tracks = []
    with open(xwb_path, "rb") as f:
        sig = f.read(4)
        if sig == SIGN_LE:
            ru32 = ru32_le
        elif sig == SIGN_BE:
            ru32 = ru32_be
        else:
            raise ValueError("Not a valid XWB file")

        version      = ru32(f)
        segments     = []
        last_segment = 3 if version <= 3 else 4
        if version >= 42:
            f.read(4)
        for _ in range(last_segment + 1):
            segments.append((ru32(f), ru32(f)))
        while len(segments) < 5:
            segments.append((0, 0))

        bank_data_off = f.tell() if version == 1 else segments[0][0]
        f.seek(bank_data_off)
        flags         = ru32(f)
        entry_count   = ru32(f)
        is_compact    = bool(flags & 0x00020000)
        f.read(16 if version in (2, 3) else 64)

        if version == 1:
            wavebank_offset   = f.tell()
            meta_element_size = 20
            alignment         = 4
            compact_format    = 0
        else:
            meta_element_size = ru32(f)
            ru32(f)
            alignment         = ru32(f)
            wavebank_offset   = segments[1][0]
            compact_format    = ru32(f) if is_compact else 0

        playregion_offset = segments[last_segment][0]
        if not playregion_offset:
            playregion_offset = wavebank_offset + (entry_count * meta_element_size)

        CODEC_NAMES = {0: "PCM", 1: "XMA", 2: "ADPCM", 3: "WMA"}

        for entry_idx in range(entry_count):
            ep = wavebank_offset + entry_idx * meta_element_size
            f.seek(ep)

            if is_compact:
                raw_val  = ru32(f)
                fmt      = compact_format
                play_off = (raw_val & 0x1FFFFF) * alignment
                if entry_idx == entry_count - 1:
                    play_len = segments[last_segment][1] - play_off
                else:
                    f.seek(ep + meta_element_size)
                    play_len = (ru32(f) & 0x1FFFFF) * alignment - play_off
            else:
                fmt = play_off = play_len = 0
                if version == 1:
                    fmt      = ru32(f)
                    play_off = ru32(f)
                    play_len = ru32(f)
                else:
                    if meta_element_size >= 4:  f.read(4)
                    if meta_element_size >= 8:  fmt      = ru32(f)
                    if meta_element_size >= 12: play_off = ru32(f)
                    if meta_element_size >= 16: play_len = ru32(f)
                if meta_element_size < 24 and not play_len:
                    play_len = segments[last_segment][1]

            play_off += playregion_offset

            if version == 1:
                codec = fmt        & 0x01
                chans = (fmt >> 1) & 0x07
                rate  = (fmt >> 5) & 0x3FFFF
                align = (fmt >>23) & 0xFF
                bits  = (fmt >>31) & 0x01
            else:
                codec = fmt        & 0x03
                chans = (fmt >> 2) & 0x07
                rate  = (fmt >> 5) & 0x3FFFF
                align = (fmt >>23) & 0xFF
                bits  = (fmt >>31) & 0x01

            if play_len == 0:
                continue

            # Calculate duration from size + format
            try:
                if codec == 0:  # PCM
                    bps = 8 << bits
                    dur = play_len / (rate * (bps / 8) * max(chans, 1)) if rate else 0
                elif codec == 2:  # ADPCM
                    block_align = (align + ADPCM_BLOCKALIGN_OFFSET) * max(chans, 1)
                    dur = (play_len / block_align) * (((block_align // max(chans,1) - 7)*2)+2) / max(rate, 1) if rate and block_align else 0
                else:
                    dur = 0
            except Exception:
                dur = 0

            tracks.append({
                "index":   entry_idx,
                "offset":  play_off,
                "size":    play_len,
                "codec":   CODEC_NAMES.get(codec, "???"),
                "codec_id": codec,
                "chans":   chans,
                "rate":    rate,
                "bits":    bits,
                "align":   align,
                "duration": dur,
                # store full format word and meta for rebuild
                "fmt_word": fmt,
                "ep":       wavebank_offset + entry_idx * meta_element_size,
            })

    return tracks


def _extract_single_track(xwb_path, track, out_path):
    """Extract one track to a WAV file (for preview)."""
    with open(xwb_path, "rb") as f:
        f.seek(track["offset"])
        data = f.read(track["size"])
    hdr = make_wav_header(
        track["codec_id"], track["chans"], track["rate"],
        track["bits"], track["align"], track["size"]
    )
    with open(out_path, "wb") as fout:
        if hdr:
            fout.write(hdr)
        fout.write(data)


def _strip_wav_header(wav_path):
    """Read a WAV file and return only the raw audio data (no header)."""
    with open(wav_path, "rb") as f:
        # Find the data chunk
        sig = f.read(4)
        if sig != b"RIFF":
            # Not a WAV — return raw bytes as-is
            f.seek(0)
            return f.read()
        f.read(4)  # RIFF size
        f.read(4)  # WAVE
        while True:
            chunk_id   = f.read(4)
            if not chunk_id or len(chunk_id) < 4:
                break
            chunk_size = struct.unpack("<I", f.read(4))[0]
            if chunk_id == b"data":
                return f.read(chunk_size)
            f.read(chunk_size)
    return b""


def _rebuild_xwb(src_path, replace_idx, wav_path, out_path):
    """
    Rebuild an XWB file, replacing track replace_idx with the audio
    from wav_path. All offsets are recalculated from scratch.
    """
    # Read entire source XWB
    with open(src_path, "rb") as f:
        src = f.read()

    # Get raw replacement audio
    new_audio = _strip_wav_header(wav_path)
    if not new_audio:
        raise ValueError("Could not read audio data from WAV file")

    # Re-parse to get track table positions
    tracks = _parse_xwb_tracks(src_path)
    if not tracks:
        raise ValueError("No tracks found in XWB")

    # Determine version and endianness
    sig = src[0:4]
    if sig == SIGN_LE:
        ru32 = lambda data, off: struct.unpack_from("<I", data, off)[0]
        wu32 = lambda v: struct.pack("<I", v)
    else:
        ru32 = lambda data, off: struct.unpack_from(">I", data, off)[0]
        wu32 = lambda v: struct.pack(">I", v)

    version = ru32(src, 4)
    last_segment = 3 if version <= 3 else 4
    hdr_offset = 8  # after sig + version
    if version >= 42:
        hdr_offset += 4

    # Read segment table offsets
    seg_offsets = []
    seg_lengths = []
    for i in range(last_segment + 1):
        o = hdr_offset + i * 8
        seg_offsets.append(ru32(src, o))
        seg_lengths.append(ru32(src, o + 4))

    # The wave data segment starts at seg_offsets[last_segment]
    wave_data_start = seg_offsets[last_segment]

    # Build new wave data block with replacement
    # Collect all track audio blobs in order
    audio_blobs = []
    for t in tracks:
        if t["index"] == replace_idx:
            blob = new_audio
        else:
            blob = src[t["offset"]: t["offset"] + t["size"]]
        audio_blobs.append(blob)

    # Build new output: copy everything up to wave data start, then patch
    out = bytearray(src[:wave_data_start])

    # Calculate new offsets relative to wave_data_start
    new_offsets = []
    current    = 0
    for blob in audio_blobs:
        new_offsets.append(current)
        current += len(blob)

    new_wave_data_size = current

    # Patch each track entry's offset and size in the metadata table
    meta_start = seg_offsets[1] if len(seg_offsets) > 1 else seg_offsets[0]

    # Read meta element size
    bank_data_off = seg_offsets[0]
    flags_val     = ru32(src, bank_data_off)
    is_compact    = bool(flags_val & 0x00020000)

    name_field_size = 16 if version in (2, 3) else 64
    meta_size_off   = bank_data_off + 8 + name_field_size
    meta_element_size = ru32(src, meta_size_off) if version != 1 else 20

    for i, t in enumerate(tracks):
        ep = t["ep"]
        # In non-compact non-version1: layout is flags(4)+fmt(4)+playoff(4)+playlen(4)+...
        if version == 1:
            play_off_field = ep + 4
            play_len_field = ep + 8
        else:
            play_off_field = ep + 8
            play_len_field = ep + 12

        new_off = wave_data_start + new_offsets[i]
        new_len = len(audio_blobs[i])

        out[play_off_field: play_off_field + 4] = wu32(new_off - wave_data_start)
        out[play_len_field: play_len_field + 4] = wu32(new_len)

    # Patch segment table: wave data length
    wave_len_field = hdr_offset + last_segment * 8 + 4
    out[wave_len_field: wave_len_field + 4] = wu32(new_wave_data_size)

    # Append new wave data
    for blob in audio_blobs:
        out += blob

    with open(out_path, "wb") as f:
        f.write(out)



# ── WAV → XWB creation ───────────────────────────────────────────────────────

def _parse_wav_info(wav_path):
    """Read WAV header and return audio parameters + raw audio bytes."""
    with open(wav_path, "rb") as f:
        sig = f.read(4)
        if sig != b"RIFF":
            raise ValueError(f"{os.path.basename(wav_path)} is not a valid WAV file")
        f.read(4)  # RIFF size
        wave = f.read(4)
        if wave != b"WAVE":
            raise ValueError(f"{os.path.basename(wav_path)} is not a valid WAV file")

        fmt_tag = channels = sample_rate = bits_per_sample = block_align = 0
        audio_data = b""

        while True:
            chunk_id = f.read(4)
            if not chunk_id or len(chunk_id) < 4:
                break
            chunk_size = struct.unpack("<I", f.read(4))[0]
            if chunk_id == b"fmt ":
                fmt_tag        = struct.unpack("<H", f.read(2))[0]
                channels       = struct.unpack("<H", f.read(2))[0]
                sample_rate    = struct.unpack("<I", f.read(4))[0]
                f.read(4)  # avg bytes per sec
                block_align    = struct.unpack("<H", f.read(2))[0]
                bits_per_sample= struct.unpack("<H", f.read(2))[0]
                remaining = chunk_size - 16
                if remaining > 0:
                    f.read(remaining)
            elif chunk_id == b"data":
                audio_data = f.read(chunk_size)
            else:
                f.read(chunk_size)

    if not audio_data:
        raise ValueError(f"No audio data in {os.path.basename(wav_path)}")
    if bits_per_sample not in (8, 16):
        raise ValueError(f"{os.path.basename(wav_path)}: only 8-bit and 16-bit PCM supported")

    return {
        "path":           wav_path,
        "channels":       channels,
        "sample_rate":    sample_rate,
        "bits_per_sample":bits_per_sample,
        "block_align":    block_align,
        "data":           audio_data,
    }


def create_xwb(wav_paths, out_path, bank_name="CustomBank"):
    """Bundle a list of WAV files into a new XWB wave bank (PCM only)."""
    if not wav_paths:
        raise ValueError("No WAV files provided")

    tracks = []
    for p in wav_paths:
        tracks.append(_parse_wav_info(p))

    num_tracks        = len(tracks)
    meta_element_size = 24   # standard WAVEBANKENTRY size

    # Layout
    hdr_size      = 4 + 4 + 4 + 5 * 8   # sig + version + hdrver + 5 segments = 52
    bankdata_size = 4 + 4 + 64 + 4 + 4 + 4 + 4 + 4  # = 92
    meta_size     = num_tracks * meta_element_size

    bankdata_off  = hdr_size                        # 52
    meta_off      = bankdata_off + bankdata_size     # 144
    wave_off_raw  = meta_off + meta_size
    wave_off      = (wave_off_raw + 3) & ~3          # align to 4 bytes

    audio_blobs   = [t["data"] for t in tracks]
    audio_offsets = []
    cur = 0
    for b in audio_blobs:
        audio_offsets.append(cur)
        cur += len(b)
    total_audio = cur

    out = bytearray()

    # WAVEBANKHEADER
    out += b"WBND"                                           # LE signature
    out += struct.pack("<I", 43)                             # version
    out += struct.pack("<I", 1)                              # dwHeaderVersion
    out += struct.pack("<II", bankdata_off, bankdata_size)   # seg 0 BANKDATA
    out += struct.pack("<II", meta_off, meta_size)           # seg 1 ENTRYMETADATA
    out += struct.pack("<II", 0, 0)                          # seg 2 SEEKTABLES
    out += struct.pack("<II", 0, 0)                          # seg 3 ENTRYNAMES
    out += struct.pack("<II", wave_off, total_audio)         # seg 4 ENTRYWAVEDATA

    # WAVEBANKDATA
    out += struct.pack("<I", 0)                              # dwFlags
    out += struct.pack("<I", num_tracks)                     # dwEntryCount
    name_b = bank_name.encode("ascii")[:64].ljust(64, b"\x00")
    out += name_b                                            # szBankName[64]
    out += struct.pack("<I", meta_element_size)              # dwEntryMetaDataElementSize
    out += struct.pack("<I", 0)                              # dwEntryNameElementSize
    out += struct.pack("<I", 4)                              # dwAlignment
    out += struct.pack("<I", 0)                              # CompactFormat
    out += struct.pack("<I", 0)                              # BuildTime

    # WAVEBANKENTRY per track
    for i, t in enumerate(tracks):
        codec = 0   # PCM
        chans = t["channels"]
        rate  = t["sample_rate"]
        bits  = 1 if t["bits_per_sample"] == 16 else 0
        align = 0
        fmt   = (codec & 0x3) | ((chans & 0x7) << 2) | ((rate & 0x3FFFF) << 5) |                 ((align & 0xFF) << 23) | ((bits & 0x1) << 31)
        out += struct.pack("<I", 0)                          # dwFlagsAndDuration
        out += struct.pack("<I", fmt)                        # Format
        out += struct.pack("<I", audio_offsets[i])           # PlayRegion.dwOffset
        out += struct.pack("<I", len(audio_blobs[i]))        # PlayRegion.dwLength
        out += struct.pack("<I", 0)                          # LoopRegion.dwOffset
        out += struct.pack("<I", 0)                          # LoopRegion.dwLength

    # Pad to wave_off
    while len(out) < wave_off:
        out += b"\x00"

    # Audio data
    for blob in audio_blobs:
        out += blob

    with open(out_path, "wb") as f:
        f.write(out)

# ── GUI ──────────────────────────────────────────────────────────────────────

BG        = "#1a1a2e"   # main background — active tab
RECENT    = "#0e2233"   # recent folders panel background
TAB_INACT = "#120d1e"   # inactive tab background — dark purple
PANEL     = "#16213e"
ACCENT    = "#e94560"
TEXT      = "#eaeaea"
MUTED     = "#7a7a9a"
SUCCESS   = "#4ecca3"
WARNING   = "#f5a623"
TAB_FG_ACT   = TEXT     # active tab text
TAB_FG_INACT = "#6a5a8a"  # inactive tab text — muted purple

FONT     = ("Consolas", 10)
FONT_BIG = ("Consolas", 13, "bold")
FONT_SM  = ("Consolas", 9)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("XWB → WAV Extractor")
        self.resizable(True, True)
        self.configure(bg=BG)

        self.input_var  = tk.StringVar()
        self.output_var = tk.StringVar()
        self.config_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready.")

        self._stop_event     = threading.Event()
        self._running        = False
        self._track_names    = {}
        self._current_tab    = "extract"
        self._recent_folders = self._load_recent()

        self._build_ui()
        self._center()
        self._try_load_config_from_cwd()
        # Lock window size so it doesn't shrink when switching tabs
        self.update_idletasks()
        self.minsize(self.winfo_width(), self.winfo_height())

    # ── Layout ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Title accent bar
        tk.Frame(self, bg=ACCENT, height=4).pack(fill="x")

        # Header
        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=20, pady=(14, 4))
        tk.Label(header, text="XWB → WAV Extractor", font=FONT_BIG,
                 bg=BG, fg=TEXT).pack(side="left")
        tk.Label(header, text="for RE4 & XACT games", font=FONT_SM,
                 bg=BG, fg=MUTED).pack(side="left", padx=(10, 0), pady=(4, 0))
        tk.Button(header, text="?", font=FONT_BIG,
                  bg=PANEL, fg=MUTED, activebackground=ACCENT, activeforeground=TEXT,
                  relief="flat", bd=0, padx=10, pady=2,
                  cursor="hand2", command=self._show_help).pack(side="right")

        # ── Tab bar ──
        tab_bar = tk.Frame(self, bg=BG)
        tab_bar.pack(fill="x")

        self._tab_btns = {}
        for key, label in [("extract", "    Extract    "), ("convert", "    Convert    "), ("inject", "    Inject    ")]:
            btn = tk.Label(tab_bar, text=label, font=FONT_BIG,
                           cursor="hand2", pady=10, padx=8)
            btn.pack(side="left")
            btn.bind("<Button-1>", lambda e, k=key: self._switch_tab(k))
            self._tab_btns[key] = btn

        # Active tab underline — a thin accent bar that slides under active tab
        self._tab_underline = tk.Frame(self, bg=ACCENT, height=2)
        self._tab_underline.pack(fill="x", padx=0)

        # Separator under tab bar
        tk.Frame(self, bg=PANEL, height=1).pack(fill="x", padx=16, pady=(0, 8))

        # ── Status bar (always at bottom) ──
        status_bar = tk.Frame(self, bg=PANEL)
        status_bar.pack(fill="x", side="bottom")
        tk.Label(status_bar, textvariable=self.status_var, font=FONT_SM,
                 bg=PANEL, fg=MUTED, anchor="w").pack(side="left", padx=10, pady=4)

        # ── Tab content frames ──
        self._extract_frame  = tk.Frame(self, bg=BG)
        self._convert_frame = tk.Frame(self, bg=TAB_INACT)
        self._inject_frame   = tk.Frame(self, bg=TAB_INACT)

        self._build_extract_tab()
        self._build_convert_tab()
        self._build_inject_tab()

        # Show extract tab by default
        self._switch_tab("extract")

    def _show_help(self):
        win = tk.Toplevel(self)
        win.title("How to use")
        win.configure(bg=BG)
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()

        tk.Frame(win, bg=ACCENT, height=3).pack(fill="x")

        content = tk.Frame(win, bg=BG)
        content.pack(padx=24, pady=16)

        sections = [
            ("Extract",
             "1. Set your XWB folder and output folder.\n"
             "2. Optionally load a config.json for custom track names.\n"
             "3. Click Extract All to dump every track from every .xwb file.\n"
             "4. Tick Individual Extraction to browse a single XWB,\n"
             "   preview tracks, and extract only the ones you want."),
            ("Inject",
             "1. Browse and load an XWB file.\n"
             "2. Select a track from the list and preview it.\n"
             "3. Browse a replacement WAV file.\n"
             "4. Click Replace & Rebuild to overwrite or save to a new folder."),
            ("Convert",
             "1. Add WAV files or a whole folder.\n"
             "2. Set an output .xwb path and bank name.\n"
             "3. Click Convert to XWB to bundle them into a new wave bank."),
        ]

        for title, body in sections:
            tk.Label(content, text=title, font=FONT, bg=BG, fg=ACCENT,
                     anchor="w").pack(anchor="w", pady=(8, 2))
            tk.Label(content, text=body, font=FONT_SM, bg=BG, fg=TEXT,
                     justify="left", anchor="w").pack(anchor="w", padx=(8, 0))

        tk.Frame(content, bg=PANEL, height=1).pack(fill="x", pady=(14, 0))
        tk.Button(content, text="Close", font=FONT_SM,
                  bg=PANEL, fg=MUTED, activebackground=ACCENT, activeforeground=TEXT,
                  relief="flat", bd=0, padx=16, pady=6,
                  cursor="hand2", command=win.destroy).pack(pady=(10, 0))

        win.update_idletasks()
        x = self.winfo_x() + (self.winfo_width()  - win.winfo_width())  // 2
        y = self.winfo_y() + (self.winfo_height() - win.winfo_height()) // 2
        win.geometry(f"+{x}+{y}")

    def _switch_tab(self, tab):
        # Stop any playing previews before switching
        if WINSOUND_OK:
            import winsound
            winsound.PlaySound(None, winsound.SND_PURGE)
        if getattr(self, "_inject_is_playing", False):
            self._inject_is_playing = False
            self._inject_preview_btn.config(text="▶  Preview Selected", fg=MUTED)
            tmp = getattr(self, "_inject_temp_wav", None)
            if tmp:
                self.after(500, lambda: self._inject_cleanup_tmp(tmp))
        if getattr(self, "_extract_is_playing", False):
            self._extract_is_playing = False
            self._extract_preview_btn.config(text="▶  Preview", fg=MUTED)
            tmp = getattr(self, "_extract_temp_wav", None)
            if tmp:
                self.after(500, lambda: self._extract_cleanup_tmp(tmp))

        self._current_tab = tab
        self._extract_frame.pack_forget()
        self._convert_frame.pack_forget()
        self._inject_frame.pack_forget()
        if tab == "extract":
            self._extract_frame.pack(fill="both", expand=True)
        elif tab == "convert":
            self._convert_frame.pack(fill="both", expand=True)
        else:
            self._inject_frame.pack(fill="both", expand=True)
        for key, btn in self._tab_btns.items():
            if key == tab:
                btn.config(bg=BG, fg=TAB_FG_ACT)
            else:
                btn.config(bg=TAB_INACT, fg=TAB_FG_INACT)

    # ── Extract tab ───────────────────────────────────────────────────────────

    def _build_extract_tab(self):
        p = self._extract_frame

        # ── Two-column layout with draggable sash ──
        self._extract_pane = tk.PanedWindow(p, orient="horizontal",
                                             bg=PANEL, sashwidth=8,
                                             sashrelief="flat", bd=0,
                                             handlesize=0)
        self._extract_pane.pack(fill="both", expand=True)

        left = tk.Frame(self._extract_pane, bg=BG)
        self._extract_pane.add(left, minsize=380, stretch="always")

        self._extract_right = tk.Frame(self._extract_pane, bg=BG, width=280)
        self._extract_right.pack_propagate(False)
        # Right panel added to pane only when checkbox is ticked
        self._extract_divider = None  # not used with PanedWindow

        # ── LEFT: folder pickers ──
        self._folder_row(left, "XWB Folder (input):", self.input_var,
                         self._browse_input, hint="Folder containing your .xwb files")
        self._folder_row(left, "Output Folder:", self.output_var,
                         self._browse_output, hint="Where extracted WAV files will be saved")
        self._folder_row(left, "Config File (optional):", self.config_var,
                         self._browse_config,
                         hint="JSON file for custom track names  —  leave blank to skip",
                         is_file=True)

        # Recent folders
        recent_outer = tk.Frame(left, bg=RECENT)
        recent_outer.pack(fill="x", padx=20, pady=(6, 2))
        tk.Label(recent_outer, text="Recent folders:", font=FONT_SM,
                 bg=RECENT, fg=MUTED).pack(anchor="w", padx=8, pady=(4, 2))
        self._recent_frame = tk.Frame(recent_outer, bg=RECENT)
        self._recent_frame.pack(fill="x", padx=8, pady=(0, 6))
        self._refresh_recent_ui()

        tk.Frame(left, bg=PANEL, height=1).pack(fill="x", padx=16, pady=(6, 8))

        # Progress
        prog_frame = tk.Frame(left, bg=BG)
        prog_frame.pack(fill="x", padx=20, pady=(0, 4))
        self._prog_label = tk.Label(prog_frame, text="", font=FONT_SM, bg=BG, fg=MUTED)
        self._prog_label.pack(side="right")

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor=PANEL, background=ACCENT,
                        bordercolor=PANEL, lightcolor=ACCENT, darkcolor=ACCENT)
        self._progress = ttk.Progressbar(left, style="Custom.Horizontal.TProgressbar",
                                          mode="determinate")
        self._progress.pack(fill="x", padx=20, pady=(0, 8))

        # Log
        log_frame = tk.Frame(left, bg=PANEL, bd=0)
        log_frame.pack(fill="both", expand=True, padx=20, pady=(0, 8))
        self._log = tk.Text(log_frame, font=FONT_SM,
                            bg=PANEL, fg=TEXT, insertbackground=TEXT,
                            relief="flat", bd=8, state="disabled",
                            wrap="word", cursor="arrow")
        self._log.pack(side="left", fill="both", expand=True)
        scroll = tk.Scrollbar(log_frame, command=self._log.yview, bg=PANEL,
                              troughcolor=PANEL, relief="flat")
        scroll.pack(side="right", fill="y")
        self._log.configure(yscrollcommand=scroll.set)
        self._log.tag_configure("ok",      foreground=SUCCESS)
        self._log.tag_configure("skip",    foreground=WARNING)
        self._log.tag_configure("error",   foreground=ACCENT)
        self._log.tag_configure("info",    foreground=MUTED)
        self._log.tag_configure("heading", foreground=TEXT)

        # Buttons
        btn_frame = tk.Frame(left, bg=BG)
        btn_frame.pack(pady=(0, 10), padx=20, anchor="w")

        self._start_btn = tk.Button(btn_frame, text="▶  Extract All",
                                    font=FONT, bg=ACCENT, fg=TEXT,
                                    activebackground="#c73652", activeforeground=TEXT,
                                    relief="flat", bd=0, padx=14, pady=8,
                                    cursor="hand2", command=self._start)
        self._start_btn.pack(side="left", padx=(0, 6))

        self._stop_btn = tk.Button(btn_frame, text="■  Stop",
                                   font=FONT, bg=PANEL, fg=MUTED,
                                   activebackground="#2a2a4e", activeforeground=TEXT,
                                   relief="flat", bd=0, padx=14, pady=8,
                                   cursor="hand2", command=self._stop,
                                   state="disabled")
        self._stop_btn.pack(side="left", padx=(0, 6))

        self._open_btn = tk.Button(btn_frame, text="📂  Open Output",
                                   font=FONT, bg="#1a4a2e", fg=SUCCESS,
                                   activebackground="#0f3320", activeforeground=SUCCESS,
                                   relief="flat", bd=0, padx=14, pady=8,
                                   cursor="hand2", command=self._open_output)
        self._open_btn.pack(side="left")

        # Config template button
        bottom_left = tk.Frame(left, bg=BG)
        bottom_left.pack(fill="x", padx=10, pady=(0, 4))

        self._indiv_var = tk.BooleanVar(value=False)
        self._indiv_btn = tk.Button(bottom_left, text="⊞  Individual Extraction",
                                    font=FONT_SM, bg=PANEL, fg=MUTED,
                                    activebackground=PANEL, activeforeground=TEXT,
                                    relief="flat", bd=0, padx=10, pady=4,
                                    cursor="hand2", command=self._toggle_track_browser)
        self._indiv_btn.pack(side="left", padx=(10, 0))
        self._indiv_pulse_step = 0
        self._indiv_pulse_id   = None
        self._start_indiv_pulse()

        self._cfg_btn = tk.Button(bottom_left, text="⚙ config template",
                                  font=FONT_SM, bg=BG, fg=MUTED,
                                  activebackground=BG, activeforeground=TEXT,
                                  relief="flat", bd=0, padx=6, pady=2,
                                  cursor="hand2", command=self._create_config)
        self._cfg_btn.pack(side="right")

        # ── RIGHT: track browser (hidden by default) ──
        r = self._extract_right
        tk.Label(r, text="Track Browser", font=FONT,
                 bg=BG, fg=TEXT).pack(anchor="w", padx=12, pady=(14, 0))
        tk.Label(r, text="preview & selective extract",
                 font=FONT_SM, bg=BG, fg=MUTED).pack(anchor="w", padx=12, pady=(0, 6))

        # Single XWB picker
        xwb_pick = tk.Frame(r, bg=BG)
        xwb_pick.pack(fill="x", padx=12, pady=(0, 4))
        self._extract_xwb_var = tk.StringVar()
        tk.Entry(xwb_pick, textvariable=self._extract_xwb_var, font=FONT_SM,
                 bg=PANEL, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=4).pack(side="left", fill="x", expand=True, padx=(0, 4))
        tk.Button(xwb_pick, text="Browse", font=FONT_SM,
                  bg=PANEL, fg=MUTED, activebackground="#2a2a4e",
                  activeforeground=TEXT, relief="flat", bd=0,
                  padx=8, pady=4, cursor="hand2",
                  command=self._extract_browse_single_xwb).pack(side="left")

        # Track listbox — fills remaining vertical space
        list_outer = tk.Frame(r, bg=PANEL)
        list_outer.pack(fill="both", expand=True, padx=12, pady=(0, 4))
        self._extract_track_listbox = tk.Listbox(list_outer,
            font=FONT_SM, bg=PANEL, fg=TEXT,
            selectbackground=ACCENT, selectforeground=TEXT,
            relief="flat", bd=8, activestyle="none",
            highlightthickness=0, selectmode="extended")
        self._extract_track_listbox.pack(side="left", fill="both", expand=True)
        track_scroll = tk.Scrollbar(list_outer, command=self._extract_track_listbox.yview,
                                    bg=PANEL, troughcolor=PANEL, relief="flat")
        track_scroll.pack(side="right", fill="y")
        self._extract_track_listbox.configure(yscrollcommand=track_scroll.set)
        self._extract_track_listbox.bind("<<ListboxSelect>>", self._extract_on_track_select)

        # Preview button + status
        self._extract_preview_btn = tk.Button(r, text="▶  Preview",
                                              font=FONT_SM, bg=PANEL, fg=MUTED,
                                              activebackground="#2a2a4e", activeforeground=TEXT,
                                              relief="flat", bd=0, padx=12, pady=5,
                                              cursor="hand2",
                                              command=self._extract_preview,
                                              state="disabled")
        self._extract_preview_btn.pack(anchor="w", padx=12, pady=(0, 4))

        self._extract_sel_lbl = tk.Label(r, text="Load an XWB to browse",
                                         font=FONT_SM, bg=BG, fg=MUTED,
                                         wraplength=240, justify="left")
        self._extract_sel_lbl.pack(anchor="w", padx=12, pady=(0, 4))

        self._extract_sel_btn = tk.Button(r, text="▶  Extract Selected",
                                          font=FONT_SM, bg="#1a3a4e", fg=SUCCESS,
                                          activebackground="#0f2a38", activeforeground=SUCCESS,
                                          relief="flat", bd=0, padx=12, pady=6,
                                          cursor="hand2",
                                          command=self._extract_selected_tracks,
                                          state="disabled")
        self._extract_sel_btn.pack(anchor="w", padx=12, pady=(0, 6))

        # Rename field — only shown when exactly 1 track is selected
        self._extract_rename_frame = tk.Frame(r, bg=BG)
        tk.Label(self._extract_rename_frame, text="rename to:",
                 font=FONT_SM, bg=BG, fg=MUTED).pack(anchor="w")
        rename_row = tk.Frame(self._extract_rename_frame, bg=BG)
        rename_row.pack(fill="x")
        self._extract_rename_var = tk.StringVar()
        self._extract_rename_entry = tk.Entry(rename_row, textvariable=self._extract_rename_var,
                                              font=FONT_SM, bg=PANEL, fg=TEXT,
                                              insertbackground=TEXT, relief="flat", bd=4)
        self._extract_rename_entry.pack(side="left", fill="x", expand=True)
        tk.Label(rename_row, text=".wav", font=FONT_SM, bg=BG, fg=MUTED).pack(side="left", padx=(2, 0))

        # Track browser state
        self._extract_tracks      = []
        self._extract_xwb_path    = None
        self._extract_is_playing  = False
        self._extract_temp_wav    = None

    def _toggle_track_browser(self):
        self._indiv_var.set(not self._indiv_var.get())
        if self._indiv_var.get():
            # Stop pulse, lock to solid accent
            self._stop_indiv_pulse()
            self._indiv_btn.config(bg=ACCENT, fg=TEXT,
                                   activebackground="#c73652", activeforeground=TEXT,
                                   text="⊠  Individual Extraction")
            self._extract_pane.add(self._extract_right, minsize=220, width=280, stretch="never")
            self.update_idletasks()
            needed_w = self.winfo_reqwidth()
            current_h = self.winfo_height()
            if self.winfo_width() < needed_w:
                self.geometry(f"{needed_w}x{current_h}")
        else:
            # Resume pulse
            self._indiv_btn.config(text="⊞  Individual Extraction",
                                   activebackground=PANEL, activeforeground=TEXT)
            self._start_indiv_pulse()
            self._extract_pane.remove(self._extract_right)

    def _start_indiv_pulse(self):
        # Pulse between PANEL and a slightly brighter highlight
        PULSE_COLORS = [
            ("#1e2a3a", "#6a7a9a"),
            ("#243040", "#8a9aba"),
            ("#2a3850", "#aabada"),
            ("#2e3e58", "#c0caea"),
            ("#2a3850", "#aabada"),
            ("#243040", "#8a9aba"),
        ]
        def _step():
            if self._indiv_var.get():
                return
            bg, fg = PULSE_COLORS[self._indiv_pulse_step % len(PULSE_COLORS)]
            self._indiv_btn.config(bg=bg, fg=fg)
            self._indiv_pulse_step += 1
            self._indiv_pulse_id = self.after(120, _step)
        self._indiv_pulse_id = self.after(120, _step)

    def _stop_indiv_pulse(self):
        if self._indiv_pulse_id:
            self.after_cancel(self._indiv_pulse_id)
            self._indiv_pulse_id = None

    # ── Extract tab — track browser ───────────────────────────────────────────

    def _extract_browse_single_xwb(self):
        path = filedialog.askopenfilename(
            title="Select XWB file to browse",
            filetypes=[("XWB files", "*.xwb"), ("All files", "*.*")]
        )
        if not path:
            return
        self._extract_xwb_var.set(path)
        self._extract_xwb_path = path
        self._extract_load_tracks(path)

    def _extract_load_tracks(self, path):
        self._extract_track_listbox.delete(0, "end")
        self._extract_tracks = []
        self._extract_preview_btn.config(state="disabled")
        self._extract_sel_btn.config(state="disabled")
        self._extract_sel_lbl.config(text="Loading tracks...")

        def _load():
            try:
                tracks = _parse_xwb_tracks(path)
                self.after(0, lambda: self._extract_populate_tracks(tracks))
            except Exception as e:
                self.after(0, lambda: self._extract_sel_lbl.config(
                    text=f"Error: {e}", fg=ACCENT))

        threading.Thread(target=_load, daemon=True).start()

    def _extract_populate_tracks(self, tracks):
        self._extract_tracks = tracks
        self._extract_track_listbox.delete(0, "end")
        for t in tracks:
            dur     = t["duration"]
            dur_str = f"{int(dur//60)}:{int(dur%60):02d}" if dur else "?"
            size_kb = t["size"] / 1024
            self._extract_track_listbox.insert("end",
                f"  {t['index']:03d}   {dur_str:>6}   {size_kb:>7.0f} KB   {t['codec']}")
        self._extract_sel_lbl.config(
            text=f"{len(tracks)} tracks  —  Ctrl+click or drag to multi-select", fg=MUTED)

    def _extract_on_track_select(self, event):
        sel = self._extract_track_listbox.curselection()

        # Auto-stop any playing preview when selection changes
        if self._extract_is_playing:
            import winsound
            self._extract_is_playing = False
            winsound.PlaySound(None, winsound.SND_PURGE)
            self._extract_preview_btn.config(text="▶  Preview", fg=MUTED)
            tmp = self._extract_temp_wav
            if tmp:
                self.after(500, lambda: self._extract_cleanup_tmp(tmp))

        if not sel:
            self._extract_preview_btn.config(state="disabled")
            self._extract_sel_btn.config(state="disabled")
            self._extract_rename_frame.pack_forget()
            return

        # Preview always enabled as long as at least 1 track is selected
        self._extract_preview_btn.config(state="normal")
        count = len(sel)
        self._extract_sel_lbl.config(
            text=f"{count} track{'s' if count > 1 else ''} selected", fg=SUCCESS)
        self._extract_sel_btn.config(state="normal")
        # Show rename box only for single selection
        if count == 1:
            self._extract_rename_var.set("")
            self._extract_rename_frame.pack(anchor="w", padx=12, pady=(0, 12), fill="x")
        else:
            self._extract_rename_frame.pack_forget()

    def _extract_preview(self):
        if not WINSOUND_OK:
            self._log_line("Preview is only available on Windows.", "skip")
            return
        sel = self._extract_track_listbox.curselection()
        if not sel or not self._extract_xwb_path:
            return

        import winsound

        # Toggle stop if already playing
        if self._extract_is_playing:
            self._extract_is_playing = False
            winsound.PlaySound(None, winsound.SND_PURGE)
            self._extract_preview_btn.config(text="▶  Preview", fg=MUTED)
            tmp = self._extract_temp_wav
            if tmp:
                self.after(500, lambda: self._extract_cleanup_tmp(tmp))
            return

        self._extract_preview_btn.config(text="Loading...", state="disabled")
        t = self._extract_tracks[sel[0]]

        def _extract_and_play():
            tmp = None
            try:
                tmp = tempfile.mktemp(suffix=".wav")
                _extract_single_track(self._extract_xwb_path, t, tmp)
                self._extract_temp_wav  = tmp
                self._extract_is_playing = True

                def _play():
                    try:
                        winsound.PlaySound(tmp, winsound.SND_FILENAME | winsound.SND_ASYNC)
                        self._extract_preview_btn.config(
                            text="■  Stop", fg=ACCENT, state="normal")
                        dur_ms = int(t["duration"] * 1000) + 1000
                        self.after(dur_ms, lambda: self._extract_auto_stop(tmp))
                    except Exception:
                        self._extract_is_playing = False
                        self._extract_cleanup_tmp(tmp)
                        self._extract_preview_btn.config(
                            text="▶  Preview", fg=MUTED, state="normal")
                self.after(0, _play)
            except Exception:
                self._extract_is_playing = False
                self.after(0, lambda: self._extract_preview_btn.config(
                    text="▶  Preview", fg=MUTED, state="normal"))

        threading.Thread(target=_extract_and_play, daemon=True).start()

    def _extract_auto_stop(self, tmp):
        if self._extract_is_playing:
            self._extract_is_playing = False
            self._extract_preview_btn.config(text="▶  Preview", fg=MUTED)
            self.after(500, lambda: self._extract_cleanup_tmp(tmp))

    def _extract_cleanup_tmp(self, path):
        try:
            if path and os.path.exists(path):
                os.remove(path)
        except Exception:
            pass

    def _extract_selected_tracks(self):
        sel = self._extract_track_listbox.curselection()
        if not sel or not self._extract_xwb_path:
            self._log_line("No tracks selected — load an XWB and select tracks first.", "skip")
            return
        output_folder = self.output_var.get().strip()
        if not output_folder:
            self._log_line("No output folder set.", "skip")
            return

        tracks_to_extract = [self._extract_tracks[i] for i in sel]
        xwb_base = os.path.splitext(os.path.basename(self._extract_xwb_path))[0]
        out_dir  = os.path.join(output_folder, xwb_base)
        os.makedirs(out_dir, exist_ok=True)

        names_for_bank = self._track_names.get(xwb_base, {})
        custom_name    = self._extract_rename_var.get().strip() if len(sel) == 1 else ""

        self._start_btn.config(state="disabled")
        self._extract_sel_btn.config(state="disabled")
        self._log_clear()
        self._log_line(f"Extracting {len(tracks_to_extract)} selected track(s) from {xwb_base}.xwb...\n", "heading")

        def _work():
            ok = fail = 0
            for t in tracks_to_extract:
                hex_name  = f"{t['index']:08x}"
                if custom_name:
                    base_name = custom_name
                else:
                    base_name = names_for_bank.get(hex_name, hex_name)
                ext       = ".wma" if t["codec"] == "WMA" else ".wav"
                out_path  = os.path.join(out_dir, base_name + ext)
                try:
                    _extract_single_track(self._extract_xwb_path, t, out_path)
                    self.after(0, lambda n=base_name: self._log_line(f"  {n}{ext}  OK", "ok"))
                    ok += 1
                except Exception as e:
                    err = str(e)
                    self.after(0, lambda n=base_name, e=err: self._log_line(f"  {n}  FAILED  ({e})", "error"))
                    fail += 1

            self.after(0, lambda: self._log_line(
                f"\nDone!  {ok} extracted, {fail} failed.\nOutput: {out_dir}", "heading"))
            self.after(0, lambda: self.status_var.set(f"Done!  {ok} extracted, {fail} failed."))
            self.after(0, lambda: self._start_btn.config(state="normal"))
            self.after(0, lambda: self._extract_sel_btn.config(state="normal"))

        threading.Thread(target=_work, daemon=True).start()

    # ── Convert tab ───────────────────────────────────────────────────────────

    def _build_convert_tab(self):
        p = self._convert_frame

        tk.Label(p, text="WAV → XWB Converter", font=FONT_BIG,
                 bg=TAB_INACT, fg=TEXT).pack(anchor="w", padx=20, pady=(14, 2))
        tk.Label(p, text="Bundle WAV files into a new XWB bank  (PCM only)",
                 font=FONT_SM, bg=TAB_INACT, fg=MUTED).pack(anchor="w", padx=20, pady=(0, 8))
        tk.Frame(p, bg=PANEL, height=1).pack(fill="x", padx=16, pady=(0, 10))

        # File list area
        list_frame = tk.Frame(p, bg=PANEL)
        list_frame.pack(fill="both", expand=True, padx=20, pady=(0, 8))

        self._convert_listbox = tk.Listbox(list_frame,
            font=FONT_SM, bg=PANEL, fg=TEXT,
            selectbackground=ACCENT, selectforeground=TEXT,
            relief="flat", bd=8, activestyle="none",
            highlightthickness=0, selectmode="extended")
        self._convert_listbox.pack(side="left", fill="both", expand=True)
        list_scroll = tk.Scrollbar(list_frame, command=self._convert_listbox.yview,
                                   bg=PANEL, troughcolor=PANEL, relief="flat")
        list_scroll.pack(side="right", fill="y")
        self._convert_listbox.configure(yscrollcommand=list_scroll.set)

        # File buttons row
        file_btn_frame = tk.Frame(p, bg=TAB_INACT)
        file_btn_frame.pack(fill="x", padx=20, pady=(0, 8))

        tk.Button(file_btn_frame, text="+ Add WAV Files", font=FONT_SM,
                  bg=PANEL, fg=MUTED, activebackground="#2a2a4e",
                  activeforeground=TEXT, relief="flat", bd=0,
                  padx=12, pady=5, cursor="hand2",
                  command=self._convert_add_files).pack(side="left", padx=(0, 6))

        tk.Button(file_btn_frame, text="+ Add Folder", font=FONT_SM,
                  bg=PANEL, fg=MUTED, activebackground="#2a2a4e",
                  activeforeground=TEXT, relief="flat", bd=0,
                  padx=12, pady=5, cursor="hand2",
                  command=self._convert_add_folder).pack(side="left", padx=(0, 6))

        tk.Button(file_btn_frame, text="✕ Remove Selected", font=FONT_SM,
                  bg=PANEL, fg=MUTED, activebackground="#2a2a4e",
                  activeforeground=TEXT, relief="flat", bd=0,
                  padx=12, pady=5, cursor="hand2",
                  command=self._convert_remove).pack(side="left", padx=(0, 6))

        tk.Button(file_btn_frame, text="Clear All", font=FONT_SM,
                  bg=PANEL, fg=MUTED, activebackground="#2a2a4e",
                  activeforeground=TEXT, relief="flat", bd=0,
                  padx=12, pady=5, cursor="hand2",
                  command=self._convert_clear).pack(side="left")

        # Output row
        out_row = tk.Frame(p, bg=TAB_INACT)
        out_row.pack(fill="x", padx=20, pady=(0, 6))
        tk.Label(out_row, text="Output XWB:", font=FONT, bg=TAB_INACT,
                 fg=TEXT, width=13, anchor="w").pack(side="left")
        self._convert_out_var = tk.StringVar()
        tk.Entry(out_row, textvariable=self._convert_out_var, font=FONT_SM,
                 bg=PANEL, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=6, width=44).pack(side="left", padx=(0, 6))
        tk.Button(out_row, text="Browse", font=FONT_SM,
                  bg=PANEL, fg=MUTED, activebackground="#2a2a4e",
                  activeforeground=TEXT, relief="flat", bd=0,
                  padx=10, pady=4, cursor="hand2",
                  command=self._convert_browse_out).pack(side="left")

        # Bank name row
        name_row = tk.Frame(p, bg=TAB_INACT)
        name_row.pack(fill="x", padx=20, pady=(0, 10))
        tk.Label(name_row, text="Bank Name:", font=FONT, bg=TAB_INACT,
                 fg=TEXT, width=13, anchor="w").pack(side="left")
        self._convert_name_var = tk.StringVar(value="CustomBank")
        tk.Entry(name_row, textvariable=self._convert_name_var, font=FONT_SM,
                 bg=PANEL, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=6, width=24).pack(side="left")
        tk.Label(name_row, text="(name stored inside the XWB file)", font=FONT_SM,
                 bg=TAB_INACT, fg=MUTED).pack(side="left", padx=(8, 0))

        # Action buttons
        act_frame = tk.Frame(p, bg=TAB_INACT)
        act_frame.pack(pady=(0, 12))

        self._convert_btn = tk.Button(act_frame, text="⬇  Convert to XWB",
                                       font=FONT, bg=ACCENT, fg=TEXT,
                                       activebackground="#c73652", activeforeground=TEXT,
                                       relief="flat", bd=0, padx=20, pady=8,
                                       cursor="hand2", command=self._convert_run)
        self._convert_btn.pack(side="left", padx=(0, 8))

        self._convert_open_btn = tk.Button(act_frame, text="📂  Open Output Folder",
                                            font=FONT, bg="#1a4a2e", fg=SUCCESS,
                                            activebackground="#0f3320", activeforeground=SUCCESS,
                                            relief="flat", bd=0, padx=20, pady=8,
                                            cursor="hand2", command=self._convert_open_folder)
        self._convert_open_btn.pack(side="left")

        self._convert_status = tk.Label(p, text="", font=FONT_SM,
                                         bg=TAB_INACT, fg=SUCCESS)
        self._convert_status.pack()

        # Internal state
        self._convert_files = []

    def _convert_add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select WAV files",
            filetypes=[("WAV files", "*.wav"), ("All files", "*.*")]
        )
        for p in paths:
            if p not in self._convert_files:
                self._convert_files.append(p)
                self._convert_listbox.insert("end", f"  {os.path.basename(p)}")

    def _convert_add_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing WAV files")
        if not folder:
            return
        wavs = sorted(f for f in os.listdir(folder) if f.lower().endswith(".wav"))
        for fname in wavs:
            full = os.path.join(folder, fname)
            if full not in self._convert_files:
                self._convert_files.append(full)
                self._convert_listbox.insert("end", f"  {fname}")

    def _convert_remove(self):
        sel = list(self._convert_listbox.curselection())
        for i in reversed(sel):
            self._convert_listbox.delete(i)
            del self._convert_files[i]

    def _convert_clear(self):
        self._convert_listbox.delete(0, "end")
        self._convert_files.clear()

    def _convert_browse_out(self):
        path = filedialog.asksaveasfilename(
            title="Save XWB file as",
            defaultextension=".xwb",
            filetypes=[("XWB files", "*.xwb")]
        )
        if path:
            self._convert_out_var.set(path)

    def _convert_open_folder(self):
        out = self._convert_out_var.get().strip()
        folder = os.path.dirname(out) if out else ""
        if folder and os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showerror("Error", "No valid output path set.")

    def _convert_run(self):
        if not self._convert_files:
            messagebox.showerror("Error", "Add at least one WAV file first.")
            return
        out_path = self._convert_out_var.get().strip()
        if not out_path:
            messagebox.showerror("Error", "Please set an output XWB path.")
            return
        bank_name = self._convert_name_var.get().strip() or "CustomBank"

        self._convert_btn.config(state="disabled", text="Converting...")
        self._convert_status.config(text="", fg=MUTED)

        files = list(self._convert_files)

        def _work():
            try:
                create_xwb(files, out_path, bank_name)
                name = os.path.basename(out_path)
                self.after(0, lambda: self._convert_status.config(
                    text=f"Done! Saved: {name}", fg=SUCCESS))
            except Exception as e:
                self.after(0, lambda: self._convert_status.config(
                    text=f"Error: {e}", fg=ACCENT))
            finally:
                self.after(0, lambda: self._convert_btn.config(
                    state="normal", text="⬇  Convert to XWB"))

        threading.Thread(target=_work, daemon=True).start()

    # ── Inject tab ────────────────────────────────────────────────────────────

    def _build_inject_tab(self):
        p = self._inject_frame

        top = tk.Frame(p, bg=TAB_INACT)
        top.pack(fill="x", padx=16, pady=(14, 6))
        tk.Label(top, text="XWB File:", font=FONT, bg=TAB_INACT,
                 fg=TEXT, width=10, anchor="w").pack(side="left")
        self._inject_xwb_var = tk.StringVar()
        tk.Entry(top, textvariable=self._inject_xwb_var, font=FONT_SM,
                 bg=PANEL, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=6, width=48).pack(side="left", padx=(0, 6))
        tk.Button(top, text="Browse", font=FONT_SM,
                  bg=PANEL, fg=MUTED, activebackground="#2a2a4e",
                  activeforeground=TEXT, relief="flat", bd=0,
                  padx=10, pady=4, cursor="hand2",
                  command=self._inject_browse_xwb).pack(side="left")

        tk.Frame(p, bg=PANEL, height=1).pack(fill="x", padx=16, pady=(4, 8))

        main = tk.Frame(p, bg=TAB_INACT)
        main.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        # Left: track list
        left = tk.Frame(main, bg=TAB_INACT)
        left.pack(side="left", fill="both", expand=True)

        tk.Label(left, text="Tracks in XWB", font=FONT, bg=TAB_INACT,
                 fg=TEXT).pack(anchor="w", pady=(0, 6))

        list_outer = tk.Frame(left, bg=PANEL)
        list_outer.pack(fill="both", expand=True)

        self._inject_listbox = tk.Listbox(list_outer,
            font=FONT_SM, bg=PANEL, fg=TEXT,
            selectbackground=ACCENT, selectforeground=TEXT,
            relief="flat", bd=0, activestyle="none",
            highlightthickness=0)
        self._inject_listbox.pack(side="left", fill="both", expand=True)
        list_scroll = tk.Scrollbar(list_outer, command=self._inject_listbox.yview,
                                   bg=PANEL, troughcolor=PANEL, relief="flat")
        list_scroll.pack(side="right", fill="y")
        self._inject_listbox.configure(yscrollcommand=list_scroll.set)
        self._inject_listbox.bind("<<ListboxSelect>>", self._inject_on_select)

        self._inject_preview_btn = tk.Button(left, text="▶  Preview Selected",
                                              font=FONT_SM, bg=PANEL, fg=MUTED,
                                              activebackground="#2a2a4e",
                                              activeforeground=TEXT,
                                              relief="flat", bd=0, padx=12, pady=6,
                                              cursor="hand2",
                                              command=self._inject_preview,
                                              state="disabled")
        self._inject_preview_btn.pack(anchor="w", pady=(6, 0))

        tk.Frame(main, bg=PANEL, width=1).pack(side="left", fill="y", padx=12)

        # Right: replace panel
        right = tk.Frame(main, bg=TAB_INACT, width=260)
        right.pack(side="left", fill="y")

        tk.Label(right, text="Replace With", font=FONT, bg=TAB_INACT,
                 fg=TEXT).pack(anchor="w", pady=(0, 6))

        self._inject_selected_lbl = tk.Label(right,
            text="No track selected", font=FONT_SM,
            bg=TAB_INACT, fg=MUTED, anchor="w", wraplength=220, justify="left")
        self._inject_selected_lbl.pack(anchor="w", pady=(0, 12))

        tk.Label(right, text="Custom WAV:", font=FONT_SM,
                 bg=TAB_INACT, fg=MUTED).pack(anchor="w")
        self._inject_wav_var = tk.StringVar()
        tk.Entry(right, textvariable=self._inject_wav_var,
                 font=FONT_SM, bg=PANEL, fg=TEXT,
                 insertbackground=TEXT, relief="flat", bd=6, width=26).pack(anchor="w", pady=(2, 4))
        tk.Button(right, text="Browse WAV", font=FONT_SM,
                  bg=PANEL, fg=MUTED, activebackground="#2a2a4e",
                  activeforeground=TEXT, relief="flat", bd=0,
                  padx=10, pady=4, cursor="hand2",
                  command=self._inject_browse_wav).pack(anchor="w")

        # Save to separate folder option
        self._inject_separate_var = tk.BooleanVar(value=False)
        tk.Checkbutton(right, text="Save to separate folder",
                       variable=self._inject_separate_var,
                       font=FONT_SM, bg=TAB_INACT, fg=TEXT,
                       selectcolor=TAB_INACT, activebackground=TAB_INACT,
                       activeforeground=TEXT,
                       command=self._inject_toggle_folder).pack(anchor="w", pady=(14, 2))

        self._inject_folder_frame = tk.Frame(right, bg=TAB_INACT)
        self._inject_out_var = tk.StringVar()
        tk.Entry(self._inject_folder_frame, textvariable=self._inject_out_var,
                 font=FONT_SM, bg=PANEL, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=4, width=24).pack(anchor="w", pady=(0, 2))
        folder_btns = tk.Frame(self._inject_folder_frame, bg=TAB_INACT)
        folder_btns.pack(anchor="w")
        tk.Button(folder_btns, text="Browse", font=FONT_SM,
                  bg=PANEL, fg=MUTED, activebackground="#2a2a4e",
                  activeforeground=TEXT, relief="flat", bd=0,
                  padx=8, pady=3, cursor="hand2",
                  command=self._inject_browse_out_folder).pack(side="left", padx=(0, 4))
        self._inject_open_folder_btn = tk.Button(folder_btns, text="📂",
                  font=FONT_SM, bg="#1a4a2e", fg=SUCCESS,
                  activebackground="#0f3320", activeforeground=SUCCESS,
                  relief="flat", bd=0, padx=8, pady=3,
                  cursor="hand2", command=self._inject_open_out_folder)
        self._inject_open_folder_btn.pack(side="left")
        # hidden by default
        # self._inject_folder_frame.pack(anchor="w")

        self._inject_replace_btn = tk.Button(right, text="⬇  Replace & Rebuild",
                                              font=FONT, bg=ACCENT, fg=TEXT,
                                              activebackground="#c73652",
                                              activeforeground=TEXT,
                                              relief="flat", bd=0,
                                              padx=16, pady=8,
                                              cursor="hand2",
                                              command=self._inject_replace,
                                              state="disabled")
        self._inject_replace_btn.pack(anchor="w", pady=(10, 0))

        self._inject_status_lbl = tk.Label(right, text="",
                                            font=FONT_SM, bg=TAB_INACT,
                                            fg=SUCCESS, wraplength=220,
                                            justify="left")
        self._inject_status_lbl.pack(anchor="w", pady=(8, 0))

        self._inject_tracks     = []
        self._inject_xwb_path   = None
        self._inject_selected   = None
        self._inject_is_playing = False

    def _inject_toggle_folder(self):
        if self._inject_separate_var.get():
            self._inject_folder_frame.pack(anchor="w", pady=(0, 4))
        else:
            self._inject_folder_frame.pack_forget()

    def _inject_browse_out_folder(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self._inject_out_var.set(path)

    def _inject_open_out_folder(self):
        path = self._inject_out_var.get().strip()
        if path and os.path.isdir(path):
            os.startfile(path)
        else:
            messagebox.showerror("Error", "No valid output folder selected.")

    def _inject_browse_xwb(self):
        path = filedialog.askopenfilename(
            title="Select XWB file to modify",
            filetypes=[("XWB files", "*.xwb"), ("All files", "*.*")]
        )
        if not path:
            return
        self._inject_xwb_var.set(path)
        self._inject_xwb_path = path
        self._inject_load_tracks(path)

    def _inject_load_tracks(self, path):
        self._inject_listbox.delete(0, "end")
        self._inject_tracks   = []
        self._inject_selected = None
        self._inject_preview_btn.config(state="disabled")
        self._inject_replace_btn.config(state="disabled")
        self._inject_selected_lbl.config(text="Loading...", fg=MUTED)

        def _load():
            try:
                tracks = _parse_xwb_tracks(path)
                self.after(0, lambda: self._inject_populate_list(tracks))
            except Exception as e:
                self.after(0, lambda: self._inject_selected_lbl.config(
                    text=f"Error: {e}", fg=ACCENT))

        threading.Thread(target=_load, daemon=True).start()

    def _inject_populate_list(self, tracks):
        self._inject_tracks = tracks
        self._inject_listbox.delete(0, "end")
        for t in tracks:
            dur     = t["duration"]
            dur_str = f"{int(dur//60)}:{int(dur%60):02d}" if dur else "?"
            size_kb = t["size"] / 1024
            self._inject_listbox.insert("end",
                f"  {t['index']:03d}   {dur_str:>6}   {size_kb:>7.0f} KB   {t['codec']}")
        self._inject_selected_lbl.config(
            text=f"{len(tracks)} tracks loaded.\nSelect one to replace.", fg=MUTED)

    def _inject_on_select(self, event):
        sel = self._inject_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(self._inject_tracks):
            return
        # Stop any currently playing preview when selecting a new track
        if getattr(self, "_inject_is_playing", False) and WINSOUND_OK:
            import winsound
            winsound.PlaySound(None, winsound.SND_ASYNC)
            self._inject_is_playing = False
            self._inject_preview_btn.config(text="▶  Preview Selected", fg=MUTED)
        self._inject_selected = self._inject_tracks[idx]
        t   = self._inject_selected
        dur = t["duration"]
        dur_str = f"{int(dur//60)}:{int(dur%60):02d}" if dur else "?"
        self._inject_selected_lbl.config(
            text=f"Track {t['index']:03d}\nDuration: {dur_str}\nSize: {t['size']//1024} KB\nCodec: {t['codec']}",
            fg=TEXT)
        self._inject_preview_btn.config(state="normal")
        if self._inject_wav_var.get():
            self._inject_replace_btn.config(state="normal")
            self._start_inject_replace_pulse()

    def _inject_browse_wav(self):
        path = filedialog.askopenfilename(
            title="Select replacement WAV file",
            filetypes=[("WAV files", "*.wav"), ("All files", "*.*")]
        )
        if not path:
            return
        self._inject_wav_var.set(path)
        if self._inject_selected is not None:
            self._inject_replace_btn.config(state="normal")
            self._start_inject_replace_pulse()

    def _inject_preview(self):
        if not self._inject_selected or not self._inject_xwb_path:
            return
        if not WINSOUND_OK:
            self._inject_status_lbl.config(text="Preview only available on Windows.", fg=WARNING)
            return

        import winsound

        # If already playing — stop and reset
        if getattr(self, "_inject_is_playing", False):
            self._inject_is_playing = False
            winsound.PlaySound(None, winsound.SND_PURGE)
            self._inject_preview_btn.config(text="▶  Preview Selected", fg=MUTED)
            tmp = getattr(self, "_inject_temp_wav", None)
            if tmp:
                self.after(500, lambda: self._inject_cleanup_tmp(tmp))
            return

        self._inject_preview_btn.config(text="Loading...", state="disabled")
        t = self._inject_selected

        def _extract_and_play():
            tmp = None
            try:
                tmp = tempfile.mktemp(suffix=".wav")
                _extract_single_track(self._inject_xwb_path, t, tmp)
                self._inject_temp_wav  = tmp
                self._inject_is_playing = True

                def _play():
                    try:
                        winsound.PlaySound(tmp, winsound.SND_FILENAME | winsound.SND_ASYNC)
                        self._inject_preview_btn.config(
                            text="■  Stop", fg=ACCENT, state="normal")
                        # Auto-reset button after track duration + 1s buffer
                        dur_ms = int(t["duration"] * 1000) + 1000
                        self.after(dur_ms, lambda: self._inject_auto_stop(tmp))
                    except Exception:
                        self._inject_is_playing = False
                        self._inject_cleanup_tmp(tmp)
                        self._inject_preview_btn.config(
                            text="▶  Preview Selected", fg=MUTED, state="normal")
                self.after(0, _play)
            except Exception:
                self._inject_is_playing = False
                self.after(0, lambda: self._inject_preview_btn.config(
                    text="▶  Preview Selected", fg=MUTED, state="normal"))

        threading.Thread(target=_extract_and_play, daemon=True).start()

    def _inject_auto_stop(self, tmp):
        """Called when the track duration has elapsed — reset button if still playing."""
        if getattr(self, "_inject_is_playing", False):
            self._inject_is_playing = False
            self._inject_preview_btn.config(text="▶  Preview Selected", fg=MUTED)
            self.after(500, lambda: self._inject_cleanup_tmp(tmp))

    def _inject_cleanup_tmp(self, path):
        try:
            if path and os.path.exists(path):
                os.remove(path)
        except Exception:
            pass

    def _inject_replace(self):
        if not self._inject_selected or not self._inject_xwb_path:
            self._inject_status_lbl.config(text="No track selected.", fg=WARNING)
            return
        wav_path = self._inject_wav_var.get().strip()
        if not wav_path or not os.path.isfile(wav_path):
            self._inject_status_lbl.config(text="No valid WAV file selected.", fg=WARNING)
            return

        xwb_fname = os.path.basename(self._inject_xwb_path)
        if self._inject_separate_var.get():
            out_folder = self._inject_out_var.get().strip()
            if not out_folder or not os.path.isdir(out_folder):
                self._inject_status_lbl.config(text="No valid output folder selected.", fg=WARNING)
                return
            out_path = os.path.join(out_folder, xwb_fname)
        else:
            out_path = self._inject_xwb_path

        self._stop_inject_replace_pulse()
        self._inject_replace_btn.config(state="disabled", text="Working...")
        self._inject_status_lbl.config(text="Rebuilding XWB...", fg=MUTED)

        track_idx = self._inject_selected["index"]

        def _work():
            try:
                _rebuild_xwb(self._inject_xwb_path, track_idx, wav_path, out_path)
                if self._inject_separate_var.get():
                    name = os.path.basename(out_path)
                    msg  = f"Saved to:\n{name}"
                else:
                    msg = "Original overwritten!"
                self.after(0, lambda m=msg: self._inject_status_lbl.config(
                    text=m, fg=SUCCESS))
            except Exception as e:
                self.after(0, lambda: self._inject_status_lbl.config(
                    text=f"Error: {e}", fg=ACCENT))
            finally:
                self.after(0, lambda: self._inject_replace_btn.config(
                    state="normal", text="⬇  Replace & Rebuild"))
                self.after(0, self._start_inject_replace_pulse)

        threading.Thread(target=_work, daemon=True).start()

    def _start_inject_replace_pulse(self):
        self._stop_inject_replace_pulse()
        PULSE = ["#e94560", "#f05575", "#ff6585", "#f05575", "#e94560", "#d03550"]
        def _step():
            if not hasattr(self, "_inject_replace_pulse_step"):
                self._inject_replace_pulse_step = 0
            try:
                state = self._inject_replace_btn.cget("state")
                if str(state) == "disabled":
                    return
                color = PULSE[self._inject_replace_pulse_step % len(PULSE)]
                self._inject_replace_btn.config(bg=color)
                self._inject_replace_pulse_step += 1
                self._inject_replace_pulse_id = self.after(130, _step)
            except Exception:
                pass
        self._inject_replace_pulse_step = 0
        self._inject_replace_pulse_id = self.after(130, _step)

    def _stop_inject_replace_pulse(self):
        pid = getattr(self, "_inject_replace_pulse_id", None)
        if pid:
            self.after_cancel(pid)
            self._inject_replace_pulse_id = None
        try:
            self._inject_replace_btn.config(bg=ACCENT)
        except Exception:
            pass

    # ── Folder row ────────────────────────────────────────────────────────────

    def _folder_row(self, parent, label, var, command, hint="", is_file=False):
        frame = tk.Frame(parent, bg=BG)
        frame.pack(fill="x", padx=20, pady=3)
        tk.Label(frame, text=label, font=FONT, bg=BG, fg=TEXT,
                 width=24, anchor="w").pack(side="left")
        entry = tk.Entry(frame, textvariable=var, font=FONT_SM,
                         bg=PANEL, fg=TEXT, insertbackground=TEXT,
                         relief="flat", bd=6, width=40)
        entry.pack(side="left", padx=(0, 6))
        tk.Button(frame, text="Browse", font=FONT_SM,
                  bg=PANEL, fg=MUTED, activebackground="#2a2a4e",
                  activeforeground=TEXT, relief="flat", bd=0,
                  padx=10, pady=4, cursor="hand2",
                  command=command).pack(side="left")
        if hint:
            tk.Label(frame, text=hint, font=FONT_SM, bg=BG, fg=MUTED).pack(side="left", padx=(8, 0))

    # ── Browse callbacks ─────────────────────────────────────────────────────

    # ── Recent folders ───────────────────────────────────────────────────────

    def _load_recent(self):
        try:
            p = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), RECENT_FILENAME)
            if os.path.exists(p):
                with open(p) as f:
                    return json.load(f)
        except Exception:
            pass
        return []

    def _save_recent(self, path):
        if not path or not os.path.isdir(path):
            return
        if path in self._recent_folders:
            self._recent_folders.remove(path)
        self._recent_folders.insert(0, path)
        self._recent_folders = self._recent_folders[:MAX_RECENT]
        try:
            p = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), RECENT_FILENAME)
            with open(p, "w") as f:
                json.dump(self._recent_folders, f)
        except Exception:
            pass
        self._refresh_recent_ui()

    def _refresh_recent_ui(self):
        for w in self._recent_frame.winfo_children():
            w.destroy()
        if not self._recent_folders:
            tk.Label(self._recent_frame, text="No recent folders yet.",
                     font=FONT_SM, bg=RECENT, fg=MUTED).pack(anchor="w")
            return
        for path in self._recent_folders:
            short = path if len(path) <= 55 else "..." + path[-52:]
            btn = tk.Button(self._recent_frame, text=short,
                            font=FONT_SM, bg=RECENT, fg=SUCCESS,
                            activebackground="#0a1a28", activeforeground=SUCCESS,
                            relief="flat", bd=0, anchor="w", cursor="hand2",
                            command=lambda p=path: self._use_recent(p))
            btn.pack(fill="x", pady=1)

    def _use_recent(self, path):
        self.input_var.set(path)
        self._log_line(f"Recent folder selected: {path}", "info")

    def _open_output(self):
        path = self.output_var.get().strip()
        if path and os.path.isdir(path):
            os.startfile(path)
        else:
            messagebox.showerror("Error", "No valid output folder selected.")

    def _browse_input(self):
        path = filedialog.askdirectory(title="Select folder containing .xwb files")
        if path:
            self.input_var.set(path)

    def _browse_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_var.set(path)

    def _browse_config(self):
        path = filedialog.askopenfilename(
            title="Select config.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if path:
            self.config_var.set(path)
            self._load_config(path)

    # ── Config ───────────────────────────────────────────────────────────────

    def _try_load_config_from_cwd(self):
        local = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), CONFIG_FILENAME)
        if os.path.exists(local):
            self.config_var.set(local)
            self._load_config(local)

    def _load_config(self, path):
        try:
            with open(path, "r") as f:
                data = json.load(f)
            self._track_names = data.get("track_names", {})
            self._log_line(f"Config loaded: {os.path.basename(path)}", "info")
            total = sum(len(v) for v in self._track_names.values())
            self._log_line(f"  {total} custom track name(s) found across {len(self._track_names)} bank(s)", "info")
        except Exception as e:
            self._log_line(f"Could not load config: {e}", "skip")
            self._track_names = {}

    def _create_config(self):
        path = filedialog.asksaveasfilename(
            title="Save config template as...",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialfile="config.json"
        )
        if not path:
            return
        try:
            with open(path, "w") as f:
                json.dump(DEFAULT_CONFIG, f, indent=4)
            self.config_var.set(path)
            self._log_line(f"Config template created: {path}", "ok")
            self._log_line("  Open it in any text editor to add custom track names.", "info")
            self._log_line('  Under "track_names", add entries like:', "info")
            self._log_line('    "bio4bgm": { "00000000": "main_menu", "00000001": "village" }', "info")
        except Exception as e:
            self._log_line(f"Failed to create config: {e}", "error")

    # ── Extraction ───────────────────────────────────────────────────────────

    def _start(self):
        input_folder  = self.input_var.get().strip()
        output_folder = self.output_var.get().strip()

        if not input_folder or not os.path.isdir(input_folder):
            self._log_line("No valid XWB input folder selected.", "skip")
            return
        if not output_folder:
            self._log_line("No output folder set.", "skip")
            return

        xwb_files = sorted(f for f in os.listdir(input_folder) if f.lower().endswith(".xwb"))
        if not xwb_files:
            self._log_line(f"No .xwb files found in: {input_folder}", "skip")
            return

        self._running = True
        self._stop_event.clear()
        self._start_btn.config(state="disabled")
        self._stop_btn.config(state="normal")
        self._progress["value"] = 0
        self._progress["maximum"] = len(xwb_files)
        self._log_clear()

        config_path = self.config_var.get().strip()
        if config_path and os.path.isfile(config_path):
            self._load_config(config_path)

        self._save_recent(input_folder)

        thread = threading.Thread(
            target=self._run,
            args=(input_folder, output_folder, xwb_files),
            daemon=True
        )
        thread.start()

    def _stop(self):
        self._stop_event.set()
        self.status_var.set("Stopping after current file...")

    def _run(self, input_folder, output_folder, xwb_files):
        total      = len(xwb_files)
        ok_count   = 0
        fail_count = 0
        ok_lines   = []
        fail_lines = []

        self._log_line(f"Starting extraction of {total} XWB files...\n", "heading")

        for i, fname in enumerate(xwb_files, 1):
            if self._stop_event.is_set():
                self._log_line("\nStopped by user.", "skip")
                break

            xwb_path = os.path.join(input_folder, fname)
            base     = os.path.splitext(fname)[0]
            out_dir  = os.path.join(output_folder, base)
            size_mb  = os.path.getsize(xwb_path) / (1024 * 1024)

            self._set_status(f"Processing {i}/{total}: {fname}")
            self._log_line(f"[{i}/{total}]  {fname}  ({size_mb:.1f} MB) ... ", "info", newline=False)

            names_for_bank = self._track_names.get(base, {})

            try:
                extracted = extract_xwb(xwb_path, out_dir,
                                        track_names_map=names_for_bank,
                                        stop_event=self._stop_event)
                if extracted:
                    self._log_line(f"OK  ({len(extracted)} tracks)", "ok")
                    ok_count += 1
                    ok_lines.append(fname)
                else:
                    self._log_line("SKIPPED  (no tracks found)", "skip")
                    fail_count += 1
                    fail_lines.append(f"{fname} - no tracks found")
            except Exception as e:
                self._log_line(f"SKIPPED  ({e})", "skip")
                fail_count += 1
                fail_lines.append(f"{fname} - {e}")

            self._set_progress(i)

        try:
            with open(os.path.join(output_folder, "converted_ok.txt"), "w") as f:
                f.write("\n".join(ok_lines))
            with open(os.path.join(output_folder, "failed.txt"), "w") as f:
                f.write("\n".join(fail_lines))
        except Exception:
            pass

        self._log_line(f"\n{'─'*50}", "info")
        self._log_line(f"Done!  Extracted: {ok_count}   Skipped: {fail_count}", "heading")
        self._log_line(f"Output folder: {output_folder}", "info")
        if fail_count:
            self._log_line(f"Check failed.txt for details on skipped files.", "skip")

        self._set_status(f"Done!  {ok_count} extracted, {fail_count} skipped.")
        self._finish()

    # ── UI helpers ───────────────────────────────────────────────────────────

    def _log_clear(self):
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

    def _log_line(self, text, tag="", newline=True):
        self._log.config(state="normal")
        self._log.insert("end", text + ("\n" if newline else ""), tag)
        self._log.see("end")
        self._log.config(state="disabled")

    def _set_status(self, text):
        self.after(0, lambda: self.status_var.set(text))

    def _set_progress(self, value):
        def _update():
            self._progress["value"] = value
            total = self._progress["maximum"]
            self._prog_label.config(text=f"{value} / {total}")
        self.after(0, _update)

    def _finish(self):
        def _update():
            self._running = False
            self._start_btn.config(state="normal")
            self._stop_btn.config(state="disabled")
        self.after(0, _update)

    def _center(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")


# ── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
