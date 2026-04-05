"""
XWB to WAV Extractor - GUI Version
Built with tkinter (included with Python, no extra installs needed)

To build as a standalone .exe:
    pip install pyinstaller
    pyinstaller --onefile --noconsole --name "XWB Extractor" xwb_extractor.py
"""

import os
import sys
import json
import struct
import threading
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

CONFIG_FILENAME = "config.json"
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
    """
    Extract all audio from one XWB file into out_dir.
    track_names_map: optional dict of {hex_name: friendly_name}
    stop_event: threading.Event to check for cancellation
    Returns list of output file paths.
    """
    with open(xwb_path, "rb") as f:
        sig = f.read(4)
        if sig == SIGN_LE:
            ru32 = ru32_le
        elif sig == SIGN_BE:
            ru32 = ru32_be
        else:
            f.seek(0)
            scan_size = min(1024 * 1024, os.path.getsize(xwb_path))
            raw = f.read(scan_size)
            found = -1
            for i in range(len(raw) - 8):
                s = raw[i:i+4]
                if (s == SIGN_LE and raw[i+7] == 0) or (s == SIGN_BE and raw[i+4] == 0):
                    found = i
                    break
            if found < 0:
                raise ValueError("Not a valid XWB file (no signature found)")
            f.seek(found + 4)
            sig  = raw[found:found+4]
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

            hex_name = f"{entry_idx:08x}"
            if track_names_map and hex_name in track_names_map:
                friendly = track_names_map[hex_name]
                base_name = f"{friendly}"
            else:
                base_name = hex_name

            ext      = ".wma" if codec == CODEC_WMA else ".wav"
            out_path = os.path.join(out_dir, base_name + ext)

            f.seek(play_off)
            with open(out_path, "wb") as fout:
                if codec != CODEC_WMA:
                    hdr = make_wav_header(codec, chans, rate, bits, align, play_len)
                    if hdr:
                        fout.write(hdr)
                copy_bytes(f, fout, play_len)

            output_files.append(out_path)

    return output_files


# ── GUI ──────────────────────────────────────────────────────────────────────

BG       = "#1a1a2e"
PANEL    = "#16213e"
ACCENT   = "#e94560"
TEXT     = "#eaeaea"
MUTED    = "#7a7a9a"
SUCCESS  = "#4ecca3"
WARNING  = "#f5a623"
FONT     = ("Consolas", 10)
FONT_BIG = ("Consolas", 13, "bold")
FONT_SM  = ("Consolas", 9)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("XWB → WAV Extractor")
        self.resizable(False, False)
        self.configure(bg=BG)

        self.input_var  = tk.StringVar()
        self.output_var = tk.StringVar()
        self.config_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready.")

        self._stop_event  = threading.Event()
        self._running     = False
        self._track_names = {}

        self._build_ui()
        self._center()
        self._try_load_config_from_cwd()

    # ── Layout ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        pad = dict(padx=16, pady=6)

        # ── Title bar ──
        title_frame = tk.Frame(self, bg=ACCENT, height=4)
        title_frame.pack(fill="x")

        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=20, pady=(16, 4))
        tk.Label(header, text="XWB → WAV Extractor", font=FONT_BIG,
                 bg=BG, fg=TEXT).pack(side="left")
        tk.Label(header, text="for RE4 & XACT games", font=FONT_SM,
                 bg=BG, fg=MUTED).pack(side="left", padx=(10, 0), pady=(4, 0))

        sep = tk.Frame(self, bg=PANEL, height=1)
        sep.pack(fill="x", padx=16, pady=(4, 12))

        # ── Folder pickers ──
        self._folder_row("XWB Folder (input):", self.input_var,
                         self._browse_input, hint="Folder containing your .xwb files")
        self._folder_row("Output Folder:", self.output_var,
                         self._browse_output, hint="Where extracted WAV files will be saved")
        self._folder_row("Config File (optional):", self.config_var,
                         self._browse_config, hint="JSON file for custom track names  —  leave blank to skip",
                         is_file=True)

        sep2 = tk.Frame(self, bg=PANEL, height=1)
        sep2.pack(fill="x", padx=16, pady=(8, 10))

        # ── Progress ──
        prog_frame = tk.Frame(self, bg=BG)
        prog_frame.pack(fill="x", padx=20, pady=(0, 4))

        self._prog_label = tk.Label(prog_frame, text="", font=FONT_SM, bg=BG, fg=MUTED)
        self._prog_label.pack(side="right")

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor=PANEL, background=ACCENT,
                        bordercolor=PANEL, lightcolor=ACCENT, darkcolor=ACCENT)

        self._progress = ttk.Progressbar(self, style="Custom.Horizontal.TProgressbar",
                                          length=560, mode="determinate")
        self._progress.pack(padx=20, pady=(0, 10))

        # ── Log box ──
        log_frame = tk.Frame(self, bg=PANEL, bd=0)
        log_frame.pack(fill="both", padx=20, pady=(0, 10))

        self._log = tk.Text(log_frame, width=72, height=14, font=FONT_SM,
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

        # ── Buttons ──
        btn_frame = tk.Frame(self, bg=BG)
        btn_frame.pack(pady=(0, 16))

        self._start_btn = tk.Button(btn_frame, text="▶  Start Extraction",
                                    font=FONT, bg=ACCENT, fg=TEXT,
                                    activebackground="#c73652", activeforeground=TEXT,
                                    relief="flat", bd=0, padx=20, pady=8,
                                    cursor="hand2", command=self._start)
        self._start_btn.pack(side="left", padx=(0, 8))

        self._stop_btn = tk.Button(btn_frame, text="■  Stop",
                                   font=FONT, bg=PANEL, fg=MUTED,
                                   activebackground="#2a2a4e", activeforeground=TEXT,
                                   relief="flat", bd=0, padx=20, pady=8,
                                   cursor="hand2", command=self._stop,
                                   state="disabled")
        self._stop_btn.pack(side="left", padx=(0, 8))

        self._config_btn = tk.Button(btn_frame, text="⚙  Create Config Template",
                                     font=FONT, bg=PANEL, fg=MUTED,
                                     activebackground="#2a2a4e", activeforeground=TEXT,
                                     relief="flat", bd=0, padx=20, pady=8,
                                     cursor="hand2", command=self._create_config)
        self._config_btn.pack(side="left")

        # ── Status bar ──
        status_bar = tk.Frame(self, bg=PANEL)
        status_bar.pack(fill="x", side="bottom")
        tk.Label(status_bar, textvariable=self.status_var, font=FONT_SM,
                 bg=PANEL, fg=MUTED, anchor="w").pack(side="left", padx=10, pady=4)

    def _folder_row(self, label, var, command, hint="", is_file=False):
        frame = tk.Frame(self, bg=BG)
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
            messagebox.showerror("Error", "Please select a valid XWB input folder.")
            return
        if not output_folder:
            messagebox.showerror("Error", "Please select an output folder.")
            return

        xwb_files = sorted(f for f in os.listdir(input_folder) if f.lower().endswith(".xwb"))
        if not xwb_files:
            messagebox.showerror("Error", f"No .xwb files found in:\n{input_folder}")
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

        self._log_line(f"Starting extraction of {total} XWB files...\n", "heading")

        ok_lines   = []
        fail_lines = []

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

        # Write log files
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
