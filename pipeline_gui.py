#!/usr/bin/env python3
"""
🎙️  Audio Analysis Pipeline — GUI Edition
==========================================
A standalone CustomTkinter app that wraps all pipeline tools
into clickable buttons with no terminal commands needed.
Visual design inspired by the noScribe GUI.
"""

import os
import sys
import re
import glob
import subprocess
import time
import threading
import platform
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import yaml
from PIL import Image

# ── Optional: drag-and-drop support ──
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

# ── Fixed Paths ────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ANALYSIS_DIR = os.path.join(BASE_DIR, "analysis")
CONFIG_FILE = os.path.join(BASE_DIR, "pipeline_config.yml")
MAX_RECENT = 5

# ── Default folder paths ─────────────────────────────
DEFAULT_PATHS = {
    "videos": os.path.join(BASE_DIR, "videos"),
    "transcripts": os.path.join(BASE_DIR, "transcripts"),
    "exports": os.path.join(BASE_DIR, "exports"),
    "graphs": os.path.join(ANALYSIS_DIR, "graphs_output"),
}

def load_config():
    """Load folder paths and recent files from YAML config."""
    paths = dict(DEFAULT_PATHS)
    recent = {"videos": [], "transcripts": [], "exports": []}
    history = []
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                saved = yaml.safe_load(f) or {}
            for key in paths:
                if key in saved and os.path.isdir(saved[key]):
                    paths[key] = saved[key]
            if "recent" in saved and isinstance(saved["recent"], dict):
                for key in recent:
                    if key in saved["recent"] and isinstance(saved["recent"][key], list):
                        recent[key] = [f for f in saved["recent"][key] if os.path.exists(f)][:MAX_RECENT]
            if "history" in saved and isinstance(saved["history"], list):
                history = saved["history"][-50:]  # keep last 50
        except Exception:
            pass
    for d in paths.values():
        os.makedirs(d, exist_ok=True)
    return paths, recent, history

def save_config(paths, recent=None, history=None):
    """Persist folder paths, recent files, and history to YAML config."""
    data = dict(paths)
    if recent:
        data["recent"] = recent
    if history is not None:
        data["history"] = history[-50:]
    with open(CONFIG_FILE, "w") as f:
        yaml.safe_dump(data, f, default_flow_style=False)

# ── Theme Data ─────────────────────────────────────────
THEMES = {
    "Professional Blue": "blue",
    "Forest Green": "green",
    "Warm Sunset": "sunset",
    "Dark Mode": "dark",
    "Soft Purple": "purple",
}

# ── Languages (from noScribe) ─────────────────────────
LANGUAGES = {
    "Auto": "auto", "English": "en", "Finnish": "fi", "German": "de",
    "French": "fr", "Spanish": "es", "Swedish": "sv", "Norwegian": "no",
    "Danish": "da", "Dutch": "nl", "Italian": "it", "Portuguese": "pt",
    "Russian": "ru", "Chinese": "zh", "Japanese": "ja", "Korean": "ko",
    "Arabic": "ar", "Hindi": "hi", "Turkish": "tr", "Polish": "pl",
    "Czech": "cs", "Romanian": "ro", "Hungarian": "hu", "Greek": "el",
    "Bulgarian": "bg", "Croatian": "hr", "Slovak": "sk", "Slovenian": "sl",
    "Estonian": "et", "Latvian": "lv", "Lithuanian": "lt", "Ukrainian": "uk",
    "Serbian": "sr", "Bosnian": "bs", "Macedonian": "mk", "Icelandic": "is",
    "Catalan": "ca", "Galician": "gl", "Welsh": "cy", "Afrikaans": "af",
    "Swahili": "sw", "Malay": "ms", "Indonesian": "id", "Vietnamese": "vi",
    "Thai": "th", "Tagalog": "tl", "Tamil": "ta", "Kannada": "kn",
    "Marathi": "mr", "Nepali": "ne", "Urdu": "ur", "Persian": "fa",
    "Hebrew": "he", "Armenian": "hy", "Azerbaijani": "az", "Belarusian": "be",
    "Kazakh": "kk", "Maori": "mi", "Multilingual": "multilingual",
}

# ── App Settings ───────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class PipelineApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Audio Analysis Pipeline")
        self.geometry("1100x765")
        self.minsize(900, 650)

        # Try to set icon
        logo_ico = os.path.join(BASE_DIR, "noScribeLogo.ico")
        if os.path.exists(logo_ico):
            try:
                self.iconbitmap(logo_ico)
            except Exception:
                pass

        # ── Load configurable paths + recent files ──
        self.paths, self.recent, self.history = load_config()

        # ── Auto-populate state ──
        self.last_transcript_path = None
        self.last_excel_path = None

        # ═══════════════════════════════════════════
        #  HEADER BANNER
        # ═══════════════════════════════════════════
        self.frame_header = ctk.CTkFrame(self, height=110, corner_radius=0)
        self.frame_header.pack(padx=0, pady=0, anchor="nw", fill="x")
        self.frame_header.pack_propagate(False)

        header_text_frame = ctk.CTkFrame(self.frame_header, fg_color="transparent")
        header_text_frame.pack(anchor="w", side="left")

        ctk.CTkLabel(
            header_text_frame, text="Audio Analysis Pipeline",
            font=ctk.CTkFont(size=36, weight="bold")
        ).pack(padx=20, pady=(30, 0), anchor="w")

        ctk.CTkLabel(
            header_text_frame, text="Transcribe · Analyse · Visualise",
            font=ctk.CTkFont(size=15, weight="bold"), text_color="#aaaaaa"
        ).pack(padx=20, pady=(0, 20), anchor="w")

        graphic_path = os.path.join(BASE_DIR, "graphic_sw.png")
        if os.path.exists(graphic_path):
            try:
                self.header_graphic = ctk.CTkImage(dark_image=Image.open(graphic_path), size=(500, 65))
                ctk.CTkLabel(self.frame_header, image=self.header_graphic, text="").pack(
                    anchor="ne", side="right", padx=(30, 30)
                )
            except Exception:
                pass

        # ═══════════════════════════════════════════
        #  MAIN LAYOUT: sidebar + content
        # ═══════════════════════════════════════════
        self.frame_main = ctk.CTkFrame(self, corner_radius=0)
        self.frame_main.pack(padx=0, pady=0, anchor="nw", expand=True, fill="both")

        # ── Sidebar ──
        self.sidebar = ctk.CTkFrame(self.frame_main, width=240, corner_radius=0, fg_color="transparent")
        self.sidebar.pack(padx=0, pady=0, fill="y", expand=False, side="left")
        self.sidebar.pack_propagate(False)

        self.nav_frame = ctk.CTkScrollableFrame(self.sidebar, fg_color="transparent", width=220)
        self.nav_frame.pack(padx=10, pady=(15, 5), fill="both", expand=True)

        ctk.CTkLabel(self.nav_frame, text="Pipeline Steps",
            font=ctk.CTkFont(size=13, weight="bold"), text_color="#888888"
        ).pack(anchor="w", padx=10, pady=(5, 10))

        for text, command in [
            ("📥  Download Audio", self.show_download),
            ("🔧  Prepare Audio", self.show_prepare_audio),
            ("🎙️  Transcribe", self.show_transcribe),
            ("🏷️  Rename Speakers", self.show_rename),
            ("📊  Excel + Theme", self.show_excel),
            ("📈  Graphs", self.show_graphs),
        ]:
            ctk.CTkButton(
                self.nav_frame, text=text, command=command,
                height=38, corner_radius=8,
                fg_color="transparent", hover_color=("gray75", "gray30"),
                anchor="w", font=ctk.CTkFont(size=14)
            ).pack(padx=5, pady=2, fill="x")

        ctk.CTkFrame(self.nav_frame, height=2, fg_color="gray30").pack(padx=15, pady=12, fill="x")

        ctk.CTkButton(
            self.nav_frame, text="🚀  Full Pipeline",
            command=self.show_full_pipeline,
            height=42, corner_radius=8,
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(padx=5, pady=2, fill="x")

        ctk.CTkFrame(self.nav_frame, height=2, fg_color="gray30").pack(padx=15, pady=12, fill="x")

        ctk.CTkButton(
            self.nav_frame, text="⚙️  Settings", command=self.show_settings,
            height=34, corner_radius=8, fg_color="transparent",
            hover_color=("gray75", "gray30"), anchor="w",
            font=ctk.CTkFont(size=13), text_color="#999999"
        ).pack(padx=5, pady=2, fill="x")

        ctk.CTkButton(
            self.nav_frame, text="📋  History", command=self.show_history,
            height=34, corner_radius=8, fg_color="transparent",
            hover_color=("gray75", "gray30"), anchor="w",
            font=ctk.CTkFont(size=13), text_color="#999999"
        ).pack(padx=5, pady=2, fill="x")

        # ── Right content area ──
        self.content_frame = ctk.CTkFrame(self.frame_main, corner_radius=0, fg_color="transparent")
        self.content_frame.pack(padx=0, pady=0, fill="both", expand=True, side="top")

        self.main_panel = ctk.CTkFrame(self.content_frame, fg_color="transparent", border_width=1, corner_radius=0)
        self.main_panel.pack(padx=(10, 30), pady=(0, 30), fill="both", expand=True)
        self.main_panel.grid_columnconfigure(0, weight=1)
        self.main_panel.grid_rowconfigure(1, weight=1)

        # ── Status Bar ──
        self.status_bar = ctk.CTkLabel(
            self, text="  Ready", anchor="w",
            font=ctk.CTkFont(size=12), text_color="lightgray",
            fg_color="gray17", height=26
        )
        self.status_bar.pack(padx=0, pady=0, fill="x", side="bottom")

        # ── Drag & Drop setup ──
        if HAS_DND:
            try:
                self.drop_target_register(DND_FILES)
                self.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

        self.show_welcome()

    # ═══════════════════════════════════════════════
    #  HELPERS
    # ═══════════════════════════════════════════════
    def clear_main(self):
        for widget in self.main_panel.winfo_children():
            widget.destroy()

    def set_status(self, text):
        self.status_bar.configure(text=f"  {text}")

    def make_title(self, text):
        ctk.CTkLabel(
            self.main_panel, text=text,
            font=ctk.CTkFont(size=22, weight="bold")
        ).grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

    def make_log(self, parent):
        return ctk.CTkTextbox(parent, font=("", 14), text_color="lightgray",
            bg_color="transparent", fg_color="transparent", wrap="word")

    def run_in_thread(self, func, *args):
        threading.Thread(target=func, args=args, daemon=True).start()

    def open_in_finder(self, path):
        """Open a file or folder in the native file manager."""
        target = path if os.path.isdir(path) else os.path.dirname(path)
        try:
            if platform.system() == "Darwin":
                subprocess.run(["open", target])
            elif platform.system() == "Windows":
                os.startfile(target)
            else:
                subprocess.run(["xdg-open", target])
        except Exception:
            pass

    def add_to_recent(self, category, filepath):
        """Add a file to the recent files list for a category."""
        if category not in self.recent:
            self.recent[category] = []
        # Remove if already present, then prepend
        self.recent[category] = [f for f in self.recent[category] if f != filepath]
        self.recent[category].insert(0, filepath)
        self.recent[category] = self.recent[category][:MAX_RECENT]
        save_config(self.paths, self.recent, self.history)

    def get_recent(self, category):
        """Get list of recent files for a category."""
        return [f for f in self.recent.get(category, []) if os.path.exists(f)]

    def _append_log(self, log_widget, text):
        """Thread-safe helper to append a line to a log textbox."""
        log_widget.insert("end", text + "\n")
        log_widget.see("end")

    def _start_elapsed_timer(self, label="Transcribing"):
        """Start a recurring 1-second timer that updates the status bar with elapsed time."""
        self._elapsed_start = time.time()
        self._elapsed_label = label
        self._elapsed_running = True
        self._tick_elapsed()

    def _tick_elapsed(self):
        """Update the status bar with the current elapsed time."""
        if not self._elapsed_running:
            return
        secs = int(time.time() - self._elapsed_start)
        mins, secs = divmod(secs, 60)
        if mins > 0:
            elapsed = f"{mins}m {secs:02d}s"
        else:
            elapsed = f"{secs}s"
        self.set_status(f"{self._elapsed_label}... ({elapsed})")
        self.after(1000, self._tick_elapsed)

    def _stop_elapsed_timer(self):
        """Stop the elapsed timer and return the formatted duration."""
        self._elapsed_running = False
        secs = int(time.time() - self._elapsed_start)
        mins, secs = divmod(secs, 60)
        return f"{mins}m {secs:02d}s" if mins > 0 else f"{secs}s"

    def _add_open_folder_btn(self, parent, path, row=None):
        """Add an 'Open in Finder' button after a task completes."""
        btn = ctk.CTkButton(
            parent, text="📂  Open in Finder", height=35, width=180,
            font=ctk.CTkFont(size=13),
            fg_color="gray40", hover_color="gray50",
            command=lambda: self.open_in_finder(path)
        )
        if row is not None:
            btn.grid(row=row, column=0, sticky="w", pady=(5, 0), padx=20)
        else:
            btn.pack(anchor="w", pady=(5, 0))
        return btn

    def _make_file_picker(self, parent, var, label, browse_cmd, category=None, row=0):
        """Create a noScribe-style file picker row with optional recent-files dropdown."""
        ctk.CTkLabel(parent, text=label, font=ctk.CTkFont(size=14)).grid(
            row=row, column=0, sticky="w", pady=5, padx=(0, 10)
        )

        picker_frame = ctk.CTkFrame(parent, fg_color="transparent")
        picker_frame.grid(row=row, column=1, sticky="ew", pady=5)
        picker_frame.grid_columnconfigure(0, weight=1)

        recent = self.get_recent(category) if category else []

        if recent:
            # Show combobox with recent files
            combo = ctk.CTkComboBox(
                picker_frame, variable=var,
                values=[os.path.basename(f) for f in recent],
                height=33, corner_radius=8, border_width=2,
                font=ctk.CTkFont(size=12),
                command=lambda choice: self._select_recent(var, category, choice)
            )
            combo.set(var.get() if var.get() else "<select file or pick recent>")
            combo.pack(side="left", fill="x", expand=True, padx=(0, 5))

            # Store full paths for lookup
            self._recent_lookup = {os.path.basename(f): f for f in recent}
        else:
            entry = ctk.CTkEntry(picker_frame, textvariable=var, height=33,
                corner_radius=8, border_width=2, font=ctk.CTkFont(size=12))
            entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        ctk.CTkButton(picker_frame, width=45, height=33, text="📂",
            command=browse_cmd
        ).pack(side="right")

    def _select_recent(self, var, category, choice):
        """Handle selection from recent files dropdown."""
        full_path = None
        if hasattr(self, '_recent_lookup') and choice in self._recent_lookup:
            full_path = self._recent_lookup[choice]
        else:
            for f in self.get_recent(category):
                if os.path.basename(f) == choice:
                    full_path = f
                    break
        if full_path:
            var.set(full_path)
            # Auto-fill output name on transcribe screen
            if category == "videos" and hasattr(self, 'transcribe_output_var'):
                self.transcribe_output_var.set(os.path.splitext(os.path.basename(full_path))[0])

    def _on_drop(self, event):
        """Handle drag-and-drop of files onto the app."""
        files = self.tk.splitlist(event.data)
        if not files:
            return

        filepath = files[0]
        ext = os.path.splitext(filepath)[1].lower()

        if ext == ".mp3":
            # Auto-navigate to transcribe screen and fill it in
            self.show_transcribe()
            self.transcribe_file_var.set(filepath)
            self.add_to_recent("videos", filepath)
            self.set_status(f"Dropped: {os.path.basename(filepath)} → Transcribe")
        elif ext == ".txt":
            self.show_rename()
            self.rename_file_var.set(filepath)
            self.add_to_recent("transcripts", filepath)
            self.set_status(f"Dropped: {os.path.basename(filepath)} → Rename Speakers")
        elif ext == ".xlsx":
            self.show_graphs()
            self.graphs_file_var.set(filepath)
            self.add_to_recent("exports", filepath)
            self.set_status(f"Dropped: {os.path.basename(filepath)} → Graphs")
        else:
            self.set_status(f"Unsupported file type: {ext}")

    # ═══════════════════════════════════════════════
    #  SETTINGS SCREEN
    # ═══════════════════════════════════════════════
    def show_settings(self):
        self.clear_main()
        self.make_title("⚙️  Settings — Folder Paths")

        content = ctk.CTkFrame(self.main_panel, fg_color="transparent")
        content.grid(row=1, column=0, padx=20, pady=10, sticky="nswe")

        ctk.CTkLabel(
            content, text="Change where files are saved and loaded from.\n"
            "All file browsers will automatically use these folders.",
            font=ctk.CTkFont(size=14), text_color="#aaaaaa", justify="left"
        ).pack(anchor="w", pady=(0, 20))

        self.settings_vars = {}
        for key, (label, hint) in {
            "videos": ("🎬  Videos / Audio", "Where downloaded .mp3 files are saved"),
            "transcripts": ("📝  Transcripts", "Where .txt transcript files are saved"),
            "exports": ("📊  Excel Exports", "Where .xlsx analysis files are saved"),
            "graphs": ("📈  Graphs", "Where visualization images are saved"),
        }.items():
            group = ctk.CTkFrame(content, fg_color="transparent")
            group.pack(fill="x", pady=(0, 12))
            ctk.CTkLabel(group, text=label, font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w")
            ctk.CTkLabel(group, text=hint, font=ctk.CTkFont(size=11), text_color="#777777").pack(anchor="w", pady=(0, 4))
            row = ctk.CTkFrame(group, fg_color="transparent")
            row.pack(fill="x")
            var = ctk.StringVar(value=self.paths[key])
            self.settings_vars[key] = var
            ctk.CTkEntry(row, textvariable=var, height=33, corner_radius=8, border_width=2, font=ctk.CTkFont(size=12)).pack(side="left", fill="x", expand=True, padx=(0, 8))
            def make_browse(v=var):
                def browse():
                    folder = filedialog.askdirectory(initialdir=v.get())
                    if folder: v.set(folder)
                return browse
            ctk.CTkButton(row, text="📂", width=45, height=33, command=make_browse()).pack(side="right")

        # Drag & drop status
        dnd_status = "✅ Drag & drop enabled (tkinterdnd2 installed)" if HAS_DND else "ℹ️  Drag & drop unavailable — install tkinterdnd2 to enable"
        ctk.CTkLabel(content, text=dnd_status, font=ctk.CTkFont(size=11), text_color="#666666").pack(anchor="w", pady=(10, 0))

        btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(15, 0))
        ctk.CTkButton(btn_frame, text="Save Settings", height=42,
            font=ctk.CTkFont(size=14, weight="bold"), command=self.do_save_settings
        ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_frame, text="Reset to Defaults", height=42,
            font=ctk.CTkFont(size=13), fg_color="gray40", hover_color="gray50",
            command=self.do_reset_settings
        ).pack(side="left")

    def do_save_settings(self):
        for key, var in self.settings_vars.items():
            folder = var.get().strip()
            if folder:
                os.makedirs(folder, exist_ok=True)
                self.paths[key] = folder
        save_config(self.paths, self.recent, self.history)
        self.set_status("✅ Settings saved!")
        messagebox.showinfo("Saved", "✅ Folder paths saved!")

    def do_reset_settings(self):
        for key, default in DEFAULT_PATHS.items():
            self.settings_vars[key].set(default)
            self.paths[key] = default
        save_config(self.paths, self.recent, self.history)
        self.set_status("Settings reset to defaults")

    # ═══════════════════════════════════════════════
    #  WELCOME SCREEN
    # ═══════════════════════════════════════════════
    def show_welcome(self):
        self.clear_main()
        self.make_title("Welcome")

        log = self.make_log(self.main_panel)
        log.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nswe")

        log.insert("end", "Hi there, the Audio Analysis Pipeline is ready!\n\n", "highlight")
        log.tag_config("highlight", foreground="darkorange")
        log.insert("end", "Powered by noScribe (Whisper + pyannote) for transcription,\n")
        log.insert("end", "openpyxl for Excel generation, and matplotlib for visualization.\n\n")
        log.insert("end", "Use the buttons on the left to select a pipeline step:\n\n")
        log.insert("end", "  📥  Download Audio    — grab audio from YouTube\n")
        log.insert("end", "  🎙️  Transcribe        — convert audio to text (AI)\n")
        log.insert("end", "  🏷️  Rename Speakers   — replace S00/S01 with real names\n")
        log.insert("end", "  📊  Excel + Theme     — create styled spreadsheet\n")
        log.insert("end", "  📈  Graphs            — generate timeline & heatmap\n\n")
        log.insert("end", "Or click  🚀 Full Pipeline  to run everything at once.\n\n")
        if HAS_DND:
            log.insert("end", "💡 Tip: Drag & drop .mp3, .txt, or .xlsx files onto this window!\n")
        log.insert("end", "💡 Tip: Set Start/Stop time to transcribe just a section.\n")
        log.configure(state="disabled")

    # ═══════════════════════════════════════════════
    #  DOWNLOAD SCREEN
    # ═══════════════════════════════════════════════
    def show_download(self):
        self.clear_main()
        self.make_title("📥  Download Audio from YouTube")

        content = ctk.CTkFrame(self.main_panel, fg_color="transparent")
        content.grid(row=1, column=0, padx=20, pady=10, sticky="nswe")
        content.grid_rowconfigure(2, weight=1)
        content.grid_columnconfigure(0, weight=1)

        opts = ctk.CTkFrame(content, fg_color="transparent")
        opts.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        opts.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(opts, text="YouTube URL:", font=ctk.CTkFont(size=14)).grid(row=0, column=0, sticky="w", pady=5, padx=(0, 10))
        self.url_entry = ctk.CTkEntry(opts, placeholder_text="https://www.youtube.com/watch?v=...", height=33, corner_radius=8, border_width=2)
        self.url_entry.grid(row=0, column=1, sticky="ew", pady=5)

        self.download_btn = ctk.CTkButton(content, text="Start Download", height=42,
            font=ctk.CTkFont(size=14, weight="bold"), command=self.do_download)
        self.download_btn.grid(row=1, column=0, sticky="w", pady=(0, 10))

        self.download_log = self.make_log(content)
        self.download_log.grid(row=2, column=0, sticky="nswe")

    def do_download(self):
        url = self.url_entry.get().strip()
        if len(url) < 10:
            messagebox.showwarning("Invalid URL", "Please enter a valid YouTube URL.")
            return

        self.download_btn.configure(state="disabled", text="Downloading...")
        self.set_status("Downloading audio...")
        self.download_log.configure(state="normal")
        self.download_log.delete("0.0", "end")
        self.download_log.insert("end", f"Downloading: {url}\n\n")

        def worker():
            result = subprocess.run(
                [os.path.join(BASE_DIR, "download_audio.sh"), url, self.paths["videos"]],
                capture_output=True, text=True, cwd=BASE_DIR
            )
            self.after(0, lambda: self._download_done(result))

        self.run_in_thread(worker)

    def _download_done(self, result):
        self.download_btn.configure(state="normal", text="Start Download")
        if result.returncode == 0:
            self.download_log.insert("end", result.stdout + "\n\n✅ Download complete!")
            self.set_status("✅ Audio downloaded")
            # Find the newest mp3 and add to recent
            mp3s = sorted(glob.glob(os.path.join(self.paths["videos"], "*.mp3")), key=os.path.getmtime, reverse=True)
            if mp3s:
                self.add_to_recent("videos", mp3s[0])
            self._add_open_folder_btn(self.download_log.master, self.paths["videos"], row=3)
        else:
            self.download_log.insert("end", f"❌ Error:\n{result.stderr}")
            self.set_status("❌ Download failed")

    # ═══════════════════════════════════════════════
    #  PREPARE AUDIO SCREEN
    # ═══════════════════════════════════════════════
    def show_prepare_audio(self):
        self.clear_main()
        self.make_title("🔧  Prepare Audio")

        content = ctk.CTkFrame(self.main_panel, fg_color="transparent")
        content.grid(row=1, column=0, padx=20, pady=10, sticky="nswe")
        content.grid_rowconfigure(3, weight=1)
        content.grid_columnconfigure(0, weight=1)

        opts = ctk.CTkFrame(content, fg_color="transparent")
        opts.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        opts.grid_columnconfigure(1, weight=1)

        # File picker
        self.prep_file_var = ctk.StringVar()
        def browse_audio():
            path = filedialog.askopenfilename(initialdir=self.paths["videos"],
                filetypes=[("Audio/Video Files", "*.mp3 *.m4a *.wav *.flac *.mp4 *.mkv"), ("All Files", "*.*")])
            if path:
                self.prep_file_var.set(path)
                self.add_to_recent("videos", path)
        self._make_file_picker(opts, self.prep_file_var, "Input File:", browse_audio, "videos", row=0)

        # Options panel
        prep_frame = ctk.CTkFrame(content, fg_color="gray20", corner_radius=8)
        prep_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        
        # Row 1: Format & Bitrate
        ctk.CTkLabel(prep_frame, text="Format:", font=ctk.CTkFont(size=13)).grid(row=0, column=0, sticky="w", padx=12, pady=(12,5))
        self.prep_fmt_var = ctk.StringVar(value="mp3")
        ctk.CTkOptionMenu(prep_frame, variable=self.prep_fmt_var, values=["mp3", "wav", "flac", "ogg"], 
                          width=90, dynamic_resizing=False).grid(row=0, column=1, sticky="w", padx=5, pady=(12,5))
                          
        ctk.CTkLabel(prep_frame, text="Bitrate:", font=ctk.CTkFont(size=13)).grid(row=0, column=2, sticky="w", padx=(20,5), pady=(12,5))
        self.prep_bitrate_var = ctk.StringVar(value="192k")
        ctk.CTkOptionMenu(prep_frame, variable=self.prep_bitrate_var, values=["128k", "192k", "256k", "320k"], 
                          width=90, dynamic_resizing=False).grid(row=0, column=3, sticky="w", padx=5, pady=(12,5))

        # Row 2: Trimming
        ctk.CTkLabel(prep_frame, text="Trim Start:", font=ctk.CTkFont(size=13)).grid(row=1, column=0, sticky="w", padx=12, pady=5)
        self.prep_start_var = ctk.StringVar()
        ctk.CTkEntry(prep_frame, textvariable=self.prep_start_var, placeholder_text="e.g. 01:30", width=90).grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        ctk.CTkLabel(prep_frame, text="Trim Stop:", font=ctk.CTkFont(size=13)).grid(row=1, column=2, sticky="w", padx=(20,5), pady=5)
        self.prep_stop_var = ctk.StringVar()
        ctk.CTkEntry(prep_frame, textvariable=self.prep_stop_var, placeholder_text="e.g. 15:45", width=90).grid(row=1, column=3, sticky="w", padx=5, pady=5)

        # Row 3: Enhancements
        self.prep_norm_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(prep_frame, text="Normalize volume (make consistent loudness)", 
                        variable=self.prep_norm_var, font=ctk.CTkFont(size=13)).grid(row=2, column=0, columnspan=4, sticky="w", padx=12, pady=(5,5))
                        
        self.prep_split_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(prep_frame, text="Split into multiple files at long silence gaps", 
                        variable=self.prep_split_var, font=ctk.CTkFont(size=13)).grid(row=3, column=0, columnspan=4, sticky="w", padx=12, pady=(5,12))

        # Action Area
        action_row = ctk.CTkFrame(content, fg_color="transparent")
        action_row.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        
        self.prepare_btn = ctk.CTkButton(action_row, text="Prepare Audio", height=42, width=200,
            font=ctk.CTkFont(size=14, weight="bold"), command=self.do_prepare_audio)
        self.prepare_btn.grid(row=0, column=0, sticky="w")
        
        self.prepare_progress = ctk.CTkProgressBar(action_row, width=400, height=10, mode="indeterminate")
        self.prepare_progress.set(0)

        self.prepare_log = self.make_log(content)
        self.prepare_log.grid(row=3, column=0, sticky="nswe")

    def do_prepare_audio(self):
        url = self.prep_file_var.get().strip()
        if not url or not os.path.exists(url):
            messagebox.showwarning("Input Error", "Please select a valid input file.")
            return

        self.add_to_recent("videos", url)
        basename = os.path.splitext(os.path.basename(url))[0]
        fmt = self.prep_fmt_var.get()
        out_name = f"{basename}_prep.{fmt}"
        output_path = os.path.join(self.paths["videos"], out_name)

        cmd = [sys.executable, os.path.join(ANALYSIS_DIR, "prepare_audio.py"), url, output_path]
        
        cmd.extend(["--format", fmt])
        cmd.extend(["--bitrate", self.prep_bitrate_var.get()])
        
        if self.prep_start_var.get().strip():
            cmd.extend(["--start", self.prep_start_var.get().strip()])
        if self.prep_stop_var.get().strip():
            cmd.extend(["--stop", self.prep_stop_var.get().strip()])
            
        if self.prep_norm_var.get():
            cmd.append("--normalize")
            
        if self.prep_split_var.get():
            cmd.append("--split")

        self.prepare_btn.configure(state="disabled", text="Processing...")
        self.set_status("Processing audio...")
        self.prepare_progress.grid(row=0, column=1, padx=20, sticky="w")
        self.prepare_progress.start()
        
        self.prepare_log.configure(state="normal")
        self.prepare_log.delete("0.0", "end")
        self.prepare_log.insert("end", f"Starting audio preparation...\nCommand: {' '.join(cmd)}\n\n")

        def worker():
            result = subprocess.run(cmd, capture_output=True, text=True, cwd=BASE_DIR)
            out_target = output_path
            if self.prep_split_var.get():
                out_target = os.path.splitext(output_path)[0] + "_parts"
            self.after(0, lambda: self._prepare_done(result, out_target))

        self.run_in_thread(worker)

    def _prepare_done(self, result, output_path):
        self.prepare_btn.configure(state="normal", text="Prepare Audio")
        self.prepare_progress.stop()
        self.prepare_progress.set(1 if result.returncode == 0 else 0)

        if result.returncode == 0:
            self.prepare_log.insert("end", f"\n{result.stdout}\n\n✅ Audio prepared successfully:\n{output_path}")
            self.set_status("✅ Audio preparation complete!")
            self.add_to_recent("videos", output_path)
            self._add_open_folder_btn(self.prepare_log.master, output_path, row=2)
        else:
            err = result.stderr or result.stdout or "Unknown error"
            self.prepare_log.insert("end", f"\n\n❌ Preparation failed:\n{err}")
            self.set_status("❌ Audio preparation failed")

    # ═══════════════════════════════════════════════
    #  TRANSCRIBE SCREEN
    # ═══════════════════════════════════════════════
    def show_transcribe(self):
        self.clear_main()
        self.make_title("🎙️  Transcribe Audio")

        content = ctk.CTkFrame(self.main_panel, fg_color="transparent")
        content.grid(row=1, column=0, padx=20, pady=10, sticky="nswe")
        content.grid_rowconfigure(3, weight=1)
        content.grid_columnconfigure(0, weight=1)

        opts = ctk.CTkFrame(content, fg_color="transparent")
        opts.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        opts.grid_columnconfigure(1, weight=1)

        # Audio file picker with recent files
        self.transcribe_file_var = ctk.StringVar()
        def browse_audio():
            path = filedialog.askopenfilename(
                initialdir=self.paths["videos"],
                filetypes=[("MP3 Files", "*.mp3"), ("All Files", "*.*")]
            )
            if path:
                self.transcribe_file_var.set(path)
                self.add_to_recent("videos", path)
                # Auto-fill output name from audio filename
                self.transcribe_output_var.set(os.path.splitext(os.path.basename(path))[0])

        self._make_file_picker(opts, self.transcribe_file_var, "Audio file:", browse_audio, "videos", row=0)

        # Batch button
        ctk.CTkButton(opts, text="📁 Batch (multiple)", width=160, height=33,
            font=ctk.CTkFont(size=12), fg_color="gray40", hover_color="gray50",
            command=self._browse_batch_audio
        ).grid(row=0, column=2, padx=(8, 0), pady=5)

        # Output name
        ctk.CTkLabel(opts, text="Output name:", font=ctk.CTkFont(size=14)).grid(row=1, column=0, sticky="w", pady=5, padx=(0, 10))
        self.transcribe_output_var = ctk.StringVar(value="my_transcript")
        ctk.CTkEntry(opts, textvariable=self.transcribe_output_var, height=33, width=250, corner_radius=8, border_width=2).grid(row=1, column=1, sticky="w", pady=5)

        # Language selector
        ctk.CTkLabel(opts, text="Language:", font=ctk.CTkFont(size=14)).grid(row=2, column=0, sticky="w", pady=5, padx=(0, 10))
        self.transcribe_lang_var = ctk.StringVar(value="Auto")
        ctk.CTkOptionMenu(opts, variable=self.transcribe_lang_var,
            values=list(LANGUAGES.keys()), width=180, dynamic_resizing=False
        ).grid(row=2, column=1, sticky="w", pady=5)

        # Start/Stop time
        ctk.CTkLabel(opts, text="Start (hh:mm:ss):", font=ctk.CTkFont(size=14)).grid(row=3, column=0, sticky="w", pady=5, padx=(0, 10))
        self.transcribe_start_var = ctk.StringVar()
        ctk.CTkEntry(opts, textvariable=self.transcribe_start_var, height=33, width=120, corner_radius=8, border_width=2, placeholder_text="00:00:00").grid(row=3, column=1, sticky="w", pady=5)

        ctk.CTkLabel(opts, text="Stop (hh:mm:ss):", font=ctk.CTkFont(size=14)).grid(row=4, column=0, sticky="w", pady=(5, 10), padx=(0, 10))
        self.transcribe_stop_var = ctk.StringVar()
        ctk.CTkEntry(opts, textvariable=self.transcribe_stop_var, height=33, width=120, corner_radius=8, border_width=2, placeholder_text="").grid(row=4, column=1, sticky="w", pady=(5, 10))

        # Start button + progress
        btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        btn_frame.grid(row=1, column=0, sticky="ew", pady=(0, 5))

        self.transcribe_btn = ctk.CTkButton(btn_frame, text="Start", height=42, width=200,
            font=ctk.CTkFont(size=14, weight="bold"), command=self.do_transcribe)
        self.transcribe_btn.pack(side="left")

        self.transcribe_progress = ctk.CTkProgressBar(btn_frame, mode="indeterminate", fg_color="gray17")
        self.transcribe_progress.pack(side="left", fill="x", expand=True, padx=(15, 0))
        self.transcribe_progress.set(0)

        self.transcribe_log = self.make_log(content)
        self.transcribe_log.grid(row=3, column=0, sticky="nswe", pady=(5, 0))

    def do_transcribe(self):
        audio_file = self.transcribe_file_var.get().strip()
        output_name = self.transcribe_output_var.get().strip()

        if not audio_file or not os.path.exists(audio_file):
            messagebox.showwarning("No file", "Please select an audio file.")
            return
        if not output_name:
            messagebox.showwarning("No name", "Please enter an output name.")
            return

        output_path = os.path.join(self.paths["transcripts"], f"{output_name}.txt")
        start_time = self.transcribe_start_var.get().strip()
        stop_time = self.transcribe_stop_var.get().strip()
        time_info = ""
        if start_time or stop_time:
            time_info = f"\nRange:  {start_time or '00:00:00'} → {stop_time or 'end'}"

        self.add_to_recent("videos", audio_file)
        self.transcribe_btn.configure(state="disabled", text="Transcribing...")
        self.transcribe_progress.start()
        self._start_elapsed_timer("Transcribing")
        self.transcribe_log.configure(state="normal")
        self.transcribe_log.delete("0.0", "end")
        self.transcribe_log.insert("end", f"Input:  {audio_file}\nOutput: {output_path}{time_info}\n\nTranscribing...\n")

        def worker():
            cmd = [sys.executable, "-u", os.path.join(BASE_DIR, "noScribe.py"),
                   audio_file, output_path, "--no-gui"]
            if start_time: cmd.extend(["--start", start_time])
            if stop_time: cmd.extend(["--stop", stop_time])
            lang_code = LANGUAGES.get(self.transcribe_lang_var.get(), "auto")
            if lang_code != "auto":
                cmd.extend(["--language", lang_code])

            proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, bufsize=1, cwd=BASE_DIR)
            for line in proc.stdout:
                stripped = line.rstrip()
                if stripped:
                    self.after(0, lambda l=stripped: self._append_log(self.transcribe_log, l))
            returncode = proc.wait()
            self.after(0, lambda: self._transcribe_done(returncode, output_path))

        self.run_in_thread(worker)

    def _transcribe_done(self, returncode, output_path):
        elapsed = self._stop_elapsed_timer()
        self.transcribe_btn.configure(state="normal", text="Start")
        self.transcribe_progress.stop()
        self.transcribe_progress.set(1 if returncode == 0 else 0)

        if returncode == 0:
            self.transcribe_log.insert("end", f"\n✅ Transcript saved to:\n{output_path}\n⏱️  Completed in {elapsed}")
            self.set_status(f"✅ Transcription complete! ({elapsed})")
            # Auto-populate for next steps
            self.last_transcript_path = output_path
            self.add_to_recent("transcripts", output_path)
            self._add_history_entry(self.transcribe_file_var.get(), output_path)
            # Open folder button
            self._add_open_folder_btn(self.transcribe_log.master, output_path, row=4)
        else:
            self.transcribe_log.insert("end", "\n❌ Transcription failed — see log above for details")
            self.set_status("❌ Transcription failed")

    def _browse_batch_audio(self):
        """Select multiple audio files for batch transcription."""
        files = filedialog.askopenfilenames(
            initialdir=self.paths["videos"],
            filetypes=[("MP3 Files", "*.mp3"), ("All Files", "*.*")]
        )
        if not files:
            return
        self._batch_files = list(files)
        self.transcribe_file_var.set(f"{len(files)} files selected")
        self.transcribe_output_var.set("(auto from filename)")
        self.set_status(f"Batch: {len(files)} files queued")
        # Override the Start button to run batch
        self.transcribe_btn.configure(command=self._do_batch_transcribe)

    def _do_batch_transcribe(self):
        """Transcribe multiple files sequentially."""
        if not hasattr(self, '_batch_files') or not self._batch_files:
            messagebox.showwarning("No files", "No batch files selected.")
            return
        files = self._batch_files
        lang_code = LANGUAGES.get(self.transcribe_lang_var.get(), "auto")
        start_time = self.transcribe_start_var.get().strip()
        stop_time = self.transcribe_stop_var.get().strip()

        self.transcribe_btn.configure(state="disabled", text=f"Batch (0/{len(files)})...")
        self.transcribe_progress.start()
        self._start_elapsed_timer(f"Batch transcribing {len(files)} files")
        self.transcribe_log.configure(state="normal")
        self.transcribe_log.delete("0.0", "end")
        self.transcribe_log.insert("end", f"Batch transcription: {len(files)} files\n{'='*50}\n\n")

        def worker():
            completed = 0
            for i, audio_file in enumerate(files):
                basename = os.path.splitext(os.path.basename(audio_file))[0]
                output_path = os.path.join(self.paths["transcripts"], f"{basename}.txt")
                self.after(0, lambda i=i, b=basename: (
                    self.transcribe_btn.configure(text=f"Batch ({i}/{len(files)})..."),
                    self._append_log(self.transcribe_log, f"\n[{i+1}/{len(files)}] {b}...")
                ))
                cmd = [sys.executable, "-u", os.path.join(BASE_DIR, "noScribe.py"),
                       audio_file, output_path, "--no-gui"]
                if lang_code != "auto":
                    cmd.extend(["--language", lang_code])
                if start_time: cmd.extend(["--start", start_time])
                if stop_time: cmd.extend(["--stop", stop_time])

                proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, bufsize=1, cwd=BASE_DIR)
                for line in proc.stdout:
                    stripped = line.rstrip()
                    if stripped:
                        self.after(0, lambda l=stripped: self._append_log(self.transcribe_log, l))
                rc = proc.wait()
                if rc == 0:
                    completed += 1
                    self.add_to_recent("transcripts", output_path)
                    self._add_history_entry(audio_file, output_path)
                    self.after(0, lambda p=output_path: self._append_log(self.transcribe_log, f"  ✅ Saved: {p}"))
                else:
                    self.after(0, lambda b=basename: self._append_log(self.transcribe_log, f"  ❌ Failed: {b}"))

            self.after(0, lambda: self._batch_done(completed, len(files)))

        self.run_in_thread(worker)

    def _batch_done(self, completed, total):
        elapsed = self._stop_elapsed_timer()
        self.transcribe_btn.configure(state="normal", text="Start", command=self.do_transcribe)
        self.transcribe_progress.stop()
        self.transcribe_progress.set(1 if completed == total else 0)
        self._append_log(self.transcribe_log, f"\n{'='*50}")
        self._append_log(self.transcribe_log, f"Batch complete: {completed}/{total} succeeded  ⏱️ {elapsed}")
        self.set_status(f"✅ Batch complete: {completed}/{total} ({elapsed})")
        if completed > 0:
            self.last_transcript_path = None  # multiple files, don't auto-populate
            self._add_open_folder_btn(self.transcribe_log.master, self.paths["transcripts"], row=4)
        self._batch_files = None

    def _add_history_entry(self, audio_file, transcript_path):
        """Log a completed transcription to history."""
        import datetime
        entry = {
            "audio": os.path.basename(audio_file),
            "transcript": transcript_path,
            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "language": self.transcribe_lang_var.get() if hasattr(self, 'transcribe_lang_var') else "Auto",
        }
        self.history.append(entry)
        save_config(self.paths, self.recent, self.history)

    # ═══════════════════════════════════════════════
    #  HISTORY SCREEN
    # ═══════════════════════════════════════════════
    def show_history(self):
        self.clear_main()
        self.make_title("📋  Transcription History")

        content = ctk.CTkFrame(self.main_panel, fg_color="transparent")
        content.grid(row=1, column=0, padx=20, pady=10, sticky="nswe")
        content.grid_rowconfigure(1, weight=1)
        content.grid_columnconfigure(0, weight=1)

        if not self.history:
            ctk.CTkLabel(content, text="No transcriptions yet.\nCompleted transcriptions will appear here.",
                font=ctk.CTkFont(size=14), text_color="#888888"
            ).grid(row=0, column=0, pady=40)
            return

        ctk.CTkLabel(content, text=f"{len(self.history)} transcription(s) recorded",
            font=ctk.CTkFont(size=13), text_color="#888888"
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        history_frame = ctk.CTkScrollableFrame(content, fg_color="transparent")
        history_frame.grid(row=1, column=0, sticky="nswe")
        history_frame.grid_columnconfigure(1, weight=1)

        for i, entry in enumerate(reversed(self.history)):
            row_frame = ctk.CTkFrame(history_frame, fg_color="gray20", corner_radius=8)
            row_frame.pack(fill="x", pady=3, padx=5)
            row_frame.grid_columnconfigure(1, weight=1)

            ctk.CTkLabel(row_frame, text=entry.get("timestamp", "?"),
                font=ctk.CTkFont(size=11), text_color="#888888", width=120
            ).grid(row=0, column=0, padx=(10, 5), pady=8, sticky="w")

            ctk.CTkLabel(row_frame, text=entry.get("audio", "?"),
                font=ctk.CTkFont(size=13, weight="bold"), anchor="w"
            ).grid(row=0, column=1, padx=5, pady=8, sticky="w")

            lang = entry.get("language", "")
            if lang and lang != "Auto":
                ctk.CTkLabel(row_frame, text=lang,
                    font=ctk.CTkFont(size=11), text_color="#aaaaaa", width=80
                ).grid(row=0, column=2, padx=5, pady=8)

            transcript = entry.get("transcript", "")
            if transcript and os.path.exists(transcript):
                ctk.CTkButton(row_frame, text="📂", width=35, height=28,
                    fg_color="gray40", hover_color="gray50",
                    command=lambda p=transcript: self.open_in_finder(p)
                ).grid(row=0, column=3, padx=(5, 10), pady=8)

        # Clear history button
        ctk.CTkButton(content, text="Clear History", height=35,
            font=ctk.CTkFont(size=12), fg_color="gray40", hover_color="gray50",
            command=self._clear_history
        ).grid(row=2, column=0, sticky="w", pady=(10, 0))

    def _clear_history(self):
        if messagebox.askyesno("Clear History", "Remove all history entries?"):
            self.history = []
            save_config(self.paths, self.recent, self.history)
            self.show_history()

    # ═══════════════════════════════════════════════
    #  RENAME SPEAKERS SCREEN
    # ═══════════════════════════════════════════════
    def show_rename(self):
        self.clear_main()
        self.make_title("🏷️  Rename Speakers")

        content = ctk.CTkFrame(self.main_panel, fg_color="transparent")
        content.grid(row=1, column=0, padx=20, pady=10, sticky="nswe")
        content.grid_rowconfigure(2, weight=1)
        content.grid_columnconfigure(0, weight=1)

        opts = ctk.CTkFrame(content, fg_color="transparent")
        opts.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        opts.grid_columnconfigure(1, weight=1)

        # Auto-populate from last transcription
        self.rename_file_var = ctk.StringVar(value=self.last_transcript_path or "")

        def browse_transcript():
            path = filedialog.askopenfilename(initialdir=self.paths["transcripts"], filetypes=[("Text Files", "*.txt")])
            if path:
                self.rename_file_var.set(path)
                self.add_to_recent("transcripts", path)

        self._make_file_picker(opts, self.rename_file_var, "Transcript:", browse_transcript, "transcripts", row=0)

        self.detect_btn = ctk.CTkButton(content, text="Detect Speakers", height=42,
            font=ctk.CTkFont(size=14, weight="bold"), command=self.do_detect_speakers)
        self.detect_btn.grid(row=1, column=0, sticky="w", pady=(0, 10))

        self.speakers_frame = ctk.CTkScrollableFrame(content, label_text="Detected Speakers")
        self.speakers_frame.grid(row=2, column=0, sticky="nswe")
        self.speakers_frame.grid_columnconfigure(1, weight=1)

        self.apply_rename_btn = ctk.CTkButton(content, text="Apply Renaming", height=42,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=("green", "green"), hover_color=("darkgreen", "darkgreen"),
            command=self.do_apply_rename)
        self.speaker_entries = {}

    def do_detect_speakers(self):
        filepath = self.rename_file_var.get().strip()
        if not filepath or not os.path.exists(filepath):
            messagebox.showwarning("No file", "Please select a transcript file.")
            return

        self.add_to_recent("transcripts", filepath)
        with open(filepath, "r", encoding="utf-8") as f:
            content = f.read()

        speaker_tags = sorted(set(re.findall(r'\b(S\d{2}):', content)))
        for widget in self.speakers_frame.winfo_children():
            widget.destroy()
        self.speaker_entries = {}

        if not speaker_tags:
            ctk.CTkLabel(self.speakers_frame,
                text="No speaker tags (S00, S01...) found.\nSpeakers may have already been renamed.",
                font=ctk.CTkFont(size=13), text_color="darkorange"
            ).grid(row=0, column=0, columnspan=3, pady=10)
            return

        for i, tag in enumerate(speaker_tags):
            pattern = re.compile(rf'^{tag}:\s*(.+)', re.MULTILINE)
            match = pattern.search(content)
            preview = match.group(1).strip()[:60] + "..." if match else "(no text)"

            ctk.CTkLabel(self.speakers_frame, text=f"{tag}  →",
                font=ctk.CTkFont(size=14, weight="bold")
            ).grid(row=i*2, column=0, padx=(10, 8), pady=(10, 0), sticky="w")

            entry = ctk.CTkEntry(self.speakers_frame, placeholder_text=f"New name for {tag}",
                height=33, corner_radius=8, border_width=2)
            entry.insert(0, tag)
            entry.grid(row=i*2, column=1, padx=(0, 10), pady=(10, 0), sticky="ew")
            self.speaker_entries[tag] = entry

            ctk.CTkLabel(self.speakers_frame, text=f'  "{preview}"',
                font=ctk.CTkFont(size=11), text_color="#777777"
            ).grid(row=i*2+1, column=0, columnspan=2, padx=15, pady=(0, 5), sticky="w")

        self.apply_rename_btn.grid(row=3, column=0, sticky="w", pady=(10, 0))
        self.set_status(f"Found {len(speaker_tags)} speakers")

    def do_apply_rename(self):
        filepath = self.rename_file_var.get().strip()
        if not filepath: return
        with open(filepath, "r", encoding="utf-8") as f:
            content = f.read()
        changes = 0
        for old_tag, entry in self.speaker_entries.items():
            new_name = entry.get().strip()
            if new_name and new_name != old_tag:
                content = content.replace(f"{old_tag}:", f"{new_name}:")
                changes += 1
        if changes == 0:
            messagebox.showinfo("No changes", "No speakers were renamed.")
            return
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(content)
        messagebox.showinfo("Success", f"✅ Renamed {changes} speaker(s) in:\n{os.path.basename(filepath)}")
        self.set_status(f"✅ Renamed {changes} speakers")
        self.do_detect_speakers()

    # ═══════════════════════════════════════════════
    #  EXCEL + THEME SCREEN
    # ═══════════════════════════════════════════════
    def show_excel(self):
        self.clear_main()
        self.make_title("📊  Generate Excel Database")

        content = ctk.CTkFrame(self.main_panel, fg_color="transparent")
        content.grid(row=1, column=0, padx=20, pady=10, sticky="nswe")
        content.grid_rowconfigure(4, weight=1)
        content.grid_columnconfigure(0, weight=1)

        opts = ctk.CTkFrame(content, fg_color="transparent")
        opts.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        opts.grid_columnconfigure(1, weight=1)

        # Auto-populate from last transcription
        self.excel_file_var = ctk.StringVar(value=self.last_transcript_path or "")

        def browse_transcript():
            path = filedialog.askopenfilename(initialdir=self.paths["transcripts"], filetypes=[("Text Files", "*.txt")])
            if path:
                self.excel_file_var.set(path)
                self.add_to_recent("transcripts", path)

        self._make_file_picker(opts, self.excel_file_var, "Transcript:", browse_transcript, "transcripts", row=0)

        # Theme dropdown
        ctk.CTkLabel(opts, text="Color theme:", font=ctk.CTkFont(size=14)).grid(row=1, column=0, sticky="w", pady=5, padx=(0, 10))
        self.theme_var = ctk.StringVar(value="Professional Blue")
        ctk.CTkOptionMenu(opts, variable=self.theme_var, values=list(THEMES.keys()),
            width=200, dynamic_resizing=False).grid(row=1, column=1, sticky="w", pady=5)

        # Output options
        options_frame = ctk.CTkFrame(content, fg_color="gray20", corner_radius=8)
        options_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        ctk.CTkLabel(options_frame, text="Output Options",
            font=ctk.CTkFont(size=13, weight="bold"), text_color="#aaaaaa"
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 4))

        self.excel_opt_summary = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(options_frame, text="Results Summary sheet",
            variable=self.excel_opt_summary, font=ctk.CTkFont(size=13)
        ).grid(row=1, column=0, sticky="w", padx=12, pady=3)

        self.excel_opt_wordfreq = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(options_frame, text="Word Frequency sheet  (top 30 words per speaker)",
            variable=self.excel_opt_wordfreq, font=ctk.CTkFont(size=13)
        ).grid(row=2, column=0, sticky="w", padx=12, pady=3)

        self.excel_opt_questions = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(options_frame, text="Question detection column  (flag turns with ?)",
            variable=self.excel_opt_questions, font=ctk.CTkFont(size=13)
        ).grid(row=3, column=0, sticky="w", padx=12, pady=3)

        self.excel_opt_csv = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(options_frame, text="Also export as CSV  (.csv alongside .xlsx)",
            variable=self.excel_opt_csv, font=ctk.CTkFont(size=13)
        ).grid(row=4, column=0, sticky="w", padx=12, pady=3)

        self.excel_opt_latex = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(options_frame, text="Export LaTeX tables  (.tex for academic papers)",
            variable=self.excel_opt_latex, font=ctk.CTkFont(size=13)
        ).grid(row=5, column=0, sticky="w", padx=12, pady=3)

        self.excel_opt_lexical = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(options_frame, text="Compute lexical diversity (TTR, MTLD, Readability)",
            variable=self.excel_opt_lexical, font=ctk.CTkFont(size=13)
        ).grid(row=6, column=0, sticky="w", padx=12, pady=(3, 10))

        self.excel_btn = ctk.CTkButton(content, text="Start", height=42, width=200,
            font=ctk.CTkFont(size=14, weight="bold"), command=self.do_generate_excel)
        self.excel_btn.grid(row=2, column=0, sticky="w", pady=(0, 10))

        self.excel_log = self.make_log(content)
        self.excel_log.grid(row=4, column=0, sticky="nswe")

    def do_generate_excel(self):
        filepath = self.excel_file_var.get().strip()
        if not filepath or not os.path.exists(filepath):
            messagebox.showwarning("No file", "Please select a transcript file.")
            return

        self.add_to_recent("transcripts", filepath)
        theme_key = THEMES[self.theme_var.get()]
        basename = os.path.splitext(os.path.basename(filepath))[0]
        output_path = os.path.join(self.paths["exports"], f"{basename}.xlsx")

        self.excel_btn.configure(state="disabled", text="Generating...")
        self.set_status("Building Excel database...")
        self.excel_log.configure(state="normal")
        self.excel_log.delete("0.0", "end")
        self.excel_log.insert("end", f"Input:  {filepath}\nOutput: {output_path}\nTheme:  {self.theme_var.get()}\n\nGenerating...\n")

        def worker():
            cmd = [sys.executable, os.path.join(ANALYSIS_DIR, "parse_to_excel.py"),
                 filepath, output_path, "--theme", theme_key]
            if not self.excel_opt_summary.get():
                cmd.append("--no-summary")
            if self.excel_opt_wordfreq.get():
                cmd.append("--word-freq")
            if self.excel_opt_questions.get():
                cmd.append("--questions")
            if self.excel_opt_csv.get():
                cmd.append("--csv")
            if self.excel_opt_latex.get():
                cmd.append("--latex")
            if self.excel_opt_lexical.get():
                cmd.append("--lexical")
            result = subprocess.run(cmd, capture_output=True, text=True, cwd=BASE_DIR)
            self.after(0, lambda: self._excel_done(result, output_path))

        self.run_in_thread(worker)

    def _excel_done(self, result, output_path):
        self.excel_btn.configure(state="normal", text="Start")
        if result.returncode == 0:
            self.excel_log.insert("end", f"\n\n✅ Excel saved to:\n{output_path}")
            self.set_status("✅ Excel database generated!")
            self.last_excel_path = output_path
            self.add_to_recent("exports", output_path)
            self._add_open_folder_btn(self.excel_log.master, output_path, row=4)
        else:
            err = result.stderr or result.stdout or "Unknown error"
            self.excel_log.insert("end", f"\n\n❌ Error:\n{err}")
            self.set_status("❌ Excel generation failed")

    # ═══════════════════════════════════════════════
    #  GRAPHS SCREEN
    # ═══════════════════════════════════════════════
    def show_graphs(self):
        self.clear_main()
        self.make_title("📈  Generate Visualizations")

        content = ctk.CTkFrame(self.main_panel, fg_color="transparent")
        content.grid(row=1, column=0, padx=20, pady=10, sticky="nswe")
        content.grid_rowconfigure(4, weight=1)
        content.grid_columnconfigure(0, weight=1)

        opts = ctk.CTkFrame(content, fg_color="transparent")
        opts.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        opts.grid_columnconfigure(1, weight=1)

        # Auto-populate from last Excel
        self.graphs_file_var = ctk.StringVar(value=self.last_excel_path or "")

        def browse_excel():
            path = filedialog.askopenfilename(initialdir=self.paths["exports"], filetypes=[("Excel Files", "*.xlsx")])
            if path:
                self.graphs_file_var.set(path)
                self.add_to_recent("exports", path)

        self._make_file_picker(opts, self.graphs_file_var, "Excel file:", browse_excel, "exports", row=0)

        # Charts to generate
        charts_frame = ctk.CTkFrame(content, fg_color="gray20", corner_radius=8)
        charts_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        ctk.CTkLabel(charts_frame, text="Charts to Generate",
            font=ctk.CTkFont(size=13, weight="bold"), text_color="#aaaaaa"
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 4))

        self.graph_opt_timeline = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(charts_frame, text="Conversational Timeline  (word count per turn)",
            variable=self.graph_opt_timeline, font=ctk.CTkFont(size=13)
        ).grid(row=1, column=0, sticky="w", padx=12, pady=3)

        self.graph_opt_friction = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(charts_frame, text="Friction Scatterplot  (disfluencies & pauses)",
            variable=self.graph_opt_friction, font=ctk.CTkFont(size=13)
        ).grid(row=2, column=0, sticky="w", padx=12, pady=3)

        self.graph_opt_pie = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(charts_frame, text="Speaker Balance Pie Chart  (% of words & turns)",
            variable=self.graph_opt_pie, font=ctk.CTkFont(size=13)
        ).grid(row=3, column=0, sticky="w", padx=12, pady=3)

        self.graph_opt_dist = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(charts_frame, text="Turn Length Distribution  (histogram)",
            variable=self.graph_opt_dist, font=ctk.CTkFont(size=13)
        ).grid(row=4, column=0, sticky="w", padx=12, pady=3)

        self.graph_opt_matrix = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(charts_frame, text="Turn-Taking Matrix  (who responds to whom)",
            variable=self.graph_opt_matrix, font=ctk.CTkFont(size=13)
        ).grid(row=5, column=0, sticky="w", padx=12, pady=3)

        self.graph_opt_wordcloud = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(charts_frame, text="Word Clouds  (visual word frequency per speaker)",
            variable=self.graph_opt_wordcloud, font=ctk.CTkFont(size=13)
        ).grid(row=6, column=0, sticky="w", padx=12, pady=(3, 10))

        # Export settings
        export_frame = ctk.CTkFrame(content, fg_color="gray20", corner_radius=8)
        export_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))

        ctk.CTkLabel(export_frame, text="Export Settings",
            font=ctk.CTkFont(size=13, weight="bold"), text_color="#aaaaaa"
        ).grid(row=0, column=0, columnspan=4, sticky="w", padx=12, pady=(8, 4))

        ctk.CTkLabel(export_frame, text="Format:", font=ctk.CTkFont(size=13)).grid(row=1, column=0, padx=(12, 5), pady=(3, 10), sticky="w")
        self.graph_format_var = ctk.StringVar(value="PNG")
        ctk.CTkOptionMenu(export_frame, variable=self.graph_format_var,
            values=["PNG", "SVG", "PDF"], width=100, dynamic_resizing=False
        ).grid(row=1, column=1, padx=(0, 20), pady=(3, 10), sticky="w")

        ctk.CTkLabel(export_frame, text="DPI:", font=ctk.CTkFont(size=13)).grid(row=1, column=2, padx=(0, 5), pady=(3, 10), sticky="w")
        self.graph_dpi_var = ctk.StringVar(value="300")
        ctk.CTkOptionMenu(export_frame, variable=self.graph_dpi_var,
            values=["150", "300", "600"], width=80, dynamic_resizing=False
        ).grid(row=1, column=3, padx=(0, 12), pady=(3, 10), sticky="w")

        self.graphs_btn = ctk.CTkButton(content, text="Start", height=42, width=200,
            font=ctk.CTkFont(size=14, weight="bold"), command=self.do_generate_graphs)
        self.graphs_btn.grid(row=3, column=0, sticky="w", pady=(0, 10))

        self.graphs_log = self.make_log(content)
        self.graphs_log.grid(row=4, column=0, sticky="nswe")

    def do_generate_graphs(self):
        filepath = self.graphs_file_var.get().strip()
        if not filepath or not os.path.exists(filepath):
            messagebox.showwarning("No file", "Please select an Excel file.")
            return

        self.add_to_recent("exports", filepath)
        basename = os.path.splitext(os.path.basename(filepath))[0]
        output_dir = os.path.join(self.paths["graphs"], basename)

        self.graphs_btn.configure(state="disabled", text="Generating...")
        self.set_status("Drawing graphs...")
        self.graphs_log.configure(state="normal")
        self.graphs_log.delete("0.0", "end")
        self.graphs_log.insert("end", f"Input:  {filepath}\nOutput: {output_dir}/\n\nGenerating...\n")

        def worker():
            cmd = [sys.executable, os.path.join(ANALYSIS_DIR, "visualize_data.py"),
                 filepath, output_dir]
            # Add selected chart flags
            if self.graph_opt_timeline.get(): cmd.append("--timeline")
            if self.graph_opt_friction.get(): cmd.append("--friction")
            if self.graph_opt_pie.get(): cmd.append("--pie")
            if self.graph_opt_dist.get(): cmd.append("--distribution")
            if self.graph_opt_matrix.get(): cmd.append("--matrix")
            if self.graph_opt_wordcloud.get(): cmd.append("--wordcloud")
            # If none selected, script defaults to all
            cmd.extend(["--format", self.graph_format_var.get().lower()])
            cmd.extend(["--dpi", self.graph_dpi_var.get()])
            result = subprocess.run(cmd, capture_output=True, text=True, cwd=BASE_DIR)
            self.after(0, lambda: self._graphs_done(result, output_dir))

        self.run_in_thread(worker)

    def _graphs_done(self, result, output_dir):
        self.graphs_btn.configure(state="normal", text="Start")
        if result.returncode == 0:
            self.graphs_log.insert("end", f"\n\n✅ Graphs saved to:\n{output_dir}/")
            self.set_status("✅ Visualizations generated!")
            self._add_open_folder_btn(self.graphs_log.master, output_dir, row=3)
        else:
            err = result.stderr or result.stdout or "Unknown error"
            self.graphs_log.insert("end", f"\n\n❌ Error:\n{err}")
            self.set_status("❌ Visualization failed")

    # ═══════════════════════════════════════════════
    #  FULL PIPELINE SCREEN
    # ═══════════════════════════════════════════════
    def show_full_pipeline(self):
        self.clear_main()
        self.make_title("🚀  Full Pipeline")

        content = ctk.CTkFrame(self.main_panel, fg_color="transparent")
        content.grid(row=1, column=0, padx=20, pady=10, sticky="nswe")
        content.grid_rowconfigure(3, weight=1)
        content.grid_columnconfigure(0, weight=1)

        opts = ctk.CTkFrame(content, fg_color="transparent")
        opts.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        opts.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(opts, text="Audio source:", font=ctk.CTkFont(size=14)).grid(row=0, column=0, sticky="w", pady=5, padx=(0, 10))
        self.pipeline_source_var = ctk.StringVar(value="existing")
        ctk.CTkOptionMenu(opts, variable=self.pipeline_source_var,
            values=["existing", "youtube"], width=150, dynamic_resizing=False).grid(row=0, column=1, sticky="w", pady=5)

        ctk.CTkLabel(opts, text="YouTube URL:", font=ctk.CTkFont(size=14)).grid(row=1, column=0, sticky="w", pady=5, padx=(0, 10))
        self.pipeline_url_entry = ctk.CTkEntry(opts, placeholder_text="(only if downloading)", height=33, corner_radius=8, border_width=2)
        self.pipeline_url_entry.grid(row=1, column=1, sticky="ew", pady=5)

        # Existing file picker with recent
        self.pipeline_file_var = ctk.StringVar()
        def browse_mp3():
            path = filedialog.askopenfilename(initialdir=self.paths["videos"], filetypes=[("MP3 Files", "*.mp3")])
            if path:
                self.pipeline_file_var.set(path)
                self.add_to_recent("videos", path)

        self._make_file_picker(opts, self.pipeline_file_var, "Existing .mp3:", browse_mp3, "videos", row=2)

        ctk.CTkLabel(opts, text="Excel theme:", font=ctk.CTkFont(size=14)).grid(row=3, column=0, sticky="w", pady=5, padx=(0, 10))
        self.pipeline_theme_var = ctk.StringVar(value="Professional Blue")
        ctk.CTkOptionMenu(opts, variable=self.pipeline_theme_var,
            values=list(THEMES.keys()), width=200, dynamic_resizing=False).grid(row=3, column=1, sticky="w", pady=5)

        self.pipeline_btn = ctk.CTkButton(content, text="Start", height=42, width=200,
            font=ctk.CTkFont(size=14, weight="bold"), command=self.do_full_pipeline)
        self.pipeline_btn.grid(row=1, column=0, sticky="w", pady=(0, 10))

        self.pipeline_log = self.make_log(content)
        self.pipeline_log.grid(row=3, column=0, sticky="nswe")

    def do_full_pipeline(self):
        source = self.pipeline_source_var.get()
        theme_key = THEMES[self.pipeline_theme_var.get()]

        if source == "youtube":
            url = self.pipeline_url_entry.get().strip()
            if len(url) < 10:
                messagebox.showwarning("No URL", "Please enter a YouTube URL.")
                return
        else:
            audio_path = self.pipeline_file_var.get().strip()
            if not audio_path or not os.path.exists(audio_path):
                messagebox.showwarning("No file", "Please select an .mp3 file.")
                return

        self.pipeline_btn.configure(state="disabled", text="Running...")
        self.pipeline_log.configure(state="normal")
        self.pipeline_log.delete("0.0", "end")

        def worker():
            try:
                audio_path_local = None

                if source == "youtube":
                    self.after(0, lambda: self.set_status("Phase 1: Downloading audio..."))
                    self.after(0, lambda: self.pipeline_log.insert("end", "📥 Phase 1: Downloading audio...\n"))

                    existing = set(glob.glob(os.path.join(self.paths["videos"], "*.mp3")))
                    result = subprocess.run(
                        [os.path.join(BASE_DIR, "download_audio.sh"), url, self.paths["videos"]],
                        capture_output=True, text=True, cwd=BASE_DIR)
                    if result.returncode != 0:
                        self.after(0, lambda: self._pipeline_error("Download failed", result.stderr))
                        return

                    new_files = set(glob.glob(os.path.join(self.paths["videos"], "*.mp3"))) - existing
                    if new_files:
                        audio_path_local = new_files.pop()
                    else:
                        all_mp3 = sorted(glob.glob(os.path.join(self.paths["videos"], "*.mp3")), key=os.path.getmtime, reverse=True)
                        audio_path_local = all_mp3[0] if all_mp3 else None

                    if not audio_path_local:
                        self.after(0, lambda: self._pipeline_error("Download failed", "No .mp3 file found"))
                        return
                    self.after(0, lambda: self.pipeline_log.insert("end", f"   ✅ Downloaded: {os.path.basename(audio_path_local)}\n\n"))
                    self.add_to_recent("videos", audio_path_local)
                else:
                    audio_path_local = audio_path
                    self.add_to_recent("videos", audio_path_local)

                basename = os.path.splitext(os.path.basename(audio_path_local))[0]
                transcript_path = os.path.join(self.paths["transcripts"], f"{basename}.txt")
                excel_path = os.path.join(self.paths["exports"], f"{basename}.xlsx")
                graphs_path = os.path.join(self.paths["graphs"], basename)

                # Phase 2: Transcribe
                self.after(0, lambda: self.set_status("Phase 2: Transcribing..."))
                self.after(0, lambda: self.pipeline_log.insert("end", "🎙️ Phase 2: Transcribing...\n"))
                result = subprocess.run(
                    [sys.executable, os.path.join(BASE_DIR, "noScribe.py"),
                     audio_path_local, transcript_path, "--no-gui"],
                    capture_output=True, text=True, cwd=BASE_DIR)
                if result.returncode != 0:
                    self.after(0, lambda: self._pipeline_error("Transcription failed", result.stderr or result.stdout))
                    return
                self.after(0, lambda: self.pipeline_log.insert("end", f"   ✅ Transcript: {transcript_path}\n\n"))
                self.last_transcript_path = transcript_path
                self.add_to_recent("transcripts", transcript_path)

                # Phase 3: Excel
                self.after(0, lambda: self.set_status("Phase 3: Generating Excel..."))
                self.after(0, lambda: self.pipeline_log.insert("end", f"📊 Phase 3: Excel ({self.pipeline_theme_var.get()})...\n"))
                result = subprocess.run(
                    [sys.executable, os.path.join(ANALYSIS_DIR, "parse_to_excel.py"),
                     transcript_path, excel_path, "--theme", theme_key],
                    capture_output=True, text=True, cwd=BASE_DIR)
                if result.returncode != 0:
                    self.after(0, lambda: self._pipeline_error("Excel generation failed", result.stderr or result.stdout))
                    return
                self.after(0, lambda: self.pipeline_log.insert("end", f"   ✅ Excel: {excel_path}\n\n"))
                self.last_excel_path = excel_path
                self.add_to_recent("exports", excel_path)

                # Phase 4: Graphs
                self.after(0, lambda: self.set_status("Phase 4: Drawing visualizations..."))
                self.after(0, lambda: self.pipeline_log.insert("end", "📈 Phase 4: Generating graphs...\n"))
                result = subprocess.run(
                    [sys.executable, os.path.join(ANALYSIS_DIR, "visualize_data.py"),
                     excel_path, graphs_path],
                    capture_output=True, text=True, cwd=BASE_DIR)
                if result.returncode != 0:
                    self.after(0, lambda: self._pipeline_error("Visualization failed", result.stderr or result.stdout))
                    return
                self.after(0, lambda: self.pipeline_log.insert("end", f"   ✅ Graphs: {graphs_path}/\n\n"))

                self.after(0, lambda: self.pipeline_log.insert("end", "🎉 Pipeline Complete!\n"))
                self.after(0, lambda: self.set_status("🎉 Full pipeline complete!"))
                self.after(0, lambda: self.pipeline_btn.configure(state="normal", text="Start"))
                # Open folder for the graphs (last output)
                self.after(0, lambda: self._add_open_folder_btn(self.pipeline_log.master, graphs_path, row=4))

            except Exception as e:
                self.after(0, lambda: self._pipeline_error("Unexpected error", str(e)))

        self.run_in_thread(worker)

    def _pipeline_error(self, phase, error_msg):
        self.pipeline_log.insert("end", f"\n\n❌ {phase}:\n{error_msg}")
        self.set_status(f"❌ Pipeline stopped: {phase}")
        self.pipeline_btn.configure(state="normal", text="Start")


# ═══════════════════════════════════════════════
#  LAUNCH
# ═══════════════════════════════════════════════
if __name__ == "__main__":
    app = PipelineApp()
    app.mainloop()
