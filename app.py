import sys
import os
import json
import math
from datetime import datetime
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

if getattr(sys, "frozen", False):
    sys.path.insert(0, sys._MEIPASS)

from engine import ComparisonEngine
from i18n import t, set_lang, get_lang, LANGUAGES

SETTINGS_FILE = Path(__file__).resolve().parent / "settings.json"


SETTINGS_DEFAULTS = {
    "lang": "tr",
    "theme": "dark",
    "font_size": 13,
    "recent_files": [],
    "default_ignore_case": False,
    "default_strip_spaces": False,
    "default_export_dir": "",
    "preview_row_limit": 200,
    "csv_encoding": "utf-8-sig",
    "max_recent": 20,
    "window_geometry": "",
}


def _load_settings():
    defaults = dict(SETTINGS_DEFAULTS)
    try:
        if SETTINGS_FILE.exists():
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            defaults.update(saved)
    except Exception:
        pass
    return defaults


def _save_settings(**kwargs):
    current = _load_settings()
    current.update(kwargs)
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(current, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _add_to_recent(paths):
    prefs = _load_settings()
    recent = prefs.get("recent_files", [])
    limit = prefs.get("max_recent", 20)
    for p in paths:
        p = str(p)
        if p in recent:
            recent.remove(p)
        recent.insert(0, p)
    _save_settings(recent_files=recent[:limit])


def _get_recent_files():
    prefs = _load_settings()
    return [p for p in prefs.get("recent_files", []) if Path(p).exists()]


HISTORY_FILE = Path(__file__).resolve().parent / "history.json"


def _load_history():
    try:
        if HISTORY_FILE.exists():
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return []


def _save_history_entry(entry):
    history = _load_history()
    history.insert(0, entry)
    history = history[:100]
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _clear_history_log():
    try:
        if HISTORY_FILE.exists():
            HISTORY_FILE.unlink()
    except Exception:
        pass


PROFILES_FILE = Path(__file__).resolve().parent / "profiles.json"


def _load_profiles():
    try:
        if PROFILES_FILE.exists():
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return []


def _save_profiles(profiles):
    try:
        with open(PROFILES_FILE, "w", encoding="utf-8") as f:
            json.dump(profiles, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _add_profile(profile):
    profiles = _load_profiles()
    profiles = [p for p in profiles if p["name"] != profile["name"]]
    profiles.insert(0, profile)
    _save_profiles(profiles)


def _delete_profile(name):
    profiles = _load_profiles()
    profiles = [p for p in profiles if p["name"] != name]
    _save_profiles(profiles)

DND_AVAILABLE = False

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _DND_IMPORT_OK = True
except Exception:
    _DND_IMPORT_OK = False

if _DND_IMPORT_OK:
    class _BaseCTk(ctk.CTk, TkinterDnD.DnDWrapper):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            global DND_AVAILABLE
            try:
                self.TkdndVersion = TkinterDnD._require(self)
                DND_AVAILABLE = True
            except Exception:
                DND_AVAILABLE = False
else:
    _BaseCTk = ctk.CTk

ctk.set_default_color_theme("blue")

SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".csv"}


class PreviewWindow(ctk.CTkToplevel):

    def __init__(self, master, file_path, engine):
        super().__init__(master)
        self.MAX_PREVIEW_ROWS = _load_settings().get("preview_row_limit", 200)
        self.title(t("preview_title", name=Path(file_path).name))
        self.geometry("800x500")
        self.minsize(500, 300)
        self._build(file_path, engine)

    def _build(self, file_path, engine):
        try:
            sheets = engine.get_sheet_names(file_path)
            if sheets:
                top = ctk.CTkFrame(self, fg_color="transparent")
                top.pack(fill="x", padx=10, pady=(10, 0))
                ctk.CTkLabel(top, text=t("sheet_label")).pack(side="left", padx=(0, 5))
                self._sheet_var = ctk.StringVar(value=sheets[0])
                ctk.CTkOptionMenu(
                    top, variable=self._sheet_var, values=sheets,
                    command=lambda _: self._load_table(file_path, engine),
                ).pack(side="left")

            import tkinter as tk
            table_frame = ctk.CTkFrame(self)
            table_frame.pack(fill="both", expand=True, padx=10, pady=10)

            style = __import__("tkinter.ttk", fromlist=["ttk"]).Style()
            style.configure("Preview.Treeview", rowheight=24)

            self._tree = __import__("tkinter.ttk", fromlist=["ttk"]).Treeview(
                table_frame, style="Preview.Treeview", show="headings"
            )

            vsb = tk.Scrollbar(table_frame, orient="vertical", command=self._tree.yview)
            hsb = tk.Scrollbar(table_frame, orient="horizontal", command=self._tree.xview)
            self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

            self._tree.grid(row=0, column=0, sticky="nsew")
            vsb.grid(row=0, column=1, sticky="ns")
            hsb.grid(row=1, column=0, sticky="ew")
            table_frame.grid_rowconfigure(0, weight=1)
            table_frame.grid_columnconfigure(0, weight=1)

            self._info_label = ctk.CTkLabel(self, text="", text_color="gray", font=("", 11))
            self._info_label.pack(pady=(0, 8))

            self._load_table(file_path, engine)
        except Exception as e:
            ctk.CTkLabel(
                self, text=t("file_read_error", e=e), text_color="red"
            ).pack(expand=True, padx=20, pady=20)

    def _load_table(self, file_path, engine):
        sheet = getattr(self, "_sheet_var", None)
        sheet = sheet.get() if sheet else None

        try:
            df = engine.load_file(file_path, sheet)
        except Exception as e:
            self._info_label.configure(text=f"{t('error_title')}: {e}")
            return

        self._tree.delete(*self._tree.get_children())

        cols = list(df.columns.astype(str))
        self._tree["columns"] = cols
        for col in cols:
            self._tree.heading(col, text=col)
            max_width = max(len(col) * 10, 80)
            self._tree.column(col, width=min(max_width, 200), minwidth=60)

        total = len(df)
        preview_df = df.head(self.MAX_PREVIEW_ROWS)
        for _, row in preview_df.iterrows():
            values = [str(v) if str(v) != "nan" else "" for v in row]
            self._tree.insert("", "end", values=values)

        if total > self.MAX_PREVIEW_ROWS:
            self._info_label.configure(
                text=t("preview_showing", limit=self.MAX_PREVIEW_ROWS, total=f"{total:,}")
            )
        else:
            self._info_label.configure(text=t("preview_total", total=f"{total:,}"))


class FileEntry(ctk.CTkFrame):

    def __init__(self, master, file_path, on_remove, on_preview=None, **kwargs):
        super().__init__(master, **kwargs)
        self.file_path = file_path
        self.on_remove = on_remove
        self.grid_columnconfigure(0, weight=1)

        name_label = ctk.CTkLabel(
            self, text=Path(file_path).name, anchor="w",
            cursor="hand2",
        )
        name_label.grid(row=0, column=0, padx=5, sticky="ew")
        if on_preview:
            name_label.bind("<Button-1>", lambda e: on_preview(file_path))

        ctk.CTkButton(
            self, text="👁", width=30, height=24,
            fg_color="#2980b9", hover_color="#3498db",
            command=lambda: on_preview(file_path) if on_preview else None,
        ).grid(row=0, column=1, padx=2, pady=2)
        ctk.CTkButton(
            self, text="✕", width=30, height=24,
            fg_color="#c0392b", hover_color="#e74c3c",
            command=lambda: self.on_remove(self),
        ).grid(row=0, column=2, padx=(0, 5), pady=2)


class App(_BaseCTk):
    def __init__(self):
        super().__init__()

        prefs = _load_settings()
        set_lang(prefs["lang"])
        ctk.set_appearance_mode(prefs["theme"])
        self._font_size = prefs["font_size"]
        self._prefs = prefs

        self.title(t("app_title"))
        geo = prefs.get("window_geometry", "") or "1100x750"
        self.geometry(geo)
        self.minsize(900, 650)
        self._saving_geo = False
        self.bind("<Configure>", self._on_window_configure)

        self.engine = ComparisonEngine()
        self.ref_files = []
        self.comp_files = []
        self.results = None
        self._last_comp_time = None
        self._last_comp_records = 0

        self._build_ui()

    def _on_window_configure(self, event):
        if event.widget is self and not self._saving_geo:
            self._saving_geo = True
            self.after(500, self._save_geometry)

    def _save_geometry(self):
        self._saving_geo = False
        if hasattr(self, "_save_geo_var") and not self._save_geo_var.get():
            return
        geo = self.geometry()
        if geo:
            _save_settings(window_geometry=geo)

    # ── Drag & Drop Yardımcıları ─────────────────────────────────

    def _parse_drop_data(self, data):
        files = []
        i = 0
        while i < len(data):
            if data[i] == "{":
                end = data.index("}", i)
                files.append(data[i + 1 : end])
                i = end + 2
            elif data[i] == " ":
                i += 1
            else:
                end = data.find(" ", i)
                if end == -1:
                    files.append(data[i:])
                    break
                files.append(data[i:end])
                i = end + 1
        return [f for f in files if Path(f).suffix.lower() in SUPPORTED_EXTENSIONS]

    def _enable_drop(self, widget, callback):
        if not DND_AVAILABLE:
            return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", callback)
        except Exception:
            pass

    # ── Arayüz Oluşturma ──────────────────────────────────────────

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._build_header()

        self._main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._main_frame.grid(row=1, column=0, sticky="nsew")
        self._main_frame.grid_columnconfigure(0, weight=1)
        self._main_frame.grid_rowconfigure(1, weight=1)

        self._settings_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._settings_frame.grid_columnconfigure(0, weight=1)

        self._history_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._history_frame.grid_columnconfigure(0, weight=1)
        self._history_frame.grid_rowconfigure(0, weight=1)

        self._build_file_section()
        self._build_results_section()
        self._build_bottom_bar()
        self._build_status_bar()
        self._build_settings_page()

        self._showing_settings = False
        self._showing_history = False

    def _apply_font_size(self):
        import tkinter.font as tkfont
        font = tkfont.Font(family="", size=self._font_size)
        for attr in ("txt_match", "txt_no_match", "txt_duplicates",
                      "txt_ref", "txt_comp", "txt_stats"):
            textbox = getattr(self, attr, None)
            if textbox:
                textbox._textbox.configure(font=font)

    def _show_main(self):
        self._settings_frame.grid_remove()
        self._history_frame.grid_remove()
        self._main_frame.grid(row=1, column=0, sticky="nsew")
        self._settings_btn.configure(text=t("settings_btn"))
        self._history_btn.configure(text=t("history_btn"))
        self._showing_settings = False
        self._showing_history = False

    def _toggle_settings(self):
        if self._showing_settings:
            self._show_main()
        else:
            self._main_frame.grid_remove()
            self._history_frame.grid_remove()
            self._settings_frame.grid(row=1, column=0, sticky="nsew", padx=15, pady=10)
            self._settings_btn.configure(text=t("settings_btn_back"))
            self._history_btn.configure(text=t("history_btn"))
            self._showing_settings = True
            self._showing_history = False

    def _toggle_history(self):
        if self._showing_history:
            self._show_main()
        else:
            self._main_frame.grid_remove()
            self._settings_frame.grid_remove()
            self._build_history_page()
            self._history_frame.grid(row=1, column=0, sticky="nsew", padx=15, pady=10)
            self._history_btn.configure(text=t("settings_btn_back"))
            self._settings_btn.configure(text=t("settings_btn"))
            self._showing_history = True
            self._showing_settings = False

    def _build_history_page(self):
        for w in self._history_frame.winfo_children():
            w.destroy()

        ctk.CTkLabel(
            self._history_frame, text=t("history_title"),
            font=("", 22, "bold"),
        ).pack(anchor="w", padx=5, pady=(0, 5))

        toolbar = ctk.CTkFrame(self._history_frame, fg_color="transparent")
        toolbar.pack(fill="x", padx=5, pady=(0, 8))

        ctk.CTkButton(
            toolbar, text=t("history_clear_btn"), width=120, height=30,
            fg_color="#c0392b", hover_color="#e74c3c",
            command=self._clear_comparison_history,
        ).pack(side="right")

        history = _load_history()

        if not history:
            ctk.CTkLabel(
                self._history_frame, text=t("history_empty"),
                text_color="gray", font=("", 14),
            ).pack(expand=True)
            return

        scroll = ctk.CTkScrollableFrame(self._history_frame)
        scroll.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        scroll.grid_columnconfigure(0, weight=1)

        for i, entry in enumerate(history):
            card = ctk.CTkFrame(scroll, corner_radius=8)
            card.grid(row=i, column=0, sticky="ew", pady=4)
            card.grid_columnconfigure(1, weight=1)

            date_str = entry.get("date", "?")
            match_c = entry.get("match_count", 0)
            only_ref_c = entry.get("only_ref", 0)
            only_comp_c = entry.get("only_comp", 0)
            pct = entry.get("match_pct", 0)
            ref_names = entry.get("ref_files", [])
            comp_names = entry.get("comp_files", [])

            if pct >= 80:
                pct_color = "#27ae60"
            elif pct >= 50:
                pct_color = "#f39c12"
            else:
                pct_color = "#e74c3c"

            pct_label = ctk.CTkLabel(
                card, text=f"%{pct}", font=("", 22, "bold"),
                text_color=pct_color, width=70,
            )
            pct_label.grid(row=0, column=0, rowspan=2, padx=(12, 8), pady=10)

            info_top = ctk.CTkLabel(
                card, font=("", 12),
                text=f"📅 {date_str}   ·   ✅ {match_c:,}  ❌ {only_ref_c + only_comp_c:,}",
                anchor="w",
            )
            info_top.grid(row=0, column=1, sticky="w", padx=5, pady=(10, 0))

            ref_short = ", ".join(ref_names[:3])
            comp_short = ", ".join(comp_names[:3])
            if len(ref_names) > 3:
                ref_short += f" +{len(ref_names)-3}"
            if len(comp_names) > 3:
                comp_short += f" +{len(comp_names)-3}"

            info_bottom = ctk.CTkLabel(
                card, font=("", 11), text_color="gray",
                text=f"{t('history_ref')}: {ref_short}  →  {t('history_comp')}: {comp_short}",
                anchor="w",
            )
            info_bottom.grid(row=1, column=1, sticky="w", padx=5, pady=(0, 10))

    def _clear_comparison_history(self):
        _clear_history_log()
        self._build_history_page()

    def _build_settings_page(self):
        sf = self._settings_frame
        prefs = _load_settings()

        ctk.CTkLabel(sf, text=t("settings_title"), font=("", 22, "bold")).pack(
            anchor="w", pady=(15, 10), padx=5
        )

        scroll = ctk.CTkScrollableFrame(sf)
        scroll.pack(fill="both", expand=True)

        # ── Dil ──
        lang_frame = ctk.CTkFrame(scroll)
        lang_frame.pack(fill="x", pady=5)

        ctk.CTkLabel(lang_frame, text=t("lang_title"), font=("", 15, "bold")).pack(
            anchor="w", padx=15, pady=(12, 5)
        )
        ctk.CTkLabel(
            lang_frame, text=t("lang_desc"),
            text_color="gray", font=("", 11),
        ).pack(anchor="w", padx=15)

        lang_labels = list(LANGUAGES.values())
        current_label = LANGUAGES.get(get_lang(), "Türkçe")
        self._lang_var = ctk.StringVar(value=current_label)

        ctk.CTkSegmentedButton(
            lang_frame, values=lang_labels,
            variable=self._lang_var,
            command=self._on_lang_change,
        ).pack(fill="x", padx=15, pady=(8, 12))

        # ── Tema ──
        tema_frame = ctk.CTkFrame(scroll)
        tema_frame.pack(fill="x", pady=5)

        ctk.CTkLabel(tema_frame, text=t("theme_title"), font=("", 15, "bold")).pack(
            anchor="w", padx=15, pady=(12, 5)
        )
        ctk.CTkLabel(
            tema_frame, text=t("theme_desc"),
            text_color="gray", font=("", 11),
        ).pack(anchor="w", padx=15)

        self._theme_var = ctk.StringVar(value=ctk.get_appearance_mode())

        def _on_theme(v):
            ctk.set_appearance_mode(v)
            _save_settings(theme=v.lower())

        ctk.CTkSegmentedButton(
            tema_frame, values=["Dark", "Light", "System"],
            variable=self._theme_var,
            command=_on_theme,
        ).pack(fill="x", padx=15, pady=(8, 12))

        # ── Yazı Boyutu ──
        font_frame = ctk.CTkFrame(scroll)
        font_frame.pack(fill="x", pady=5)

        ctk.CTkLabel(font_frame, text=t("font_size_title"), font=("", 15, "bold")).pack(
            anchor="w", padx=15, pady=(12, 5)
        )
        ctk.CTkLabel(
            font_frame, text=t("font_size_desc"),
            text_color="gray", font=("", 11),
        ).pack(anchor="w", padx=15)

        size_row = ctk.CTkFrame(font_frame, fg_color="transparent")
        size_row.pack(fill="x", padx=15, pady=(8, 12))

        self._font_size_label = ctk.CTkLabel(
            size_row, text=str(self._font_size), font=("", 16, "bold"), width=35
        )

        def _save_font():
            _save_settings(font_size=self._font_size)

        def decrease():
            self._font_size = max(self._font_size - 1, 9)
            self._font_size_label.configure(text=str(self._font_size))
            self._font_slider.set(self._font_size)
            self._apply_font_size()
            _save_font()

        def increase():
            self._font_size = min(self._font_size + 1, 28)
            self._font_size_label.configure(text=str(self._font_size))
            self._font_slider.set(self._font_size)
            self._apply_font_size()
            _save_font()

        def on_slider(value):
            self._font_size = int(value)
            self._font_size_label.configure(text=str(self._font_size))
            self._apply_font_size()
            _save_font()

        ctk.CTkButton(
            size_row, text="−", width=35, height=32, font=("", 16, "bold"),
            command=decrease,
        ).pack(side="left")

        self._font_slider = ctk.CTkSlider(
            size_row, from_=9, to=28, number_of_steps=19,
            command=on_slider,
        )
        self._font_slider.set(self._font_size)
        self._font_slider.pack(side="left", fill="x", expand=True, padx=8)

        ctk.CTkButton(
            size_row, text="+", width=35, height=32, font=("", 16, "bold"),
            command=increase,
        ).pack(side="left")

        self._font_size_label.pack(side="left", padx=(8, 0))

        # ── Varsayılan Karşılaştırma ──
        comp_frame = ctk.CTkFrame(scroll)
        comp_frame.pack(fill="x", pady=5)

        ctk.CTkLabel(comp_frame, text=t("set_defaults_title"), font=("", 15, "bold")).pack(
            anchor="w", padx=15, pady=(12, 5)
        )
        ctk.CTkLabel(
            comp_frame, text=t("set_defaults_desc"),
            text_color="gray", font=("", 11),
        ).pack(anchor="w", padx=15)

        def _on_default_case(v=None):
            _save_settings(default_ignore_case=self._def_case_var.get())

        def _on_default_strip(v=None):
            _save_settings(default_strip_spaces=self._def_strip_var.get())

        switches_row = ctk.CTkFrame(comp_frame, fg_color="transparent")
        switches_row.pack(fill="x", padx=15, pady=(8, 12))

        self._def_case_var = ctk.BooleanVar(value=prefs.get("default_ignore_case", False))
        ctk.CTkSwitch(
            switches_row, text=t("ignore_case"),
            variable=self._def_case_var, command=_on_default_case,
            font=("", 12),
        ).pack(side="left", padx=(0, 20))

        self._def_strip_var = ctk.BooleanVar(value=prefs.get("default_strip_spaces", False))
        ctk.CTkSwitch(
            switches_row, text=t("strip_spaces"),
            variable=self._def_strip_var, command=_on_default_strip,
            font=("", 12),
        ).pack(side="left")

        # ── Dosya & Dışa Aktarma ──
        export_frame = ctk.CTkFrame(scroll)
        export_frame.pack(fill="x", pady=5)

        ctk.CTkLabel(export_frame, text=t("set_export_title"), font=("", 15, "bold")).pack(
            anchor="w", padx=15, pady=(12, 5)
        )

        dir_row = ctk.CTkFrame(export_frame, fg_color="transparent")
        dir_row.pack(fill="x", padx=15, pady=(5, 5))
        dir_row.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(dir_row, text=t("set_export_dir"), font=("", 12)).grid(
            row=0, column=0, sticky="w", padx=(0, 8)
        )
        self._export_dir_label = ctk.CTkLabel(
            dir_row, text=prefs.get("default_export_dir", "") or t("set_not_set"),
            font=("", 11), text_color="gray", anchor="w",
        )
        self._export_dir_label.grid(row=0, column=1, sticky="ew")

        def _pick_export_dir():
            d = filedialog.askdirectory()
            if d:
                _save_settings(default_export_dir=d)
                self._export_dir_label.configure(text=d)

        def _clear_export_dir():
            _save_settings(default_export_dir="")
            self._export_dir_label.configure(text=t("set_not_set"))

        dir_btn_row = ctk.CTkFrame(dir_row, fg_color="transparent")
        dir_btn_row.grid(row=0, column=2, padx=(8, 0))
        ctk.CTkButton(dir_btn_row, text=t("set_choose"), width=60, height=26,
                       font=("", 11), command=_pick_export_dir).pack(side="left", padx=2)
        ctk.CTkButton(dir_btn_row, text=t("clear"), width=50, height=26,
                       font=("", 11), fg_color="#c0392b", hover_color="#e74c3c",
                       command=_clear_export_dir).pack(side="left", padx=2)

        preview_row = ctk.CTkFrame(export_frame, fg_color="transparent")
        preview_row.pack(fill="x", padx=15, pady=(5, 5))

        ctk.CTkLabel(preview_row, text=t("set_preview_limit"), font=("", 12)).pack(
            side="left", padx=(0, 8)
        )
        self._preview_limit_var = ctk.StringVar(
            value=str(prefs.get("preview_row_limit", 200))
        )
        ctk.CTkOptionMenu(
            preview_row, variable=self._preview_limit_var,
            values=["100", "200", "500", "1000"],
            command=lambda v: _save_settings(preview_row_limit=int(v)),
            width=90,
        ).pack(side="left")

        enc_row = ctk.CTkFrame(export_frame, fg_color="transparent")
        enc_row.pack(fill="x", padx=15, pady=(5, 12))

        ctk.CTkLabel(enc_row, text=t("set_csv_encoding"), font=("", 12)).pack(
            side="left", padx=(0, 8)
        )
        self._csv_enc_var = ctk.StringVar(
            value=prefs.get("csv_encoding", "utf-8-sig")
        )
        ctk.CTkOptionMenu(
            enc_row, variable=self._csv_enc_var,
            values=["utf-8-sig", "utf-8", "windows-1254", "iso-8859-9"],
            command=lambda v: _save_settings(csv_encoding=v),
            width=140,
        ).pack(side="left")

        # ── Geçmiş ──
        history_frame = ctk.CTkFrame(scroll)
        history_frame.pack(fill="x", pady=5)

        ctk.CTkLabel(history_frame, text=t("set_history_title"), font=("", 15, "bold")).pack(
            anchor="w", padx=15, pady=(12, 5)
        )

        hist_row = ctk.CTkFrame(history_frame, fg_color="transparent")
        hist_row.pack(fill="x", padx=15, pady=(5, 5))

        ctk.CTkLabel(hist_row, text=t("set_max_recent"), font=("", 12)).pack(
            side="left", padx=(0, 8)
        )
        self._max_recent_var = ctk.StringVar(
            value=str(prefs.get("max_recent", 20))
        )
        ctk.CTkOptionMenu(
            hist_row, variable=self._max_recent_var,
            values=["10", "20", "50"],
            command=lambda v: _save_settings(max_recent=int(v)),
            width=80,
        ).pack(side="left")

        self._recent_count_label = ctk.CTkLabel(
            hist_row, text="", font=("", 11), text_color="gray"
        )
        self._recent_count_label.pack(side="left", padx=(15, 0))
        recent_count = len(_get_recent_files())
        self._recent_count_label.configure(
            text=t("set_recent_count", count=recent_count)
        )

        def _clear_history():
            _save_settings(recent_files=[])
            self._recent_count_label.configure(text=t("set_recent_count", count=0))

        ctk.CTkButton(
            hist_row, text=t("set_clear_history"), width=120, height=28,
            font=("", 11), fg_color="#c0392b", hover_color="#e74c3c",
            command=_clear_history,
        ).pack(side="right")

        # ── Genel ──
        general_frame = ctk.CTkFrame(scroll)
        general_frame.pack(fill="x", pady=(5, 5))

        ctk.CTkLabel(general_frame, text=t("set_general_title"), font=("", 15, "bold")).pack(
            anchor="w", padx=15, pady=(12, 5)
        )
        ctk.CTkLabel(
            general_frame, text=t("set_window_geo_desc"),
            text_color="gray", font=("", 11),
        ).pack(anchor="w", padx=15)

        self._save_geo_var = ctk.BooleanVar(
            value=bool(prefs.get("window_geometry", ""))
        )

        def _on_geo_toggle():
            if not self._save_geo_var.get():
                _save_settings(window_geometry="")

        ctk.CTkSwitch(
            general_frame, text=t("set_remember_window"),
            variable=self._save_geo_var, command=_on_geo_toggle,
            font=("", 12),
        ).pack(anchor="w", padx=15, pady=(8, 8))

        def _reset_all():
            if messagebox.askyesno(t("warn_title"), t("set_reset_confirm")):
                try:
                    SETTINGS_FILE.unlink(missing_ok=True)
                except Exception:
                    pass
                set_lang("tr")
                ctk.set_appearance_mode("dark")
                self._font_size = 13
                self.geometry("1100x750")
                self._rebuild_all_ui()

        ctk.CTkButton(
            general_frame, text=t("set_reset_all"), width=180, height=32,
            font=("", 12), fg_color="#c0392b", hover_color="#e74c3c",
            command=_reset_all,
        ).pack(anchor="w", padx=15, pady=(5, 12))

        # ── Hakkında ──
        about_frame = ctk.CTkFrame(scroll)
        about_frame.pack(fill="x", pady=(5, 15))

        ctk.CTkLabel(about_frame, text=t("about_title"), font=("", 15, "bold")).pack(
            anchor="w", padx=15, pady=(12, 8)
        )

        info_lines = [
            (t("about_label_app"), t("about_app_name")),
            (t("about_label_version"), "1.0.0"),
            (t("about_label_developer"), "Recep Aydın & Hülya Tokay"),
            (t("about_label_license"), "MIT"),
            (t("about_label_platform"), "Windows / macOS"),
            (t("about_label_tech"), "Python · Pandas · CustomTkinter"),
        ]

        for label, value in info_lines:
            row = ctk.CTkFrame(about_frame, fg_color="transparent")
            row.pack(fill="x", padx=15, pady=1)
            ctk.CTkLabel(row, text=f"{label}:", font=("", 12, "bold"), width=90,
                         anchor="w").pack(side="left")
            ctk.CTkLabel(row, text=value, font=("", 12),
                         text_color="gray").pack(side="left")

        ctk.CTkLabel(
            about_frame, text=t("about_copyright"),
            font=("", 10), text_color="gray",
        ).pack(anchor="w", padx=15, pady=(8, 12))

    def _on_lang_change(self, label):
        code = None
        for k, v in LANGUAGES.items():
            if v == label:
                code = k
                break
        if code and code != get_lang():
            set_lang(code)
            _save_settings(lang=code)
            self._rebuild_all_ui()

    def _rebuild_all_ui(self):
        self.title(t("app_title"))

        self._header_title.configure(text=t("header_title"))
        self._header_subtitle.configure(text=t("header_subtitle"))

        if self._showing_settings:
            self._settings_btn.configure(text=t("settings_btn_back"))
        else:
            self._settings_btn.configure(text=t("settings_btn"))

        if self._showing_history:
            self._history_btn.configure(text=t("settings_btn_back"))
        else:
            self._history_btn.configure(text=t("history_btn"))

        for w in self._settings_frame.winfo_children():
            w.destroy()
        self._build_settings_page()

        for w in self._main_frame.winfo_children():
            w.destroy()
        self._build_file_section()
        self._build_results_section()
        self._build_bottom_bar()
        self._build_status_bar()

        self._refresh_ref_list()
        self._refresh_comp_list()
        if self.ref_files:
            self._rebuild_col_checks("ref")
        if self.comp_files:
            self._rebuild_col_checks("comp")

        if self.results:
            self._display_results()
            self.export_btn.configure(state="normal")
            self._report_btn.configure(state="normal")

        self._update_status_bar()

    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, padx=15, pady=(15, 0), sticky="ew")

        self._header_title = ctk.CTkLabel(
            header, text=t("header_title"), font=("", 24, "bold")
        )
        self._header_title.pack(side="left")

        self._header_subtitle = ctk.CTkLabel(
            header, text=t("header_subtitle"),
            font=("", 14), text_color="gray",
        )
        self._header_subtitle.pack(side="left", padx=10)

        self._settings_btn = ctk.CTkButton(
            header, text=t("settings_btn"), width=90, height=32,
            font=("", 13), command=self._toggle_settings,
            fg_color=("gray75", "gray25"), hover_color=("gray65", "gray35"),
            text_color=("gray10", "gray90"),
        )
        self._settings_btn.pack(side="right")

        self._history_btn = ctk.CTkButton(
            header, text=t("history_btn"), width=90, height=32,
            font=("", 13), command=self._toggle_history,
            fg_color=("gray75", "gray25"), hover_color=("gray65", "gray35"),
            text_color=("gray10", "gray90"),
        )
        self._history_btn.pack(side="right", padx=(0, 8))

    def _build_file_section(self):
        container = ctk.CTkFrame(self._main_frame)
        container.grid(row=0, column=0, padx=15, pady=10, sticky="ew")
        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=1)

        self._build_ref_panel(container)
        self._build_comp_panel(container)

        btn_row = ctk.CTkFrame(container, fg_color="transparent")
        btn_row.grid(row=1, column=0, columnspan=2, padx=5, pady=10, sticky="ew")
        btn_row.grid_columnconfigure(0, weight=1)

        left_group = ctk.CTkFrame(btn_row, fg_color="transparent")
        left_group.grid(row=0, column=0, sticky="w")

        ctk.CTkButton(
            left_group, text=t("swap_sides"), font=("", 13),
            height=35, width=160, command=self._swap_sides,
            fg_color="#8e44ad", hover_color="#9b59b6",
        ).pack(side="left")

        ctk.CTkButton(
            left_group, text=t("profile_save_btn"), font=("", 12),
            height=35, width=110, command=self._save_profile,
            fg_color=("gray75", "gray25"), hover_color=("gray65", "gray35"),
            text_color=("gray10", "gray90"),
        ).pack(side="left", padx=(10, 2))

        ctk.CTkButton(
            left_group, text=t("profile_load_btn"), font=("", 12),
            height=35, width=110, command=self._show_profiles,
            fg_color=("gray75", "gray25"), hover_color=("gray65", "gray35"),
            text_color=("gray10", "gray90"),
        ).pack(side="left", padx=2)

        right_group = ctk.CTkFrame(btn_row, fg_color="transparent")
        right_group.grid(row=0, column=0, sticky="e")

        prefs = _load_settings()
        self._ignore_case_var = ctk.BooleanVar(value=prefs.get("default_ignore_case", False))
        ctk.CTkSwitch(
            right_group, text=t("ignore_case"),
            variable=self._ignore_case_var,
            font=("", 12),
        ).pack(side="left", padx=(0, 10))

        self._strip_spaces_var = ctk.BooleanVar(value=prefs.get("default_strip_spaces", False))
        ctk.CTkSwitch(
            right_group, text=t("strip_spaces"),
            variable=self._strip_spaces_var,
            font=("", 12),
        ).pack(side="left", padx=(0, 15))

        ctk.CTkButton(
            right_group, text=t("compare_btn"), font=("", 16, "bold"),
            height=45, command=self._run_comparison,
            fg_color="#27ae60", hover_color="#2ecc71",
        ).pack(side="left")

        ctk.CTkButton(
            right_group, text=t("detailed_compare_btn"), font=("", 13, "bold"),
            height=45, command=self._run_detailed_comparison,
            fg_color="#2980b9", hover_color="#3498db",
        ).pack(side="left", padx=(8, 0))

    def _build_ref_panel(self, parent):
        panel = ctk.CTkFrame(parent)
        panel.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        ctk.CTkLabel(panel, text=t("ref_panel_title"), font=("", 16, "bold")).pack(
            pady=(10, 5)
        )
        ctk.CTkLabel(
            panel, text=t("ref_panel_desc"), text_color="gray", font=("", 11)
        ).pack()

        self.ref_list_frame = ctk.CTkScrollableFrame(panel, height=60)
        self.ref_list_frame.pack(fill="x", padx=15, pady=5)
        self.ref_list_frame.grid_columnconfigure(0, weight=1)

        drop_hint = t("no_files_dnd") if DND_AVAILABLE else t("no_files")
        self.ref_empty_label = ctk.CTkLabel(
            self.ref_list_frame, text=drop_hint, text_color="#888"
        )
        self.ref_empty_label.grid(row=0, column=0)

        self._enable_drop(panel, self._on_ref_drop)

        btn_row = ctk.CTkFrame(panel, fg_color="transparent")
        btn_row.pack(pady=5)
        ctk.CTkButton(
            btn_row, text=t("add_file"), command=self._add_ref_files, width=100
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            btn_row, text=t("add_folder"), command=self._add_ref_folder, width=100,
            fg_color="#2980b9", hover_color="#3498db",
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            btn_row, text=t("recent_btn"), command=lambda: self._show_recent("ref"),
            width=100, fg_color="#e67e22", hover_color="#f39c12",
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            btn_row, text=t("clear"), command=self._clear_ref_files, width=80,
            fg_color="#c0392b", hover_color="#e74c3c",
        ).pack(side="left", padx=2)

        ctk.CTkLabel(panel, text=t("columns"), font=("", 12), anchor="w").pack(
            fill="x", padx=15
        )
        self.ref_col_frame = ctk.CTkScrollableFrame(panel, height=50)
        self.ref_col_frame.pack(fill="x", padx=15, pady=(0, 10))
        self.ref_col_checks = {}
        ctk.CTkLabel(
            self.ref_col_frame, text=t("add_file_first"), text_color="#888"
        ).pack()

    def _build_comp_panel(self, parent):
        panel = ctk.CTkFrame(parent)
        panel.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

        ctk.CTkLabel(
            panel, text=t("comp_panel_title"), font=("", 16, "bold")
        ).pack(pady=(10, 5))
        ctk.CTkLabel(
            panel, text=t("comp_panel_desc"),
            text_color="gray", font=("", 11),
        ).pack()

        self.comp_list_frame = ctk.CTkScrollableFrame(panel, height=60)
        self.comp_list_frame.pack(fill="x", padx=15, pady=5)
        self.comp_list_frame.grid_columnconfigure(0, weight=1)

        drop_hint = t("no_files_dnd") if DND_AVAILABLE else t("no_files")
        self.comp_empty_label = ctk.CTkLabel(
            self.comp_list_frame, text=drop_hint, text_color="#888"
        )
        self.comp_empty_label.grid(row=0, column=0)

        self._enable_drop(panel, self._on_comp_drop)

        btn_row = ctk.CTkFrame(panel, fg_color="transparent")
        btn_row.pack(pady=5)
        ctk.CTkButton(
            btn_row, text=t("add_file"), command=self._add_comp_files, width=100
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            btn_row, text=t("add_folder"), command=self._add_comp_folder, width=100,
            fg_color="#2980b9", hover_color="#3498db",
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            btn_row, text=t("recent_btn"), command=lambda: self._show_recent("comp"),
            width=100, fg_color="#e67e22", hover_color="#f39c12",
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            btn_row, text=t("clear"), command=self._clear_comp_files, width=80,
            fg_color="#c0392b", hover_color="#e74c3c",
        ).pack(side="left", padx=2)

        ctk.CTkLabel(panel, text=t("columns"), font=("", 12), anchor="w").pack(
            fill="x", padx=15
        )
        self.comp_col_frame = ctk.CTkScrollableFrame(panel, height=50)
        self.comp_col_frame.pack(fill="x", padx=15, pady=(0, 10))
        self.comp_col_checks = {}
        ctk.CTkLabel(
            self.comp_col_frame, text=t("add_file_first"), text_color="#888"
        ).pack()

    def _build_results_section(self):
        results_frame = ctk.CTkFrame(self._main_frame)
        results_frame.grid(row=1, column=0, padx=15, pady=(0, 5), sticky="nsew")
        results_frame.grid_rowconfigure(1, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)

        search_bar = ctk.CTkFrame(results_frame, fg_color="transparent")
        search_bar.grid(row=0, column=0, padx=10, pady=(8, 0), sticky="ew")
        search_bar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(search_bar, text=t("search_label"), font=("", 13)).grid(
            row=0, column=0, padx=(0, 5)
        )
        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", self._on_search_changed)
        self.search_entry = ctk.CTkEntry(
            search_bar, textvariable=self.search_var,
            placeholder_text=t("search_placeholder"), height=30,
        )
        self.search_entry.grid(row=0, column=1, sticky="ew")

        self.search_count_label = ctk.CTkLabel(
            search_bar, text="", font=("", 11), text_color="gray"
        )
        self.search_count_label.grid(row=0, column=2, padx=(8, 0))

        self._copy_btn = ctk.CTkButton(
            search_bar, text=t("copy_btn"), width=90, height=28,
            font=("", 12), command=self._copy_active_tab,
            fg_color=("gray75", "gray25"), hover_color=("gray65", "gray35"),
            text_color=("gray10", "gray90"),
        )
        self._copy_btn.grid(row=0, column=3, padx=(8, 0))

        self.tabview = ctk.CTkTabview(results_frame)
        self.tabview.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="nsew")

        self._tab_full_texts = {}

        tabs = {
            t("tab_matches"): "txt_match",
            t("tab_no_matches"): "txt_no_match",
            t("tab_duplicates"): "txt_duplicates",
            t("tab_only_ref"): "txt_ref",
            t("tab_only_comp"): "txt_comp",
            t("tab_stats"): "txt_stats",
        }
        for tab_name, attr_name in tabs.items():
            tab = self.tabview.add(tab_name)
            textbox = ctk.CTkTextbox(tab, state="disabled")
            textbox.pack(fill="both", expand=True, padx=5, pady=5)
            setattr(self, attr_name, textbox)

        chart_tab = self.tabview.add(t("tab_charts"))

        chart_btn_bar = ctk.CTkFrame(chart_tab, fg_color="transparent")
        chart_btn_bar.pack(fill="x", padx=5, pady=(5, 0))

        self._chart_mode = ctk.StringVar(value="pie_bar")
        self._chart_seg = ctk.CTkSegmentedButton(
            chart_btn_bar,
            values=[t("chart_mode_pie_bar"), t("chart_mode_venn")],
            command=self._on_chart_mode_change,
            font=("", 12),
        )
        self._chart_seg.set(t("chart_mode_pie_bar"))
        self._chart_seg.pack()

        self._chart_canvas = tk.Canvas(chart_tab, highlightthickness=0)
        self._chart_canvas.pack(fill="both", expand=True, padx=5, pady=(3, 5))
        self._chart_canvas.bind("<Configure>", self._on_chart_resize)

    def _build_bottom_bar(self):
        bar = ctk.CTkFrame(self._main_frame, fg_color="transparent")
        bar.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="ew")
        bar.grid_columnconfigure(1, weight=1)

        self.status_label = ctk.CTkLabel(bar, text=t("status_ready"), text_color="gray")
        self.status_label.grid(row=0, column=0, sticky="w")

        self.progress_bar = ctk.CTkProgressBar(bar, height=12)
        self.progress_bar.grid(row=0, column=1, padx=10, sticky="ew")
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(bar, text="", text_color="gray", font=("", 11))
        self.progress_label.grid(row=0, column=2, padx=(0, 10))

        self.export_btn = ctk.CTkButton(
            bar, text=t("export_btn"),
            command=self._export_results, state="disabled",
        )
        self.export_btn.grid(row=0, column=3, padx=(0, 5), sticky="e")

        self._report_btn = ctk.CTkButton(
            bar, text=t("report_btn"), width=90,
            command=self._export_html_report, state="disabled",
            fg_color=("gray75", "gray25"), hover_color=("gray65", "gray35"),
            text_color=("gray10", "gray90"),
        )
        self._report_btn.grid(row=0, column=4, sticky="e")

    def _build_status_bar(self):
        sbar = ctk.CTkFrame(self._main_frame, height=28,
                             fg_color=("gray85", "gray20"), corner_radius=0)
        sbar.grid(row=3, column=0, sticky="ew")
        sbar.grid_columnconfigure(1, weight=1)

        self._sb_files = ctk.CTkLabel(
            sbar, text="", font=("", 11), text_color=("gray30", "gray70"),
        )
        self._sb_files.grid(row=0, column=0, padx=(12, 20))

        self._sb_comp = ctk.CTkLabel(
            sbar, text="", font=("", 11), text_color=("gray30", "gray70"),
        )
        self._sb_comp.grid(row=0, column=1, sticky="w")

        self._sb_time = ctk.CTkLabel(
            sbar, text="", font=("", 11), text_color=("gray30", "gray70"),
        )
        self._sb_time.grid(row=0, column=2, padx=(20, 12))

        self._update_status_bar()

    def _update_status_bar(self):
        if not hasattr(self, "_sb_files"):
            return
        ref_count = len(self.ref_files)
        comp_count = len(self.comp_files)
        self._sb_files.configure(
            text=t("sb_files", ref=ref_count, comp=comp_count)
        )

        if self._last_comp_time and self.results:
            s = self.results["stats"]
            self._sb_comp.configure(
                text=t("sb_result",
                       match=s["match_count"],
                       total=self._last_comp_records)
            )
            self._sb_time.configure(
                text=t("sb_time", time=self._last_comp_time)
            )
        else:
            self._sb_comp.configure(text=t("sb_no_comparison"))
            self._sb_time.configure(text="")

    # ── Dosya / Klasör Seçimi ────────────────────────────────────────

    def _scan_folder(self, folder_path):
        folder = Path(folder_path)
        files = []
        for ext in SUPPORTED_EXTENSIONS:
            files.extend(folder.rglob(f"*{ext}"))
        return sorted(str(f) for f in files if f.is_file())

    def _add_ref_folder(self):
        folder = filedialog.askdirectory(title=t("select_ref_folder"))
        if not folder:
            return
        files = self._scan_folder(folder)
        if not files:
            messagebox.showinfo(t("info_title"), t("no_supported_files"))
            return
        self._add_files_to_ref(files)

    def _add_comp_folder(self):
        folder = filedialog.askdirectory(title=t("select_comp_folder"))
        if not folder:
            return
        files = self._scan_folder(folder)
        if not files:
            messagebox.showinfo(t("info_title"), t("no_supported_files"))
            return
        self._add_files_to_comp(files)

    def _file_types(self):
        return [
            (t("excel_files"), "*.xlsx *.xls"),
            (t("csv_files"), "*.csv"),
            (t("all_files"), "*.*"),
        ]

    def _on_ref_drop(self, event):
        paths = self._parse_drop_data(event.data)
        if paths:
            self._add_files_to_ref(paths)

    def _preview_file(self, file_path):
        PreviewWindow(self, file_path, self.engine)

    def _on_comp_drop(self, event):
        paths = self._parse_drop_data(event.data)
        if paths:
            self._add_files_to_comp(paths)

    def _add_ref_files(self):
        paths = filedialog.askopenfilenames(
            title=t("select_ref_files"), filetypes=self._file_types()
        )
        if not paths:
            return
        self._add_files_to_ref(paths)

    def _add_files_to_ref(self, paths):
        existing_paths = {p for p, _ in self.ref_files}
        first_new = len(self.ref_files) == 0
        added = []

        for p in paths:
            if p not in existing_paths:
                try:
                    sheets = self.engine.get_sheet_names(p)
                    sheet = sheets[0] if sheets else None
                    self.ref_files.append((p, sheet))
                    added.append(p)
                except Exception as e:
                    messagebox.showerror(
                        t("error_file_read_title"),
                        t("error_file_read", name=Path(p).name, e=e),
                    )

        if added:
            _add_to_recent(added)

        self._refresh_ref_list()

        if first_new and self.ref_files:
            self._rebuild_col_checks("ref")

    def _remove_ref_file(self, entry_widget):
        self.ref_files = [
            (p, s) for p, s in self.ref_files if p != entry_widget.file_path
        ]
        self._refresh_ref_list()
        if not self.ref_files:
            self._rebuild_col_checks("ref")

    def _clear_ref_files(self):
        self.ref_files = []
        self._refresh_ref_list()
        self._rebuild_col_checks("ref")

    def _refresh_ref_list(self):
        for widget in self.ref_list_frame.winfo_children():
            widget.destroy()

        if not self.ref_files:
            self.ref_empty_label = ctk.CTkLabel(
                self.ref_list_frame, text=t("no_files"), text_color="#888"
            )
            self.ref_empty_label.grid(row=0, column=0)
            return

        for i, (path, _) in enumerate(self.ref_files):
            entry = FileEntry(self.ref_list_frame, path, self._remove_ref_file, self._preview_file)
            entry.grid(row=i, column=0, sticky="ew", pady=1)

        self._update_status_bar()

    def _add_comp_files(self):
        paths = filedialog.askopenfilenames(
            title=t("select_comp_files"), filetypes=self._file_types()
        )
        if not paths:
            return
        self._add_files_to_comp(paths)

    def _add_files_to_comp(self, paths):
        existing_paths = {p for p, _ in self.comp_files}
        first_new = len(self.comp_files) == 0
        added = []

        for p in paths:
            if p not in existing_paths:
                try:
                    sheets = self.engine.get_sheet_names(p)
                    sheet = sheets[0] if sheets else None
                    self.comp_files.append((p, sheet))
                    added.append(p)
                except Exception as e:
                    messagebox.showerror(
                        t("error_file_read_title"),
                        t("error_file_read", name=Path(p).name, e=e),
                    )

        if added:
            _add_to_recent(added)

        self._refresh_comp_list()

        if first_new and self.comp_files:
            self._rebuild_col_checks("comp")

    def _remove_comp_file(self, entry_widget):
        self.comp_files = [
            (p, s) for p, s in self.comp_files if p != entry_widget.file_path
        ]
        self._refresh_comp_list()
        if not self.comp_files:
            self._rebuild_col_checks("comp")

    def _clear_comp_files(self):
        self.comp_files = []
        self._refresh_comp_list()
        self._rebuild_col_checks("comp")

    def _refresh_comp_list(self):
        for widget in self.comp_list_frame.winfo_children():
            widget.destroy()

        if not self.comp_files:
            self.comp_empty_label = ctk.CTkLabel(
                self.comp_list_frame, text=t("no_files"), text_color="#888"
            )
            self.comp_empty_label.grid(row=0, column=0)
            return

        for i, (path, _) in enumerate(self.comp_files):
            entry = FileEntry(self.comp_list_frame, path, self._remove_comp_file, self._preview_file)
            entry.grid(row=i, column=0, sticky="ew", pady=1)

        self._update_status_bar()

    # ── Sütun Checkbox Yönetimi ────────────────────────────────────

    def _rebuild_col_checks(self, side):
        if side == "ref":
            frame = self.ref_col_frame
            files = self.ref_files
            checks_attr = "ref_col_checks"
        else:
            frame = self.comp_col_frame
            files = self.comp_files
            checks_attr = "comp_col_checks"

        for widget in frame.winfo_children():
            widget.destroy()

        checks = {}
        setattr(self, checks_attr, checks)

        if not files:
            ctk.CTkLabel(frame, text=t("add_file_first"), text_color="#888").pack()
            return

        all_columns = []
        seen = set()
        for path, sheet in files:
            try:
                cols = self.engine.get_columns(path, sheet)
                for c in cols:
                    if c not in seen:
                        all_columns.append(c)
                        seen.add(c)
            except Exception:
                pass

        if not all_columns:
            ctk.CTkLabel(frame, text=t("no_columns"), text_color="#888").pack()
            return

        for col in all_columns:
            var = ctk.BooleanVar(value=False)
            cb = ctk.CTkCheckBox(frame, text=col, variable=var)
            cb.pack(anchor="w", padx=5, pady=1)
            checks[col] = var

    def _show_recent(self, side):
        recent = _get_recent_files()
        if not recent:
            messagebox.showinfo(t("info_title"), t("no_recent"))
            return

        win = ctk.CTkToplevel(self)
        win.title(t("recent_title"))
        win.geometry("500x400")
        win.resizable(False, True)
        win.transient(self)
        win.grab_set()

        ctk.CTkLabel(
            win, text=t("recent_title"), font=("", 16, "bold")
        ).pack(pady=(12, 5))
        ctk.CTkLabel(
            win, text=t("recent_hint"), font=("", 11), text_color="gray"
        ).pack(pady=(0, 8))

        scroll = ctk.CTkScrollableFrame(win)
        scroll.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        scroll.grid_columnconfigure(0, weight=1)

        selected = {}
        for i, path in enumerate(recent):
            var = ctk.BooleanVar(value=False)
            selected[path] = var
            row = ctk.CTkFrame(scroll, fg_color="transparent")
            row.grid(row=i, column=0, sticky="ew", pady=1)
            row.grid_columnconfigure(1, weight=1)
            ctk.CTkCheckBox(row, text="", variable=var, width=24).grid(
                row=0, column=0, padx=(0, 4)
            )
            ctk.CTkLabel(
                row, text=Path(path).name, font=("", 12), anchor="w"
            ).grid(row=0, column=1, sticky="ew")
            ctk.CTkLabel(
                row, text=str(Path(path).parent), font=("", 10),
                text_color="gray", anchor="w",
            ).grid(row=1, column=1, sticky="ew")

        def _add_selected():
            chosen = [p for p, v in selected.items() if v.get()]
            if not chosen:
                return
            win.destroy()
            if side == "ref":
                self._add_files_to_ref(chosen)
            else:
                self._add_files_to_comp(chosen)

        btn_frame = ctk.CTkFrame(win, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(
            btn_frame, text=t("recent_add_selected"), command=_add_selected,
            fg_color="#27ae60", hover_color="#2ecc71",
        ).pack(side="right", padx=2)
        ctk.CTkButton(
            btn_frame, text=t("recent_cancel"), command=win.destroy,
            fg_color=("gray75", "gray25"), hover_color=("gray65", "gray35"),
            text_color=("gray10", "gray90"),
        ).pack(side="right", padx=2)

    def _swap_sides(self):
        self.ref_files, self.comp_files = self.comp_files, self.ref_files
        self._refresh_ref_list()
        self._refresh_comp_list()
        self._rebuild_col_checks("ref")
        self._rebuild_col_checks("comp")

    def _save_profile(self):
        ref_cols = self._get_selected_columns("ref")
        comp_cols = self._get_selected_columns("comp")

        dialog = ctk.CTkInputDialog(
            text=t("profile_name_prompt"), title=t("profile_save_title"),
        )
        name = dialog.get_input()
        if not name or not name.strip():
            return
        name = name.strip()

        profile = {
            "name": name,
            "ref_columns": ref_cols,
            "comp_columns": comp_cols,
            "ignore_case": self._ignore_case_var.get(),
            "strip_spaces": self._strip_spaces_var.get(),
            "created": datetime.now().strftime("%d.%m.%Y %H:%M"),
        }
        _add_profile(profile)
        messagebox.showinfo(t("success_title"), t("profile_saved", name=name))

    def _show_profiles(self):
        profiles = _load_profiles()
        if not profiles:
            messagebox.showinfo(t("info_title"), t("profile_empty"))
            return

        win = ctk.CTkToplevel(self)
        win.title(t("profile_list_title"))
        win.geometry("480x420")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()

        ctk.CTkLabel(
            win, text=t("profile_list_title"), font=("", 18, "bold"),
        ).pack(pady=(12, 8))

        scroll = ctk.CTkScrollableFrame(win)
        scroll.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        scroll.grid_columnconfigure(0, weight=1)

        for i, p in enumerate(profiles):
            card = ctk.CTkFrame(scroll, corner_radius=8)
            card.grid(row=i, column=0, sticky="ew", pady=4)
            card.grid_columnconfigure(0, weight=1)

            info_frame = ctk.CTkFrame(card, fg_color="transparent")
            info_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=8)
            info_frame.grid_columnconfigure(0, weight=1)

            ctk.CTkLabel(
                info_frame, text=f"📌 {p['name']}",
                font=("", 14, "bold"), anchor="w",
            ).grid(row=0, column=0, sticky="w")

            ref_c = ", ".join(p.get("ref_columns", [])) or "—"
            comp_c = ", ".join(p.get("comp_columns", [])) or "—"
            flags = []
            if p.get("ignore_case"):
                flags.append(t("ignore_case"))
            if p.get("strip_spaces"):
                flags.append(t("strip_spaces"))
            flag_str = " | ".join(flags) if flags else "—"

            detail = f"{t('profile_ref_cols')}: {ref_c}\n{t('profile_comp_cols')}: {comp_c}\n{t('profile_flags')}: {flag_str}"
            ctk.CTkLabel(
                info_frame, text=detail,
                font=("", 11), text_color="gray", anchor="w", justify="left",
            ).grid(row=1, column=0, sticky="w", pady=(2, 0))

            date_str = p.get("created", "")
            if date_str:
                ctk.CTkLabel(
                    info_frame, text=date_str,
                    font=("", 10), text_color="gray",
                ).grid(row=0, column=1, sticky="e")

            btn_frame = ctk.CTkFrame(card, fg_color="transparent")
            btn_frame.grid(row=0, column=1, padx=(0, 10), pady=8)

            profile_name = p["name"]
            ctk.CTkButton(
                btn_frame, text=t("profile_apply"), width=70, height=28,
                font=("", 11), fg_color="#27ae60", hover_color="#2ecc71",
                command=lambda pn=profile_name, w=win: self._apply_profile(pn, w),
            ).pack(pady=(0, 3))
            ctk.CTkButton(
                btn_frame, text=t("profile_delete"), width=70, height=28,
                font=("", 11), fg_color="#c0392b", hover_color="#e74c3c",
                command=lambda pn=profile_name, w=win: self._delete_and_refresh_profiles(pn, w),
            ).pack()

    def _apply_profile(self, profile_name, win):
        profiles = _load_profiles()
        profile = next((p for p in profiles if p["name"] == profile_name), None)
        if not profile:
            return
        win.destroy()

        for col, var in self.ref_col_checks.items():
            var.set(col in profile.get("ref_columns", []))
        for col, var in self.comp_col_checks.items():
            var.set(col in profile.get("comp_columns", []))

        self._ignore_case_var.set(profile.get("ignore_case", False))
        self._strip_spaces_var.set(profile.get("strip_spaces", False))

        messagebox.showinfo(
            t("success_title"),
            t("profile_applied", name=profile_name),
        )

    def _delete_and_refresh_profiles(self, profile_name, win):
        _delete_profile(profile_name)
        win.destroy()
        self._show_profiles()

    def _get_selected_columns(self, side):
        checks = self.ref_col_checks if side == "ref" else self.comp_col_checks
        return [col for col, var in checks.items() if var.get()]

    # ── Karşılaştırma ─────────────────────────────────────────────

    def _run_comparison(self):
        if not self.ref_files:
            messagebox.showwarning(t("warn_title"), t("warn_no_ref"))
            return
        if not self.comp_files:
            messagebox.showwarning(t("warn_title"), t("warn_no_comp"))
            return

        ref_cols = self._get_selected_columns("ref")
        comp_cols = self._get_selected_columns("comp")
        if not ref_cols or not comp_cols:
            messagebox.showwarning(t("warn_title"), t("warn_no_cols"))
            return

        self.progress_bar.set(0)
        self.progress_label.configure(text="")
        self.status_label.configure(text=t("comparing"), text_color="yellow")
        self.update()

        def on_progress(checked, total, phase):
            ratio = checked / total if total > 0 else 0
            self.progress_bar.set(ratio)
            pct = int(ratio * 100)
            self.progress_label.configure(
                text=t("rows_checked", checked=f"{checked:,}", total=f"{total:,}", pct=pct)
            )
            self.status_label.configure(text=phase, text_color="yellow")
            self.update_idletasks()

        try:
            self.results = self.engine.compare(
                self.ref_files, ref_cols, self.comp_files, comp_cols,
                on_progress=on_progress,
                ignore_case=self._ignore_case_var.get(),
                strip_spaces=self._strip_spaces_var.get(),
            )
            self.progress_bar.set(1)
            self._display_results()
            self.export_btn.configure(state="normal")
            self._report_btn.configure(state="normal")
            s = self.results["stats"]
            no_match = s["only_ref_count"] + s["only_comp_count"]
            self.progress_label.configure(text="")
            self.status_label.configure(
                text=t("completed", match=s["match_count"], no_match=no_match),
                text_color="#2ecc71",
            )
            self._last_comp_time = datetime.now().strftime("%H:%M")
            self._last_comp_records = s["ref_total"] + s["comp_total"]
            self._update_status_bar()

            _save_history_entry({
                "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
                "ref_files": [Path(p).name for p, _ in self.ref_files],
                "comp_files": [Path(p).name for p, _ in self.comp_files],
                "match_count": s["match_count"],
                "only_ref": s["only_ref_count"],
                "only_comp": s["only_comp_count"],
                "match_pct": s["match_percentage"],
                "ref_total": s["ref_total"],
                "comp_total": s["comp_total"],
            })
        except Exception as e:
            self.progress_bar.set(0)
            self.progress_label.configure(text="")
            messagebox.showerror(t("error_title"), t("error_comparison", e=e))
            self.status_label.configure(text=t("error_occurred"), text_color="red")

    def _run_detailed_comparison(self):
        if not self.ref_files:
            messagebox.showwarning(t("warn_title"), t("warn_no_ref"))
            return
        if not self.comp_files:
            messagebox.showwarning(t("warn_title"), t("warn_no_comp"))
            return

        all_columns = []
        seen = set()
        for path, sheet in self.ref_files:
            try:
                cols = self.engine.get_columns(path, sheet)
                for c in cols:
                    if c not in seen:
                        all_columns.append(c)
                        seen.add(c)
            except Exception:
                pass

        if len(all_columns) < 2:
            messagebox.showwarning(t("warn_title"), t("detail_need_cols"))
            return

        self._show_detail_column_picker(all_columns)

    def _show_detail_column_picker(self, all_columns):
        win = ctk.CTkToplevel(self)
        win.title(t("detail_picker_title"))
        win.geometry("520x550")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()

        ctk.CTkLabel(
            win, text=t("detail_picker_title"), font=("", 18, "bold"),
        ).pack(pady=(15, 5))
        ctk.CTkLabel(
            win, text=t("detail_picker_desc"),
            font=("", 12), text_color="gray", wraplength=480,
        ).pack(padx=15, pady=(0, 10))

        key_frame = ctk.CTkFrame(win)
        key_frame.pack(fill="x", padx=15, pady=(0, 5))
        ctk.CTkLabel(
            key_frame, text=t("detail_key_label"),
            font=("", 13, "bold"), text_color="#e67e22",
        ).pack(anchor="w", padx=10, pady=(8, 2))
        ctk.CTkLabel(
            key_frame, text=t("detail_key_hint"),
            font=("", 11), text_color="gray",
        ).pack(anchor="w", padx=10, pady=(0, 5))

        key_scroll = ctk.CTkScrollableFrame(key_frame, height=100)
        key_scroll.pack(fill="x", padx=10, pady=(0, 8))

        key_vars = {}
        for col in all_columns:
            var = ctk.BooleanVar(value=False)
            ctk.CTkCheckBox(key_scroll, text=col, variable=var, font=("", 12)).pack(
                anchor="w", pady=1
            )
            key_vars[col] = var

        detail_frame = ctk.CTkFrame(win)
        detail_frame.pack(fill="both", expand=True, padx=15, pady=(5, 5))
        ctk.CTkLabel(
            detail_frame, text=t("detail_detail_label"),
            font=("", 13, "bold"), text_color="#2980b9",
        ).pack(anchor="w", padx=10, pady=(8, 2))
        ctk.CTkLabel(
            detail_frame, text=t("detail_detail_hint"),
            font=("", 11), text_color="gray",
        ).pack(anchor="w", padx=10, pady=(0, 5))

        detail_scroll = ctk.CTkScrollableFrame(detail_frame, height=120)
        detail_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        detail_vars = {}
        for col in all_columns:
            var = ctk.BooleanVar(value=False)
            ctk.CTkCheckBox(detail_scroll, text=col, variable=var, font=("", 12)).pack(
                anchor="w", pady=1
            )
            detail_vars[col] = var

        def on_run():
            keys = [c for c, v in key_vars.items() if v.get()]
            details = [c for c, v in detail_vars.items() if v.get()]
            if not keys:
                messagebox.showwarning(t("warn_title"), t("detail_no_key"))
                return
            if not details:
                messagebox.showwarning(t("warn_title"), t("detail_no_detail"))
                return
            overlap = set(keys) & set(details)
            if overlap:
                messagebox.showwarning(
                    t("warn_title"),
                    t("detail_overlap", cols=", ".join(overlap)),
                )
                return
            win.destroy()
            self._execute_detailed_comparison(keys, details)

        ctk.CTkButton(
            win, text=t("detail_run_btn"), font=("", 15, "bold"),
            height=42, command=on_run,
            fg_color="#2980b9", hover_color="#3498db",
        ).pack(pady=(5, 15))

    def _execute_detailed_comparison(self, key_cols, detail_cols):
        self.progress_bar.set(0)
        self.progress_label.configure(text="")
        self.status_label.configure(text=t("comparing"), text_color="yellow")
        self.update()

        def on_progress(checked, total, phase):
            ratio = checked / total if total > 0 else 0
            self.progress_bar.set(ratio)
            pct = int(ratio * 100)
            self.progress_label.configure(
                text=t("rows_checked", checked=f"{checked:,}", total=f"{total:,}", pct=pct)
            )
            self.status_label.configure(text=phase, text_color="yellow")
            self.update_idletasks()

        try:
            detail_results = self.engine.detailed_compare(
                self.ref_files, self.comp_files,
                key_cols, detail_cols,
                on_progress=on_progress,
                ignore_case=self._ignore_case_var.get(),
                strip_spaces=self._strip_spaces_var.get(),
            )
            self.progress_bar.set(1)
            self.progress_label.configure(text="")
            s = detail_results["stats"]
            self.status_label.configure(
                text=t("detail_completed",
                       matched=s["total_matched"],
                       diffs=s["total_with_diffs"]),
                text_color="#2ecc71",
            )
            self._show_detail_results(detail_results, key_cols, detail_cols)
        except Exception as e:
            self.progress_bar.set(0)
            self.progress_label.configure(text="")
            messagebox.showerror(t("error_title"), t("error_comparison", e=e))
            self.status_label.configure(text=t("error_occurred"), text_color="red")

    def _show_detail_results(self, results, key_cols, detail_cols):
        win = ctk.CTkToplevel(self)
        win.title(t("detail_results_title"))
        win.geometry("900x650")
        win.transient(self)

        s = results["stats"]

        header = ctk.CTkFrame(win, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(12, 5))

        ctk.CTkLabel(
            header, text=t("detail_results_title"),
            font=("", 20, "bold"),
        ).pack(side="left")

        stats_text = (
            f"  {t('detail_stat_matched')}: {s['total_matched']:,}"
            f"  |  {t('detail_stat_diffs')}: {s['total_with_diffs']:,}"
            f"  |  {t('detail_stat_identical')}: {s['total_identical']:,}"
        )
        ctk.CTkLabel(
            header, text=stats_text, font=("", 12), text_color="gray",
        ).pack(side="left", padx=(15, 0))

        tabview = ctk.CTkTabview(win)
        tabview.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # ── Farklar sekmesi ──
        diff_tab = tabview.add(t("detail_tab_diffs"))
        diff_text = ctk.CTkTextbox(diff_tab)
        diff_text.pack(fill="both", expand=True, padx=5, pady=5)

        diffs = results["differences"]
        if diffs:
            lines = []
            lines.append(t("detail_diff_header", count=len(diffs)))
            lines.append("=" * 60)
            lines.append("")
            for i, entry in enumerate(diffs, 1):
                key_str = " | ".join(f"{k}: {v}" for k, v in entry["key"].items())
                lines.append(f"{'─' * 50}")
                lines.append(f"  #{i}  🔑 {key_str}")
                lines.append(f"       {t('detail_ref_file')}: {entry['ref_file']}")
                lines.append(f"       {t('detail_comp_file')}: {entry['comp_file']}")
                for d in entry["diffs"]:
                    lines.append(f"       ⚠ {d['column']}:")
                    lines.append(f"           {t('detail_ref_val')}: \"{d['ref_value']}\"")
                    lines.append(f"           {t('detail_comp_val')}: \"{d['comp_value']}\"")
                lines.append("")
            content = "\n".join(lines)
        else:
            content = t("detail_no_diffs")

        diff_text.insert("1.0", content)
        diff_text.configure(state="disabled")

        # ── Özet sekmesi ──
        summary_tab = tabview.add(t("detail_tab_summary"))
        summary_text = ctk.CTkTextbox(summary_tab)
        summary_text.pack(fill="both", expand=True, padx=5, pady=5)

        summary_lines = [
            "╔══════════════════════════════════════════════════╗",
            f"║     {t('detail_summary_title'):^44} ║",
            "╠══════════════════════════════════════════════════╣",
            f"║  {t('detail_key_cols')+':':<34} {', '.join(key_cols):>12} ║",
            f"║  {t('detail_detail_cols')+':':<34} {', '.join(detail_cols):>12} ║",
            "╠══════════════════════════════════════════════════╣",
            f"║  {t('detail_stat_matched')+':':<34} {s['total_matched']:>12,} ║",
            f"║  {t('detail_stat_diffs')+':':<34} {s['total_with_diffs']:>12,} ║",
            f"║  {t('detail_stat_identical')+':':<34} {s['total_identical']:>12,} ║",
            f"║  {t('stats_only_ref')+':':<34} {s['only_ref_count']:>12,} ║",
            f"║  {t('stats_only_comp')+':':<34} {s['only_comp_count']:>12,} ║",
            "╚══════════════════════════════════════════════════╝",
        ]

        if diffs:
            col_diff_counts = {}
            for entry in diffs:
                for d in entry["diffs"]:
                    col_diff_counts[d["column"]] = col_diff_counts.get(d["column"], 0) + 1

            summary_lines.append("")
            summary_lines.append(t("detail_col_breakdown"))
            summary_lines.append("─" * 40)
            for col, cnt in sorted(col_diff_counts.items(), key=lambda x: -x[1]):
                summary_lines.append(f"  {col:<25} {cnt:>6} {t('detail_diff_count')}")

        summary_text.insert("1.0", "\n".join(summary_lines))
        summary_text.configure(state="disabled")

    def _display_results(self):
        r = self.results
        s = r["stats"]

        if r["matches"]:
            lines = [f"  • {v}" for v in r["matches"]]
            match_text = (
                t("matched_records", count=s["match_count"]) + "\n"
                + "=" * 50 + "\n\n"
                + "\n".join(lines)
            )
        else:
            match_text = t("no_matched_records")
        self._fill_textbox(self.txt_match, match_text)

        no_match_lines = []
        if r["only_in_reference"]:
            no_match_lines.append(
                t("ref_not_in_comp", count=s["only_ref_count"]) + "\n"
            )
            for v in r["only_in_reference"]:
                no_match_lines.append(f"  ✗ {v}")
            no_match_lines.append("")

        if r["only_in_comparison"]:
            no_match_lines.append(
                t("comp_not_in_ref", count=s["only_comp_count"]) + "\n"
            )
            for v in r["only_in_comparison"]:
                sources = ", ".join(r["source_map"].get(v, ["?"]))
                no_match_lines.append(f"  ✗ {v}  ← [{sources}]")

        total_no_match = s["only_ref_count"] + s["only_comp_count"]
        if no_match_lines:
            no_match_text = (
                t("unmatched_records", count=total_no_match) + "\n"
                + "=" * 50 + "\n\n"
                + "\n".join(no_match_lines)
            )
        else:
            no_match_text = t("all_matched")
        self._fill_textbox(self.txt_no_match, no_match_text)

        dup_lines = []
        if r["ref_duplicates"]:
            dup_lines.append(
                t("ref_dup_header", count=s["ref_dup_count"]) + "\n"
            )
            for val, details in sorted(r["ref_duplicates"].items()):
                for d in details:
                    dup_lines.append(f"  ⚠ {val}  ← {d}")
            dup_lines.append("")

        if r["comp_duplicates"]:
            dup_lines.append(
                t("comp_dup_header", count=s["comp_dup_count"]) + "\n"
            )
            for val, details in sorted(r["comp_duplicates"].items()):
                for d in details:
                    dup_lines.append(f"  ⚠ {val}  ← {d}")

        total_dups = s["ref_dup_count"] + s["comp_dup_count"]
        if dup_lines:
            dup_text = (
                t("dup_records", count=total_dups) + "\n"
                + "=" * 50 + "\n\n"
                + "\n".join(dup_lines)
            )
        else:
            dup_text = t("no_duplicates")
        self._fill_textbox(self.txt_duplicates, dup_text)

        if r["only_in_reference"]:
            lines = [f"  • {v}" for v in r["only_in_reference"]]
            ref_text = (
                t("only_ref_header", count=s["only_ref_count"]) + "\n"
                + "=" * 50 + "\n\n"
                + "\n".join(lines)
            )
        else:
            ref_text = t("all_records_match")
        self._fill_textbox(self.txt_ref, ref_text)

        if r["only_in_comparison"]:
            lines = []
            for v in r["only_in_comparison"]:
                sources = ", ".join(r["source_map"].get(v, ["?"]))
                lines.append(f"  • {v}  ← [{sources}]")
            comp_text = (
                t("only_comp_header", count=s["only_comp_count"]) + "\n"
                + "=" * 50 + "\n\n"
                + "\n".join(lines)
            )
        else:
            comp_text = t("no_extra_comp")
        self._fill_textbox(self.txt_comp, comp_text)

        stats_text = (
            "╔════════════════════════════════════════════╗\n"
            f"║        {t('stats_title'):^36} ║\n"
            "╠════════════════════════════════════════════╣\n"
            f"║  {t('stats_ref_total')+':':<30} {s['ref_total']:>6}       ║\n"
            f"║  {t('stats_comp_total')+':':<30} {s['comp_total']:>6}       ║\n"
            "╠════════════════════════════════════════════╣\n"
            f"║  {t('stats_matched')+':':<30} {s['match_count']:>6}       ║\n"
            f"║  {t('stats_only_ref')+':':<30} {s['only_ref_count']:>6}       ║\n"
            f"║  {t('stats_only_comp')+':':<30} {s['only_comp_count']:>6}       ║\n"
            "╠════════════════════════════════════════════╣\n"
            f"║  {t('stats_match_pct')+':':<30} %{s['match_percentage']:>5}       ║\n"
            "╚════════════════════════════════════════════╝"
        )
        self._fill_textbox(self.txt_stats, stats_text)
        self.after(100, self._draw_charts)

    def _on_chart_resize(self, event=None):
        if self.results:
            self._draw_charts()

    def _on_chart_mode_change(self, value):
        mode_map = {
            t("chart_mode_pie_bar"): "pie_bar",
            t("chart_mode_venn"): "venn",
        }
        self._chart_mode.set(mode_map.get(value, "pie_bar"))
        if self.results:
            self._draw_charts()

    def _draw_charts(self):
        c = self._chart_canvas
        c.delete("all")
        c.update_idletasks()

        w = c.winfo_width()
        h = c.winfo_height()
        if w < 50 or h < 50:
            return

        s = self.results["stats"]
        is_dark = ctk.get_appearance_mode().lower() == "dark"
        bg = "#2b2b2b" if is_dark else "#f0f0f0"
        fg = "#ffffff" if is_dark else "#1a1a1a"
        fg_dim = "#aaaaaa" if is_dark else "#666666"
        c.configure(bg=bg)

        match_count = s["match_count"]
        only_ref = s["only_ref_count"]
        only_comp = s["only_comp_count"]
        total = match_count + only_ref + only_comp
        if total == 0:
            c.create_text(w // 2, h // 2, text=t("no_matched_records"),
                          fill=fg, font=("", 14))
            return

        colors = {"match": "#27ae60", "ref": "#e74c3c", "comp": "#f39c12"}
        labels = {
            "match": t("stats_matched"),
            "ref": t("stats_only_ref"),
            "comp": t("stats_only_comp"),
        }
        tab_map = {
            "match": t("tab_matches"),
            "ref": t("tab_only_ref"),
            "comp": t("tab_only_comp"),
        }
        values = {"match": match_count, "ref": only_ref, "comp": only_comp}

        def _go_tab(tab_name):
            def handler(event=None):
                self.tabview.set(tab_name)
            return handler

        hover_outline = "#ffffff" if is_dark else "#000000"

        def _hover_in(tag):
            def handler(event=None):
                c.itemconfigure(tag, outline=hover_outline, width=3)
                c.configure(cursor="hand2")
            return handler

        def _hover_out(tag):
            def handler(event=None):
                c.itemconfigure(tag, outline=bg, width=2)
                c.configure(cursor="")
            return handler

        mode = self._chart_mode.get()

        if mode == "venn":
            self._draw_venn(c, w, h, bg, fg, fg_dim, colors, labels, tab_map,
                            values, match_count, only_ref, only_comp, total,
                            _go_tab)
        else:
            self._draw_pie_bar(c, w, h, bg, fg, fg_dim, is_dark, colors, labels,
                               tab_map, values, total, _go_tab, _hover_in,
                               _hover_out)

    def _draw_pie_bar(self, c, w, h, bg, fg, fg_dim, is_dark, colors, labels,
                      tab_map, values, total, _go_tab, _hover_in, _hover_out):

        margin_top = 15
        margin_bot = 40
        usable_h = h - margin_top - margin_bot

        # ── Pasta Grafik (sol) ──
        pie_cx = w * 0.25
        pie_cy = margin_top + usable_h * 0.5
        pie_r = min(w * 0.18, usable_h * 0.38, 120)

        c.create_text(pie_cx, margin_top + 5,
                       text=t("chart_pie_title"), fill=fg, font=("", 14, "bold"))

        start = 90
        for key in ("match", "ref", "comp"):
            val = values[key]
            if val == 0:
                continue
            extent = -(val / total) * 360
            tag = f"pie_{key}"
            c.create_arc(
                pie_cx - pie_r, pie_cy - pie_r,
                pie_cx + pie_r, pie_cy + pie_r,
                start=start, extent=extent,
                fill=colors[key], outline=bg, width=2,
                tags=(tag,),
            )
            c.tag_bind(tag, "<Button-1>", _go_tab(tab_map[key]))
            c.tag_bind(tag, "<Enter>", _hover_in(tag))
            c.tag_bind(tag, "<Leave>", _hover_out(tag))

            mid_angle = math.radians(start + extent / 2)
            lx = pie_cx + (pie_r * 0.6) * math.cos(mid_angle)
            ly = pie_cy - (pie_r * 0.6) * math.sin(mid_angle)
            pct = round(val / total * 100, 1)
            if pct >= 5:
                pct_tag = f"pie_pct_{key}"
                c.create_text(lx, ly, text=f"%{pct}",
                              fill="white", font=("", 11, "bold"),
                              tags=(pct_tag,))
                c.tag_bind(pct_tag, "<Button-1>", _go_tab(tab_map[key]))
                c.tag_bind(pct_tag, "<Enter>", _hover_in(tag))
                c.tag_bind(pct_tag, "<Leave>", _hover_out(tag))
            start += extent

        # ── Bar Grafik (sağ) ──
        bar_left = w * 0.52
        bar_right = w - 20
        bar_w = bar_right - bar_left
        bar_area_top = margin_top + 25
        bar_area_bot = margin_top + usable_h - 10
        bar_area_h = bar_area_bot - bar_area_top

        c.create_text(bar_left + bar_w / 2, margin_top + 5,
                       text=t("chart_bar_title"), fill=fg, font=("", 14, "bold"))

        max_val = max(values.values()) or 1
        bar_h = min(bar_area_h / 3 * 0.55, 35)
        gap = bar_area_h / 3

        for i, key in enumerate(("match", "ref", "comp")):
            val = values[key]
            y_top = bar_area_top + i * gap
            filled_w = (val / max_val) * bar_w if max_val > 0 else 0

            bg_tag = f"bar_bg_{key}"
            c.create_rectangle(
                bar_left, y_top, bar_left + bar_w, y_top + bar_h,
                fill="#3a3a3a" if is_dark else "#dcdcdc", outline="",
                tags=(bg_tag,),
            )
            c.tag_bind(bg_tag, "<Button-1>", _go_tab(tab_map[key]))
            c.tag_bind(bg_tag, "<Enter>", lambda e: c.configure(cursor="hand2"))
            c.tag_bind(bg_tag, "<Leave>", lambda e: c.configure(cursor=""))

            if filled_w > 0:
                bar_tag = f"bar_{key}"
                c.create_rectangle(
                    bar_left, y_top, bar_left + filled_w, y_top + bar_h,
                    fill=colors[key], outline="",
                    tags=(bar_tag,),
                )
                c.tag_bind(bar_tag, "<Button-1>", _go_tab(tab_map[key]))
                c.tag_bind(bar_tag, "<Enter>", lambda e: c.configure(cursor="hand2"))
                c.tag_bind(bar_tag, "<Leave>", lambda e: c.configure(cursor=""))

            lbl_tag = f"bar_lbl_{key}"
            c.create_text(
                bar_left, y_top + bar_h + 14,
                text=f"{labels[key]}: {val:,}", anchor="w",
                fill=fg_dim, font=("", 11),
                tags=(lbl_tag,),
            )
            c.tag_bind(lbl_tag, "<Button-1>", _go_tab(tab_map[key]))
            c.tag_bind(lbl_tag, "<Enter>", lambda e: c.configure(cursor="hand2"))
            c.tag_bind(lbl_tag, "<Leave>", lambda e: c.configure(cursor=""))

        # ── Lejant ──
        legend_y = h - margin_bot + 10
        legend_x_start = w * 0.08
        spacing = w * 0.31

        for i, key in enumerate(("match", "ref", "comp")):
            lx = legend_x_start + i * spacing
            leg_rect_tag = f"leg_rect_{key}"
            c.create_rectangle(lx, legend_y, lx + 14, legend_y + 14,
                                fill=colors[key], outline="",
                                tags=(leg_rect_tag,))
            pct = round(values[key] / total * 100, 1) if total > 0 else 0
            leg_txt_tag = f"leg_txt_{key}"
            c.create_text(lx + 20, legend_y + 7,
                          text=f"{labels[key]} ({values[key]:,} — %{pct})",
                          anchor="w", fill=fg, font=("", 11),
                          tags=(leg_txt_tag,))
            for lt in (leg_rect_tag, leg_txt_tag):
                c.tag_bind(lt, "<Button-1>", _go_tab(tab_map[key]))
                c.tag_bind(lt, "<Enter>", lambda e: c.configure(cursor="hand2"))
                c.tag_bind(lt, "<Leave>", lambda e: c.configure(cursor=""))

    def _draw_venn(self, c, w, h, bg, fg, fg_dim, colors, labels, tab_map,
                   values, match_count, only_ref, only_comp, total, _go_tab):

        venn_cx = w * 0.5
        venn_cy = h * 0.50
        venn_r = min(w * 0.18, h * 0.28, 130)
        offset = venn_r * 0.60

        ref_total_unique = match_count + only_ref
        comp_total_unique = match_count + only_comp

        ref_color = "#e74c3c"
        comp_color = "#3498db"
        intersect_color = "#27ae60"

        c.create_text(w / 2, 18,
                       text=t("chart_venn_title"), fill=fg, font=("", 16, "bold"))

        # Sol daire etiketi
        c.create_text(
            venn_cx - offset - venn_r * 0.3, venn_cy - venn_r - 30,
            text=f"{t('venn_ref_label')} ({ref_total_unique:,})",
            fill=ref_color, font=("", 11, "bold"),
        )
        # Sağ daire etiketi
        c.create_text(
            venn_cx + offset + venn_r * 0.3, venn_cy - venn_r - 30,
            text=f"{t('venn_comp_label')} ({comp_total_unique:,})",
            fill=comp_color, font=("", 11, "bold"),
        )

        c.create_oval(
            venn_cx - offset - venn_r, venn_cy - venn_r,
            venn_cx - offset + venn_r, venn_cy + venn_r,
            fill="", outline=ref_color, width=3,
            tags=("venn_ref",),
        )
        c.create_oval(
            venn_cx + offset - venn_r, venn_cy - venn_r,
            venn_cx + offset + venn_r, venn_cy + venn_r,
            fill="", outline=comp_color, width=3,
            tags=("venn_comp",),
        )

        c.create_text(
            venn_cx - offset - venn_r * 0.5, venn_cy - 10,
            text=f"{only_ref:,}", fill=ref_color,
            font=("", 18, "bold"),
        )
        c.create_text(
            venn_cx - offset - venn_r * 0.5, venn_cy + 12,
            text=t("stats_only_ref"), fill=fg_dim, font=("", 9),
        )

        c.create_text(
            venn_cx, venn_cy - 10,
            text=f"{match_count:,}", fill=intersect_color,
            font=("", 20, "bold"),
        )
        c.create_text(
            venn_cx, venn_cy + 14,
            text=t("stats_matched"), fill=fg_dim, font=("", 9),
        )

        c.create_text(
            venn_cx + offset + venn_r * 0.5, venn_cy - 10,
            text=f"{only_comp:,}", fill=comp_color,
            font=("", 18, "bold"),
        )
        c.create_text(
            venn_cx + offset + venn_r * 0.5, venn_cy + 12,
            text=t("stats_only_comp"), fill=fg_dim, font=("", 9),
        )

        pct_match = round(match_count / total * 100, 1) if total else 0
        c.create_text(
            venn_cx, venn_cy + venn_r + 28,
            text=f"{t('stats_match_pct')}: %{pct_match}",
            fill=fg, font=("", 14, "bold"),
        )

        c.tag_bind("venn_ref", "<Button-1>", _go_tab(tab_map["ref"]))
        c.tag_bind("venn_ref", "<Enter>", lambda e: c.configure(cursor="hand2"))
        c.tag_bind("venn_ref", "<Leave>", lambda e: c.configure(cursor=""))
        c.tag_bind("venn_comp", "<Button-1>", _go_tab(tab_map["comp"]))
        c.tag_bind("venn_comp", "<Enter>", lambda e: c.configure(cursor="hand2"))
        c.tag_bind("venn_comp", "<Leave>", lambda e: c.configure(cursor=""))

    def _fill_textbox(self, textbox, text):
        self._tab_full_texts[textbox] = text
        textbox.configure(state="normal")
        textbox.delete("0.0", "end")
        textbox.insert("0.0", text)
        textbox.configure(state="disabled")

    def _copy_active_tab(self):
        tab_name = self.tabview.get()
        tabs = {
            t("tab_matches"): self.txt_match,
            t("tab_no_matches"): self.txt_no_match,
            t("tab_duplicates"): self.txt_duplicates,
            t("tab_only_ref"): self.txt_ref,
            t("tab_only_comp"): self.txt_comp,
            t("tab_stats"): self.txt_stats,
        }
        textbox = tabs.get(tab_name)
        if not textbox:
            return
        content = self._tab_full_texts.get(textbox, "")
        if not content:
            return
        self.clipboard_clear()
        self.clipboard_append(content)
        old_text = self._copy_btn.cget("text")
        self._copy_btn.configure(text=t("copied"))
        self.after(1500, lambda: self._copy_btn.configure(text=old_text))

    def _on_search_changed(self, *_args):
        query = self.search_var.get().strip().lower()
        total_found = 0

        for textbox, full_text in self._tab_full_texts.items():
            if not query:
                filtered = full_text
            else:
                lines = full_text.split("\n")
                filtered_lines = [
                    l for l in lines if query in l.lower() or l.startswith("─") or l.startswith("=")
                ]
                total_found += sum(1 for l in lines if query in l.lower() and not l.startswith("─") and not l.startswith("="))
                filtered = "\n".join(filtered_lines) if filtered_lines else t("search_not_found", query=query)

            textbox.configure(state="normal")
            textbox.delete("0.0", "end")
            textbox.insert("0.0", filtered)
            textbox.configure(state="disabled")

        if query:
            self.search_count_label.configure(text=t("search_results", count=total_found))
        else:
            self.search_count_label.configure(text="")

    # ── Dışa Aktarma ──────────────────────────────────────────────

    def _export_html_report(self):
        if not self.results:
            return

        initial_dir = _load_settings().get("default_export_dir", "") or None
        path = filedialog.asksaveasfilename(
            title=t("save_report"),
            defaultextension=".html",
            initialdir=initial_dir,
            filetypes=[("HTML", "*.html")],
        )
        if not path:
            return

        try:
            self._generate_html_report(path)
            messagebox.showinfo(t("success_title"), t("report_saved", path=path))
            import webbrowser
            webbrowser.open(f"file://{os.path.abspath(path)}")
        except Exception as e:
            messagebox.showerror(t("error_title"), t("export_error", e=e))

    def _generate_html_report(self, output_path):
        r = self.results
        s = r["stats"]
        now = datetime.now().strftime("%d.%m.%Y %H:%M")

        match_count = s["match_count"]
        only_ref = s["only_ref_count"]
        only_comp = s["only_comp_count"]
        total = match_count + only_ref + only_comp
        pct_match = round(match_count / total * 100, 1) if total else 0
        pct_ref = round(only_ref / total * 100, 1) if total else 0
        pct_comp = round(only_comp / total * 100, 1) if total else 0

        ref_names = ", ".join(Path(p).name for p, _ in self.ref_files)
        comp_names = ", ".join(Path(p).name for p, _ in self.comp_files)

        match_rows = ""
        for v in r["matches"][:200]:
            match_rows += f"<tr><td>{v}</td></tr>\n"
        if len(r["matches"]) > 200:
            match_rows += f'<tr><td style="color:#888">... +{len(r["matches"])-200} {t("more_records")}</td></tr>\n'

        only_ref_rows = ""
        for v in r["only_in_reference"][:200]:
            only_ref_rows += f"<tr><td>{v}</td></tr>\n"
        if len(r["only_in_reference"]) > 200:
            only_ref_rows += f'<tr><td style="color:#888">... +{len(r["only_in_reference"])-200} {t("more_records")}</td></tr>\n'

        only_comp_rows = ""
        for v in r["only_in_comparison"][:200]:
            src = ", ".join(r["source_map"].get(v, ["?"]))
            only_comp_rows += f"<tr><td>{v}</td><td style='color:#888;font-size:12px'>{src}</td></tr>\n"
        if len(r["only_in_comparison"]) > 200:
            only_comp_rows += f'<tr><td colspan="2" style="color:#888">... +{len(r["only_in_comparison"])-200} {t("more_records")}</td></tr>\n'

        html = f"""<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="UTF-8">
<title>{t("report_title")}</title>
<style>
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{ font-family: 'Segoe UI', Tahoma, sans-serif; background:#f5f7fa; color:#1a1a2e; padding:30px; }}
  .container {{ max-width:900px; margin:0 auto; }}
  .header {{ text-align:center; margin-bottom:30px; }}
  .header h1 {{ font-size:24px; color:#16213e; }}
  .header .date {{ color:#888; font-size:13px; margin-top:4px; }}
  .files {{ background:#fff; border-radius:10px; padding:18px 22px; margin-bottom:20px; box-shadow:0 2px 8px rgba(0,0,0,0.06); }}
  .files .label {{ font-weight:600; color:#555; font-size:13px; }}
  .files .value {{ color:#16213e; margin-bottom:6px; }}
  .stats-grid {{ display:grid; grid-template-columns:1fr 1fr 1fr; gap:14px; margin-bottom:24px; }}
  .stat-card {{ background:#fff; border-radius:10px; padding:20px; text-align:center; box-shadow:0 2px 8px rgba(0,0,0,0.06); }}
  .stat-card .num {{ font-size:32px; font-weight:700; }}
  .stat-card .lbl {{ font-size:13px; color:#666; margin-top:4px; }}
  .stat-card.green .num {{ color:#27ae60; }}
  .stat-card.red .num {{ color:#e74c3c; }}
  .stat-card.orange .num {{ color:#f39c12; }}
  .chart-area {{ background:#fff; border-radius:10px; padding:24px; margin-bottom:24px; box-shadow:0 2px 8px rgba(0,0,0,0.06); text-align:center; }}
  .chart-area h3 {{ margin-bottom:16px; font-size:16px; }}
  .bar-chart {{ width:100%; }}
  .bar-row {{ display:flex; align-items:center; margin-bottom:10px; }}
  .bar-label {{ width:200px; text-align:right; padding-right:12px; font-size:13px; color:#555; }}
  .bar-track {{ flex:1; height:28px; background:#eee; border-radius:6px; overflow:hidden; position:relative; }}
  .bar-fill {{ height:100%; border-radius:6px; display:flex; align-items:center; padding-left:10px; color:#fff; font-size:12px; font-weight:600; min-width:40px; }}
  .bar-fill.green {{ background:#27ae60; }}
  .bar-fill.red {{ background:#e74c3c; }}
  .bar-fill.orange {{ background:#f39c12; }}
  .section {{ background:#fff; border-radius:10px; padding:18px 22px; margin-bottom:18px; box-shadow:0 2px 8px rgba(0,0,0,0.06); }}
  .section h3 {{ font-size:15px; color:#16213e; margin-bottom:10px; border-bottom:2px solid #eee; padding-bottom:6px; }}
  table {{ width:100%; border-collapse:collapse; }}
  table td {{ padding:5px 8px; border-bottom:1px solid #f0f0f0; font-size:13px; }}
  .footer {{ text-align:center; color:#aaa; font-size:11px; margin-top:30px; }}
  .pct-badge {{ display:inline-block; background:#16213e; color:#fff; border-radius:20px; padding:4px 14px; font-size:22px; font-weight:700; margin:8px 0; }}
  @media print {{
    body {{ padding:10px; }}
    .stat-card, .section, .chart-area, .files {{ box-shadow:none; border:1px solid #ddd; }}
  }}
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>{t("report_title")}</h1>
    <div class="date">{now}</div>
  </div>

  <div class="files">
    <div class="label">{t("stats_ref_total")}:</div>
    <div class="value">{ref_names} ({s["ref_total"]:,} {t("more_records")})</div>
    <div class="label">{t("stats_comp_total")}:</div>
    <div class="value">{comp_names} ({s["comp_total"]:,} {t("more_records")})</div>
  </div>

  <div class="stats-grid">
    <div class="stat-card green">
      <div class="num">{match_count:,}</div>
      <div class="lbl">{t("stats_matched")}</div>
    </div>
    <div class="stat-card red">
      <div class="num">{only_ref:,}</div>
      <div class="lbl">{t("stats_only_ref")}</div>
    </div>
    <div class="stat-card orange">
      <div class="num">{only_comp:,}</div>
      <div class="lbl">{t("stats_only_comp")}</div>
    </div>
  </div>

  <div class="chart-area">
    <h3>{t("chart_bar_title")}</h3>
    <div class="pct-badge">%{s["match_percentage"]} {t("stats_match_pct")}</div>
    <div class="bar-chart">
      <div class="bar-row">
        <div class="bar-label">{t("stats_matched")} ({match_count:,})</div>
        <div class="bar-track"><div class="bar-fill green" style="width:{pct_match}%">{pct_match}%</div></div>
      </div>
      <div class="bar-row">
        <div class="bar-label">{t("stats_only_ref")} ({only_ref:,})</div>
        <div class="bar-track"><div class="bar-fill red" style="width:{pct_ref}%">{pct_ref}%</div></div>
      </div>
      <div class="bar-row">
        <div class="bar-label">{t("stats_only_comp")} ({only_comp:,})</div>
        <div class="bar-track"><div class="bar-fill orange" style="width:{pct_comp}%">{pct_comp}%</div></div>
      </div>
    </div>
  </div>

  <div class="section">
    <h3>{t("tab_matches")} ({match_count:,})</h3>
    <table>{match_rows if match_rows else f"<tr><td>{t('no_matched_records')}</td></tr>"}</table>
  </div>

  <div class="section">
    <h3>{t("tab_only_ref")} ({only_ref:,})</h3>
    <table>{only_ref_rows if only_ref_rows else f"<tr><td>{t('all_records_match')}</td></tr>"}</table>
  </div>

  <div class="section">
    <h3>{t("tab_only_comp")} ({only_comp:,})</h3>
    <table>{only_comp_rows if only_comp_rows else f"<tr><td>{t('no_extra_comp')}</td></tr>"}</table>
  </div>

  <div class="footer">
    {t("report_footer")} — {now}
  </div>
</div>
</body>
</html>"""
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)

    def _export_results(self):
        if not self.results:
            return

        initial_dir = _load_settings().get("default_export_dir", "") or None
        path = filedialog.asksaveasfilename(
            title=t("save_results"),
            defaultextension=".xlsx",
            initialdir=initial_dir,
            filetypes=[
                (t("excel_file"), "*.xlsx"),
                (t("csv_files"), "*.csv"),
            ],
        )
        if not path:
            return

        try:
            enc = _load_settings().get("csv_encoding", "utf-8-sig")
            self.engine.export_results(self.results, path, csv_encoding=enc)
            if path.lower().endswith(".csv"):
                stem = Path(path).stem
                folder = str(Path(path).parent)
                messagebox.showinfo(
                    t("success_title"),
                    t("results_saved", path=f"{folder}/{stem}_*.csv"),
                )
            else:
                messagebox.showinfo(t("success_title"), t("results_saved", path=path))
        except Exception as e:
            messagebox.showerror(t("error_title"), t("export_error", e=e))


if __name__ == "__main__":
    app = App()
    app.mainloop()
