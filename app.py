import sys
import os
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path

if getattr(sys, "frozen", False):
    sys.path.insert(0, sys._MEIPASS)

from engine import ComparisonEngine

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

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".csv"}


class PreviewWindow(ctk.CTkToplevel):
    """Dosya içeriğini önizleme penceresi."""

    MAX_PREVIEW_ROWS = 200

    def __init__(self, master, file_path, engine):
        super().__init__(master)
        self.title(f"Önizleme — {Path(file_path).name}")
        self.geometry("800x500")
        self.minsize(500, 300)

        self._build(file_path, engine)

    def _build(self, file_path, engine):
        try:
            sheets = engine.get_sheet_names(file_path)
            if sheets:
                top = ctk.CTkFrame(self, fg_color="transparent")
                top.pack(fill="x", padx=10, pady=(10, 0))
                ctk.CTkLabel(top, text="Sayfa:").pack(side="left", padx=(0, 5))
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
                self, text=f"Dosya okunamadı:\n{e}", text_color="red"
            ).pack(expand=True, padx=20, pady=20)

    def _load_table(self, file_path, engine):
        sheet = getattr(self, "_sheet_var", None)
        sheet = sheet.get() if sheet else None

        try:
            df = engine.load_file(file_path, sheet)
        except Exception as e:
            self._info_label.configure(text=f"Hata: {e}")
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
                text=f"İlk {self.MAX_PREVIEW_ROWS} / {total:,} satır gösteriliyor"
            )
        else:
            self._info_label.configure(text=f"Toplam {total:,} satır")


class FileEntry(ctk.CTkFrame):
    """Dosya listesinde tek bir dosya satırı."""

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
            self,
            text="✕",
            width=30,
            height=24,
            fg_color="#c0392b",
            hover_color="#e74c3c",
            command=lambda: self.on_remove(self),
        ).grid(row=0, column=2, padx=(0, 5), pady=2)


class App(_BaseCTk):
    def __init__(self):
        super().__init__()

        self.title("Karşılaştırıcı - Dosya Karşılaştırma Aracı")
        self.geometry("1100x750")
        self.minsize(900, 650)

        self.engine = ComparisonEngine()
        self.ref_files = []
        self.comp_files = []
        self.results = None

        self._build_ui()

    # ── Drag & Drop Yardımcıları ─────────────────────────────────

    def _parse_drop_data(self, data):
        """Platform bağımsız sürüklenen dosya yollarını ayrıştır."""
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
        """Bir widget'a drop zone özelliği ekle (DnD varsa)."""
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
        self.grid_rowconfigure(2, weight=1)

        self._build_header()
        self._build_file_section()
        self._build_results_section()
        self._build_bottom_bar()

    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, padx=15, pady=(15, 0), sticky="ew")
        ctk.CTkLabel(header, text="Karşılaştırıcı", font=("", 24, "bold")).pack(
            side="left"
        )
        ctk.CTkLabel(
            header,
            text="Dosya Karşılaştırma Aracı",
            font=("", 14),
            text_color="gray",
        ).pack(side="left", padx=10)

    def _build_file_section(self):
        container = ctk.CTkFrame(self)
        container.grid(row=1, column=0, padx=15, pady=10, sticky="ew")
        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=1)

        self._build_ref_panel(container)
        self._build_comp_panel(container)

        btn_row = ctk.CTkFrame(container, fg_color="transparent")
        btn_row.grid(row=1, column=0, columnspan=2, padx=5, pady=10, sticky="ew")
        btn_row.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(
            btn_row,
            text="⇄  Tarafları Değiştir",
            font=("", 13),
            height=35,
            width=160,
            command=self._swap_sides,
            fg_color="#8e44ad",
            hover_color="#9b59b6",
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkButton(
            btn_row,
            text="KARŞILAŞTIR",
            font=("", 16, "bold"),
            height=45,
            command=self._run_comparison,
            fg_color="#27ae60",
            hover_color="#2ecc71",
        ).grid(row=0, column=0, sticky="e")

    def _build_ref_panel(self, parent):
        panel = ctk.CTkFrame(parent)
        panel.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        ctk.CTkLabel(panel, text="Referans Dosyaları", font=("", 16, "bold")).pack(
            pady=(10, 5)
        )
        ctk.CTkLabel(
            panel, text="Ana dosyalarınızı ekleyin", text_color="gray", font=("", 11)
        ).pack()

        self.ref_list_frame = ctk.CTkScrollableFrame(panel, height=60)
        self.ref_list_frame.pack(fill="x", padx=15, pady=5)
        self.ref_list_frame.grid_columnconfigure(0, weight=1)

        drop_hint = "Dosya eklenmedi\n(veya buraya sürükle-bırak)" if DND_AVAILABLE else "Dosya eklenmedi"
        self.ref_empty_label = ctk.CTkLabel(
            self.ref_list_frame, text=drop_hint, text_color="#888"
        )
        self.ref_empty_label.grid(row=0, column=0)

        self._enable_drop(panel, self._on_ref_drop)

        btn_row = ctk.CTkFrame(panel, fg_color="transparent")
        btn_row.pack(pady=5)
        ctk.CTkButton(
            btn_row, text="Dosya Ekle", command=self._add_ref_files, width=100
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            btn_row, text="Klasör Ekle", command=self._add_ref_folder, width=100,
            fg_color="#2980b9", hover_color="#3498db",
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            btn_row,
            text="Temizle",
            command=self._clear_ref_files,
            width=80,
            fg_color="#c0392b",
            hover_color="#e74c3c",
        ).pack(side="left", padx=2)

        ctk.CTkLabel(panel, text="Sütunlar:", font=("", 12), anchor="w").pack(
            fill="x", padx=15
        )
        self.ref_col_frame = ctk.CTkScrollableFrame(panel, height=50)
        self.ref_col_frame.pack(fill="x", padx=15, pady=(0, 10))
        self.ref_col_checks = {}
        ctk.CTkLabel(
            self.ref_col_frame, text="Önce dosya ekleyin", text_color="#888"
        ).pack()

    def _build_comp_panel(self, parent):
        panel = ctk.CTkFrame(parent)
        panel.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

        ctk.CTkLabel(
            panel, text="Karşılaştırma Dosyaları", font=("", 16, "bold")
        ).pack(pady=(10, 5))
        ctk.CTkLabel(
            panel,
            text="Karşılaştırılacak dosyaları ekleyin",
            text_color="gray",
            font=("", 11),
        ).pack()

        self.comp_list_frame = ctk.CTkScrollableFrame(panel, height=60)
        self.comp_list_frame.pack(fill="x", padx=15, pady=5)
        self.comp_list_frame.grid_columnconfigure(0, weight=1)

        drop_hint = "Dosya eklenmedi\n(veya buraya sürükle-bırak)" if DND_AVAILABLE else "Dosya eklenmedi"
        self.comp_empty_label = ctk.CTkLabel(
            self.comp_list_frame, text=drop_hint, text_color="#888"
        )
        self.comp_empty_label.grid(row=0, column=0)

        self._enable_drop(panel, self._on_comp_drop)

        btn_row = ctk.CTkFrame(panel, fg_color="transparent")
        btn_row.pack(pady=5)
        ctk.CTkButton(
            btn_row, text="Dosya Ekle", command=self._add_comp_files, width=100
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            btn_row, text="Klasör Ekle", command=self._add_comp_folder, width=100,
            fg_color="#2980b9", hover_color="#3498db",
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            btn_row,
            text="Temizle",
            command=self._clear_comp_files,
            width=80,
            fg_color="#c0392b",
            hover_color="#e74c3c",
        ).pack(side="left", padx=2)

        ctk.CTkLabel(panel, text="Sütunlar:", font=("", 12), anchor="w").pack(
            fill="x", padx=15
        )
        self.comp_col_frame = ctk.CTkScrollableFrame(panel, height=50)
        self.comp_col_frame.pack(fill="x", padx=15, pady=(0, 10))
        self.comp_col_checks = {}
        ctk.CTkLabel(
            self.comp_col_frame, text="Önce dosya ekleyin", text_color="#888"
        ).pack()

    def _build_results_section(self):
        results_frame = ctk.CTkFrame(self)
        results_frame.grid(row=2, column=0, padx=15, pady=(0, 5), sticky="nsew")
        results_frame.grid_rowconfigure(1, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)

        search_bar = ctk.CTkFrame(results_frame, fg_color="transparent")
        search_bar.grid(row=0, column=0, padx=10, pady=(8, 0), sticky="ew")
        search_bar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(search_bar, text="Ara:", font=("", 13)).grid(
            row=0, column=0, padx=(0, 5)
        )
        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", self._on_search_changed)
        self.search_entry = ctk.CTkEntry(
            search_bar, textvariable=self.search_var,
            placeholder_text="Sonuçlarda ara...", height=30,
        )
        self.search_entry.grid(row=0, column=1, sticky="ew")

        self.search_count_label = ctk.CTkLabel(
            search_bar, text="", font=("", 11), text_color="gray"
        )
        self.search_count_label.grid(row=0, column=2, padx=(8, 0))

        self.tabview = ctk.CTkTabview(results_frame)
        self.tabview.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="nsew")

        self._tab_full_texts = {}

        tabs = {
            "Eşleşenler": "txt_match",
            "Eşleşmeyenler": "txt_no_match",
            "Tekrar Edenler": "txt_duplicates",
            "Yalnızca Referansta": "txt_ref",
            "Yalnızca Karşılaştırmada": "txt_comp",
            "İstatistikler": "txt_stats",
        }
        for tab_name, attr_name in tabs.items():
            tab = self.tabview.add(tab_name)
            textbox = ctk.CTkTextbox(tab, state="disabled")
            textbox.pack(fill="both", expand=True, padx=5, pady=5)
            setattr(self, attr_name, textbox)

    def _build_bottom_bar(self):
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=3, column=0, padx=15, pady=(0, 10), sticky="ew")
        bar.grid_columnconfigure(1, weight=1)

        self.status_label = ctk.CTkLabel(bar, text="Hazır", text_color="gray")
        self.status_label.grid(row=0, column=0, sticky="w")

        self.progress_bar = ctk.CTkProgressBar(bar, height=12)
        self.progress_bar.grid(row=0, column=1, padx=10, sticky="ew")
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(bar, text="", text_color="gray", font=("", 11))
        self.progress_label.grid(row=0, column=2, padx=(0, 10))

        self.export_btn = ctk.CTkButton(
            bar,
            text="Sonuçları Dışa Aktar (Excel)",
            command=self._export_results,
            state="disabled",
        )
        self.export_btn.grid(row=0, column=3, sticky="e")

    # ── Dosya / Klasör Seçimi ────────────────────────────────────────

    def _scan_folder(self, folder_path):
        """Klasördeki tüm desteklenen dosyaları bul (alt klasörler dahil)."""
        folder = Path(folder_path)
        files = []
        for ext in SUPPORTED_EXTENSIONS:
            files.extend(folder.rglob(f"*{ext}"))
        return sorted(str(f) for f in files if f.is_file())

    def _add_ref_folder(self):
        folder = filedialog.askdirectory(title="Referans Klasörü Seçin")
        if not folder:
            return
        files = self._scan_folder(folder)
        if not files:
            messagebox.showinfo("Bilgi", "Klasörde desteklenen dosya bulunamadı.\n(.xlsx, .xls, .csv)")
            return
        self._add_files_to_ref(files)

    def _add_comp_folder(self):
        folder = filedialog.askdirectory(title="Karşılaştırma Klasörü Seçin")
        if not folder:
            return
        files = self._scan_folder(folder)
        if not files:
            messagebox.showinfo("Bilgi", "Klasörde desteklenen dosya bulunamadı.\n(.xlsx, .xls, .csv)")
            return
        self._add_files_to_comp(files)

    def _file_types(self):
        return [
            ("Excel Dosyaları", "*.xlsx *.xls"),
            ("CSV Dosyaları", "*.csv"),
            ("Tüm Dosyalar", "*.*"),
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
            title="Referans Dosyalarını Seçin", filetypes=self._file_types()
        )
        if not paths:
            return
        self._add_files_to_ref(paths)

    def _add_files_to_ref(self, paths):

        existing_paths = {p for p, _ in self.ref_files}
        first_new = len(self.ref_files) == 0

        for p in paths:
            if p not in existing_paths:
                try:
                    sheets = self.engine.get_sheet_names(p)
                    sheet = sheets[0] if sheets else None
                    self.ref_files.append((p, sheet))
                except Exception as e:
                    messagebox.showerror(
                        "Dosya Okuma Hatası",
                        f"'{Path(p).name}' dosyası okunamadı:\n{e}",
                    )

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
                self.ref_list_frame, text="Dosya eklenmedi", text_color="#888"
            )
            self.ref_empty_label.grid(row=0, column=0)
            return

        for i, (path, _) in enumerate(self.ref_files):
            entry = FileEntry(self.ref_list_frame, path, self._remove_ref_file, self._preview_file)
            entry.grid(row=i, column=0, sticky="ew", pady=1)

    def _add_comp_files(self):
        paths = filedialog.askopenfilenames(
            title="Karşılaştırma Dosyalarını Seçin", filetypes=self._file_types()
        )
        if not paths:
            return
        self._add_files_to_comp(paths)

    def _add_files_to_comp(self, paths):
        existing_paths = {p for p, _ in self.comp_files}
        first_new = len(self.comp_files) == 0

        for p in paths:
            if p not in existing_paths:
                try:
                    sheets = self.engine.get_sheet_names(p)
                    sheet = sheets[0] if sheets else None
                    self.comp_files.append((p, sheet))
                except Exception as e:
                    messagebox.showerror(
                        "Dosya Okuma Hatası",
                        f"'{Path(p).name}' dosyası okunamadı:\n{e}",
                    )

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
                self.comp_list_frame, text="Dosya eklenmedi", text_color="#888"
            )
            self.comp_empty_label.grid(row=0, column=0)
            return

        for i, (path, _) in enumerate(self.comp_files):
            entry = FileEntry(self.comp_list_frame, path, self._remove_comp_file, self._preview_file)
            entry.grid(row=i, column=0, sticky="ew", pady=1)

    # ── Sütun Checkbox Yönetimi ────────────────────────────────────

    def _rebuild_col_checks(self, side):
        """Dosya listesine göre sütun checkbox'larını yeniden oluştur."""
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
            ctk.CTkLabel(frame, text="Önce dosya ekleyin", text_color="#888").pack()
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
            ctk.CTkLabel(frame, text="Sütun bulunamadı", text_color="#888").pack()
            return

        for col in all_columns:
            var = ctk.BooleanVar(value=False)
            cb = ctk.CTkCheckBox(frame, text=col, variable=var)
            cb.pack(anchor="w", padx=5, pady=1)
            checks[col] = var

    def _swap_sides(self):
        """Referans ve karşılaştırma dosyalarını yer değiştirir."""
        self.ref_files, self.comp_files = self.comp_files, self.ref_files
        self._refresh_ref_list()
        self._refresh_comp_list()
        self._rebuild_col_checks("ref")
        self._rebuild_col_checks("comp")

    def _get_selected_columns(self, side):
        checks = self.ref_col_checks if side == "ref" else self.comp_col_checks
        return [col for col, var in checks.items() if var.get()]

    # ── Karşılaştırma ─────────────────────────────────────────────

    def _run_comparison(self):
        if not self.ref_files:
            messagebox.showwarning("Uyarı", "Lütfen en az bir referans dosya ekleyin.")
            return
        if not self.comp_files:
            messagebox.showwarning(
                "Uyarı", "Lütfen en az bir karşılaştırma dosyası ekleyin."
            )
            return

        ref_cols = self._get_selected_columns("ref")
        comp_cols = self._get_selected_columns("comp")
        if not ref_cols or not comp_cols:
            messagebox.showwarning(
                "Uyarı", "Lütfen her iki taraftan da en az bir sütun seçin."
            )
            return

        self.progress_bar.set(0)
        self.progress_label.configure(text="")
        self.status_label.configure(text="Karşılaştırılıyor...", text_color="yellow")
        self.update()

        def on_progress(checked, total, phase):
            ratio = checked / total if total > 0 else 0
            self.progress_bar.set(ratio)
            pct = int(ratio * 100)
            self.progress_label.configure(
                text=f"{checked:,} / {total:,} satır  ({pct}%)"
            )
            self.status_label.configure(text=phase, text_color="yellow")
            self.update_idletasks()

        try:
            self.results = self.engine.compare(
                self.ref_files, ref_cols, self.comp_files, comp_cols,
                on_progress=on_progress,
            )
            self.progress_bar.set(1)
            self._display_results()
            self.export_btn.configure(state="normal")
            s = self.results["stats"]
            no_match = s["only_ref_count"] + s["only_comp_count"]
            self.progress_label.configure(text="")
            self.status_label.configure(
                text=f"Tamamlandı! {s['match_count']} eşleşen, {no_match} eşleşmeyen.",
                text_color="#2ecc71",
            )
        except Exception as e:
            self.progress_bar.set(0)
            self.progress_label.configure(text="")
            messagebox.showerror(
                "Hata", f"Karşılaştırma sırasında hata oluştu:\n{e}"
            )
            self.status_label.configure(text="Hata oluştu", text_color="red")

    def _display_results(self):
        r = self.results
        s = r["stats"]

        if r["matches"]:
            lines = [f"  • {v}" for v in r["matches"]]
            match_text = (
                f"Eşleşen Kayıtlar ({s['match_count']} adet)\n"
                + "=" * 50
                + "\n\n"
                + "\n".join(lines)
            )
        else:
            match_text = "Eşleşen kayıt bulunamadı."
        self._fill_textbox(self.txt_match, match_text)

        # Eşleşmeyenler: her iki taraftaki eşleşmeyenleri birleştir
        no_match_lines = []
        if r["only_in_reference"]:
            no_match_lines.append(
                f"── Referansta olup karşılaştırmada OLMAYAN ({s['only_ref_count']} adet) ──\n"
            )
            for v in r["only_in_reference"]:
                no_match_lines.append(f"  ✗ {v}")
            no_match_lines.append("")

        if r["only_in_comparison"]:
            no_match_lines.append(
                f"── Karşılaştırmada olup referansta OLMAYAN ({s['only_comp_count']} adet) ──\n"
            )
            for v in r["only_in_comparison"]:
                sources = ", ".join(r["source_map"].get(v, ["?"]))
                no_match_lines.append(f"  ✗ {v}  ← [{sources}]")

        total_no_match = s["only_ref_count"] + s["only_comp_count"]
        if no_match_lines:
            no_match_text = (
                f"Eşleşmeyen Kayıtlar (toplam {total_no_match} adet)\n"
                + "=" * 50
                + "\n\n"
                + "\n".join(no_match_lines)
            )
        else:
            no_match_text = "Tüm kayıtlar eşleşiyor, eşleşmeyen kayıt yok!"
        self._fill_textbox(self.txt_no_match, no_match_text)

        # Tekrar Edenler
        dup_lines = []
        if r["ref_duplicates"]:
            dup_lines.append(
                f"── Referansta Tekrar Edenler ({s['ref_dup_count']} farklı kayıt) ──\n"
            )
            for val, details in sorted(r["ref_duplicates"].items()):
                for d in details:
                    dup_lines.append(f"  ⚠ {val}  ← {d}")
            dup_lines.append("")

        if r["comp_duplicates"]:
            dup_lines.append(
                f"── Karşılaştırmada Tekrar Edenler ({s['comp_dup_count']} farklı kayıt) ──\n"
            )
            for val, details in sorted(r["comp_duplicates"].items()):
                for d in details:
                    dup_lines.append(f"  ⚠ {val}  ← {d}")

        total_dups = s["ref_dup_count"] + s["comp_dup_count"]
        if dup_lines:
            dup_text = (
                f"Tekrar Eden Kayıtlar (toplam {total_dups} farklı kayıt)\n"
                + "=" * 50
                + "\n\n"
                + "\n".join(dup_lines)
            )
        else:
            dup_text = "Tekrar eden kayıt bulunamadı."
        self._fill_textbox(self.txt_duplicates, dup_text)

        if r["only_in_reference"]:
            lines = [f"  • {v}" for v in r["only_in_reference"]]
            ref_text = (
                f"Yalnızca Referansta ({s['only_ref_count']} adet)\n"
                + "=" * 50
                + "\n\n"
                + "\n".join(lines)
            )
        else:
            ref_text = "Tüm kayıtlar eşleşiyor."
        self._fill_textbox(self.txt_ref, ref_text)

        if r["only_in_comparison"]:
            lines = []
            for v in r["only_in_comparison"]:
                sources = ", ".join(r["source_map"].get(v, ["?"]))
                lines.append(f"  • {v}  ← [{sources}]")
            comp_text = (
                f"Yalnızca Karşılaştırmada ({s['only_comp_count']} adet)\n"
                + "=" * 50
                + "\n\n"
                + "\n".join(lines)
            )
        else:
            comp_text = "Karşılaştırma dosyalarında fazladan kayıt yok."
        self._fill_textbox(self.txt_comp, comp_text)

        stats_text = (
            "╔════════════════════════════════════════════╗\n"
            "║        KARŞILAŞTIRMA SONUÇLARI             ║\n"
            "╠════════════════════════════════════════════╣\n"
            f"║  Referans Kayıt Sayısı:       {s['ref_total']:>6}       ║\n"
            f"║  Karşılaştırma Kayıt Sayısı:  {s['comp_total']:>6}       ║\n"
            "╠════════════════════════════════════════════╣\n"
            f"║  Eşleşen:                     {s['match_count']:>6}       ║\n"
            f"║  Yalnızca Referansta:         {s['only_ref_count']:>6}       ║\n"
            f"║  Yalnızca Karşılaştırmada:   {s['only_comp_count']:>6}       ║\n"
            "╠════════════════════════════════════════════╣\n"
            f"║  Eşleşme Oranı:              %{s['match_percentage']:>5}       ║\n"
            "╚════════════════════════════════════════════╝"
        )
        self._fill_textbox(self.txt_stats, stats_text)

    def _fill_textbox(self, textbox, text):
        self._tab_full_texts[textbox] = text
        textbox.configure(state="normal")
        textbox.delete("0.0", "end")
        textbox.insert("0.0", text)
        textbox.configure(state="disabled")

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
                filtered = "\n".join(filtered_lines) if filtered_lines else f"'{query}' bulunamadı."

            textbox.configure(state="normal")
            textbox.delete("0.0", "end")
            textbox.insert("0.0", filtered)
            textbox.configure(state="disabled")

        if query:
            self.search_count_label.configure(text=f"{total_found} sonuç")
        else:
            self.search_count_label.configure(text="")

    # ── Dışa Aktarma ──────────────────────────────────────────────

    def _export_results(self):
        if not self.results:
            return

        path = filedialog.asksaveasfilename(
            title="Sonuçları Kaydet",
            defaultextension=".xlsx",
            filetypes=[("Excel Dosyası", "*.xlsx")],
        )
        if not path:
            return

        try:
            self.engine.export_results(self.results, path)
            messagebox.showinfo("Başarılı", f"Sonuçlar kaydedildi:\n{path}")
        except Exception as e:
            messagebox.showerror("Hata", f"Dışa aktarma hatası:\n{e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
