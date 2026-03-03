import sys
import os
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path

if getattr(sys, "frozen", False):
    sys.path.insert(0, sys._MEIPASS)

from engine import ComparisonEngine

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class FileEntry(ctk.CTkFrame):
    """Dosya listesinde tek bir dosya satırı."""

    def __init__(self, master, file_path, on_remove, **kwargs):
        super().__init__(master, **kwargs)
        self.file_path = file_path
        self.on_remove = on_remove
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text=Path(file_path).name, anchor="w").grid(
            row=0, column=0, padx=5, sticky="ew"
        )
        ctk.CTkButton(
            self,
            text="✕",
            width=30,
            height=24,
            fg_color="#c0392b",
            hover_color="#e74c3c",
            command=lambda: self.on_remove(self),
        ).grid(row=0, column=1, padx=5, pady=2)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Karşılaştırıcı - Dosya Karşılaştırma Aracı")
        self.geometry("1100x750")
        self.minsize(900, 650)

        self.engine = ComparisonEngine()
        self.ref_path = None
        self.comp_files = []
        self.results = None

        self._build_ui()

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

        ctk.CTkButton(
            container,
            text="KARŞILAŞTIR",
            font=("", 16, "bold"),
            height=45,
            command=self._run_comparison,
            fg_color="#27ae60",
            hover_color="#2ecc71",
        ).grid(row=1, column=0, columnspan=2, padx=5, pady=10, sticky="ew")

    def _build_ref_panel(self, parent):
        panel = ctk.CTkFrame(parent)
        panel.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        ctk.CTkLabel(panel, text="Referans Dosya", font=("", 16, "bold")).pack(
            pady=(10, 5)
        )
        ctk.CTkLabel(
            panel, text="Ana dosyanızı seçin", text_color="gray", font=("", 11)
        ).pack()

        self.ref_label = ctk.CTkLabel(
            panel, text="Henüz dosya seçilmedi", text_color="#888"
        )
        self.ref_label.pack(pady=5)

        ctk.CTkButton(panel, text="Dosya Seç", command=self._select_ref_file).pack(
            pady=5
        )

        opts = ctk.CTkFrame(panel, fg_color="transparent")
        opts.pack(fill="x", padx=15, pady=(5, 10))
        opts.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(opts, text="Sayfa:").grid(row=0, column=0, padx=(0, 5), sticky="w")
        self.ref_sheet_var = ctk.StringVar(value="--")
        self.ref_sheet_menu = ctk.CTkOptionMenu(
            opts,
            variable=self.ref_sheet_var,
            values=["--"],
            command=self._on_ref_sheet_change,
            dynamic_resizing=False,
        )
        self.ref_sheet_menu.grid(row=0, column=1, sticky="ew", pady=2)

        ctk.CTkLabel(opts, text="Sütun:").grid(row=1, column=0, padx=(0, 5), sticky="w")
        self.ref_col_var = ctk.StringVar(value="--")
        self.ref_col_menu = ctk.CTkOptionMenu(
            opts,
            variable=self.ref_col_var,
            values=["--"],
            dynamic_resizing=False,
        )
        self.ref_col_menu.grid(row=1, column=1, sticky="ew", pady=2)

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

        self.comp_empty_label = ctk.CTkLabel(
            self.comp_list_frame, text="Dosya eklenmedi", text_color="#888"
        )
        self.comp_empty_label.grid(row=0, column=0)

        btn_row = ctk.CTkFrame(panel, fg_color="transparent")
        btn_row.pack(pady=5)
        ctk.CTkButton(
            btn_row, text="Dosya Ekle", command=self._add_comp_files, width=120
        ).pack(side="left", padx=3)
        ctk.CTkButton(
            btn_row,
            text="Tümünü Temizle",
            command=self._clear_comp_files,
            width=120,
            fg_color="#c0392b",
            hover_color="#e74c3c",
        ).pack(side="left", padx=3)

        comp_opts = ctk.CTkFrame(panel, fg_color="transparent")
        comp_opts.pack(fill="x", padx=15, pady=(5, 10))
        comp_opts.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(comp_opts, text="Sütun:").grid(
            row=0, column=0, padx=(0, 5), sticky="w"
        )
        self.comp_col_var = ctk.StringVar(value="--")
        self.comp_col_menu = ctk.CTkOptionMenu(
            comp_opts,
            variable=self.comp_col_var,
            values=["--"],
            dynamic_resizing=False,
        )
        self.comp_col_menu.grid(row=0, column=1, sticky="ew", pady=2)

    def _build_results_section(self):
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=2, column=0, padx=15, pady=(0, 5), sticky="nsew")

        tabs = {
            "Eşleşenler": "txt_match",
            "Eşleşmeyenler": "txt_no_match",
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

        self.status_label = ctk.CTkLabel(bar, text="Hazır", text_color="gray")
        self.status_label.pack(side="left")

        self.export_btn = ctk.CTkButton(
            bar,
            text="Sonuçları Dışa Aktar (Excel)",
            command=self._export_results,
            state="disabled",
        )
        self.export_btn.pack(side="right")

    # ── Dosya Seçimi ───────────────────────────────────────────────

    def _file_types(self):
        return [
            ("Excel Dosyaları", "*.xlsx *.xls"),
            ("CSV Dosyaları", "*.csv"),
            ("Tüm Dosyalar", "*.*"),
        ]

    def _select_ref_file(self):
        path = filedialog.askopenfilename(
            title="Referans Dosya Seçin", filetypes=self._file_types()
        )
        if not path:
            return

        self.ref_path = path
        self.ref_label.configure(text=Path(path).name, text_color="white")

        try:
            sheets = self.engine.get_sheet_names(path)
            if sheets:
                self.ref_sheet_menu.configure(values=sheets)
                self.ref_sheet_var.set(sheets[0])
                self._on_ref_sheet_change(sheets[0])
            else:
                self.ref_sheet_menu.configure(values=["--"])
                self.ref_sheet_var.set("--")
                self._load_ref_columns(path, None)
        except Exception as e:
            messagebox.showerror("Dosya Okuma Hatası", f"Dosya okunamadı:\n{e}")

    def _on_ref_sheet_change(self, sheet_name):
        if not self.ref_path or sheet_name == "--":
            return
        self._load_ref_columns(self.ref_path, sheet_name)

    def _load_ref_columns(self, path, sheet):
        try:
            cols = self.engine.get_columns(path, sheet)
            self.ref_col_menu.configure(values=cols if cols else ["--"])
            self.ref_col_var.set(cols[0] if cols else "--")
        except Exception as e:
            messagebox.showerror("Hata", str(e))

    def _add_comp_files(self):
        paths = filedialog.askopenfilenames(
            title="Karşılaştırma Dosyalarını Seçin", filetypes=self._file_types()
        )
        if not paths:
            return

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
            path, sheet = self.comp_files[0]
            try:
                cols = self.engine.get_columns(path, sheet)
                self.comp_col_menu.configure(values=cols if cols else ["--"])
                self.comp_col_var.set(cols[0] if cols else "--")
            except Exception as e:
                messagebox.showerror("Hata", str(e))

    def _remove_comp_file(self, entry_widget):
        self.comp_files = [
            (p, s) for p, s in self.comp_files if p != entry_widget.file_path
        ]
        self._refresh_comp_list()
        if not self.comp_files:
            self.comp_col_menu.configure(values=["--"])
            self.comp_col_var.set("--")

    def _clear_comp_files(self):
        self.comp_files = []
        self._refresh_comp_list()
        self.comp_col_menu.configure(values=["--"])
        self.comp_col_var.set("--")

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
            entry = FileEntry(self.comp_list_frame, path, self._remove_comp_file)
            entry.grid(row=i, column=0, sticky="ew", pady=1)

    # ── Karşılaştırma ─────────────────────────────────────────────

    def _run_comparison(self):
        if not self.ref_path:
            messagebox.showwarning("Uyarı", "Lütfen bir referans dosya seçin.")
            return
        if not self.comp_files:
            messagebox.showwarning(
                "Uyarı", "Lütfen en az bir karşılaştırma dosyası ekleyin."
            )
            return

        ref_col = self.ref_col_var.get()
        comp_col = self.comp_col_var.get()
        if ref_col == "--" or comp_col == "--":
            messagebox.showwarning("Uyarı", "Lütfen karşılaştırılacak sütunları seçin.")
            return

        ref_sheet = self.ref_sheet_var.get()
        if ref_sheet == "--":
            ref_sheet = None

        self.status_label.configure(text="Karşılaştırılıyor...", text_color="yellow")
        self.update()

        try:
            self.results = self.engine.compare(
                self.ref_path, ref_sheet, ref_col, self.comp_files, comp_col
            )
            self._display_results()
            self.export_btn.configure(state="normal")
            s = self.results["stats"]
            no_match = s["only_ref_count"] + s["only_comp_count"]
            self.status_label.configure(
                text=f"Tamamlandı! {s['match_count']} eşleşen, {no_match} eşleşmeyen.",
                text_color="#2ecc71",
            )
        except Exception as e:
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
        textbox.configure(state="normal")
        textbox.delete("0.0", "end")
        textbox.insert("0.0", text)
        textbox.configure(state="disabled")

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
