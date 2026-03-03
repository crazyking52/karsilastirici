import pandas as pd
from pathlib import Path


class ComparisonEngine:

    def load_file(self, file_path, sheet_name=None):
        suffix = Path(file_path).suffix.lower()
        if suffix in (".xlsx", ".xls"):
            if sheet_name and sheet_name != "--":
                return pd.read_excel(file_path, sheet_name=sheet_name)
            return pd.read_excel(file_path)
        elif suffix == ".csv":
            return pd.read_csv(file_path)
        else:
            raise ValueError(f"Desteklenmeyen dosya formatı: {suffix}")

    def get_sheet_names(self, file_path):
        if Path(file_path).suffix.lower() in (".xlsx", ".xls"):
            return pd.ExcelFile(file_path).sheet_names
        return []

    def get_columns(self, file_path, sheet_name=None):
        df = self.load_file(file_path, sheet_name)
        return list(df.columns.astype(str))

    def _count_total_rows(self, files):
        """Toplam satır sayısını dosyaları okumadan önce hesapla."""
        total = 0
        for file_path, sheet_name in files:
            df = self.load_file(file_path, sheet_name)
            total += len(df)
        return total

    def compare(self, ref_files, ref_columns, comp_files, comp_columns,
                on_progress=None):
        """
        ref_files: [(dosya_yolu, sayfa_adı), ...] listesi
        ref_columns: [sütun_adı, ...] listesi
        comp_files: [(dosya_yolu, sayfa_adı), ...] listesi
        comp_columns: [sütun_adı, ...] listesi
        on_progress: callback(checked, total, phase) -- ilerleme bildirimi
        """
        total_rows = self._count_total_rows(ref_files) + self._count_total_rows(comp_files)
        checked = 0

        ref_values = set()
        ref_source_map = {}
        ref_duplicates = {}

        for file_path, sheet_name in ref_files:
            ref_df = self.load_file(file_path, sheet_name)
            fname = Path(file_path).name
            for col in ref_columns:
                if col not in ref_df.columns:
                    raise ValueError(
                        f"'{col}' sütunu '{fname}' dosyasında bulunamadı.\n"
                        f"Mevcut sütunlar: {', '.join(ref_df.columns.astype(str))}"
                    )
            row_count = len(ref_df)
            for col in ref_columns:
                col_values = ref_df[col].dropna().astype(str).str.strip()
                counts = col_values.value_counts()
                for val, cnt in counts.items():
                    if cnt > 1:
                        key = f"{val}"
                        if key not in ref_duplicates:
                            ref_duplicates[key] = []
                        ref_duplicates[key].append(f"{fname} → {col}: {cnt} kez")
                values = set(col_values)
                ref_values.update(values)
                for v in values:
                    if v not in ref_source_map:
                        ref_source_map[v] = []
                    source = f"{fname} → {col}"
                    if source not in ref_source_map[v]:
                        ref_source_map[v].append(source)
            checked += row_count
            if on_progress:
                on_progress(checked, total_rows, "Referans okunuyor...")

        all_comp_values = set()
        source_map = {}
        comp_duplicates = {}

        for file_path, sheet_name in comp_files:
            comp_df = self.load_file(file_path, sheet_name)
            fname = Path(file_path).name
            for col in comp_columns:
                if col not in comp_df.columns:
                    raise ValueError(
                        f"'{col}' sütunu '{fname}' dosyasında bulunamadı.\n"
                        f"Mevcut sütunlar: {', '.join(comp_df.columns.astype(str))}"
                    )
            row_count = len(comp_df)
            for col in comp_columns:
                col_values = comp_df[col].dropna().astype(str).str.strip()
                counts = col_values.value_counts()
                for val, cnt in counts.items():
                    if cnt > 1:
                        key = f"{val}"
                        if key not in comp_duplicates:
                            comp_duplicates[key] = []
                        comp_duplicates[key].append(f"{fname} → {col}: {cnt} kez")
                values = set(col_values)
                all_comp_values.update(values)
                for v in values:
                    if v not in source_map:
                        source_map[v] = []
                    source = f"{fname} → {col}"
                    if source not in source_map[v]:
                        source_map[v].append(source)
            checked += row_count
            if on_progress:
                on_progress(checked, total_rows, "Karşılaştırılıyor...")

        matches = ref_values & all_comp_values
        only_in_ref = ref_values - all_comp_values
        only_in_comp = all_comp_values - ref_values

        total = len(ref_values)
        match_pct = round(len(matches) / total * 100, 1) if total > 0 else 0

        return {
            "matches": sorted(matches),
            "only_in_reference": sorted(only_in_ref),
            "only_in_comparison": sorted(only_in_comp),
            "ref_source_map": ref_source_map,
            "source_map": source_map,
            "ref_duplicates": ref_duplicates,
            "comp_duplicates": comp_duplicates,
            "stats": {
                "ref_total": total,
                "comp_total": len(all_comp_values),
                "match_count": len(matches),
                "only_ref_count": len(only_in_ref),
                "only_comp_count": len(only_in_comp),
                "match_percentage": match_pct,
                "ref_dup_count": len(ref_duplicates),
                "comp_dup_count": len(comp_duplicates),
            },
        }

    def export_results(self, results, output_path):
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            if results["matches"]:
                pd.DataFrame({"Eşleşen Kayıtlar": results["matches"]}).to_excel(
                    writer, sheet_name="Eşleşenler", index=False
                )

            no_match_rows = []
            for v in results["only_in_reference"]:
                ref_src = ", ".join(results["ref_source_map"].get(v, ["?"]))
                no_match_rows.append({"Kayıt": v, "Durum": f"Referansta var ({ref_src}), karşılaştırmada yok"})
            for v in results["only_in_comparison"]:
                sources = ", ".join(results["source_map"].get(v, ["?"]))
                no_match_rows.append({"Kayıt": v, "Durum": f"Karşılaştırmada var ({sources}), referansta yok"})
            if no_match_rows:
                pd.DataFrame(no_match_rows).to_excel(
                    writer, sheet_name="Eşleşmeyenler", index=False
                )

            if results["only_in_reference"]:
                pd.DataFrame(
                    {"Yalnızca Referansta": results["only_in_reference"]}
                ).to_excel(writer, sheet_name="Yalnızca Referans", index=False)

            if results["only_in_comparison"]:
                data = {
                    "Kayıt": results["only_in_comparison"],
                    "Kaynak Dosya": [
                        ", ".join(results["source_map"].get(v, ["?"]))
                        for v in results["only_in_comparison"]
                    ],
                }
                pd.DataFrame(data).to_excel(
                    writer, sheet_name="Yalnızca Karşılaştırma", index=False
                )

            dup_rows = []
            for val, details in results["ref_duplicates"].items():
                for d in details:
                    dup_rows.append({"Kayıt": val, "Konum": d, "Taraf": "Referans"})
            for val, details in results["comp_duplicates"].items():
                for d in details:
                    dup_rows.append({"Kayıt": val, "Konum": d, "Taraf": "Karşılaştırma"})
            if dup_rows:
                pd.DataFrame(dup_rows).to_excel(
                    writer, sheet_name="Tekrar Edenler", index=False
                )

            stats = results["stats"]
            stats_df = pd.DataFrame(
                {
                    "Metrik": [
                        "Referans Toplam Kayıt",
                        "Karşılaştırma Toplam Kayıt",
                        "Eşleşen Kayıt",
                        "Yalnızca Referansta",
                        "Yalnızca Karşılaştırmada",
                        "Eşleşme Oranı",
                    ],
                    "Değer": [
                        stats["ref_total"],
                        stats["comp_total"],
                        stats["match_count"],
                        stats["only_ref_count"],
                        stats["only_comp_count"],
                        f"%{stats['match_percentage']}",
                    ],
                }
            )
            stats_df.to_excel(writer, sheet_name="İstatistikler", index=False)
