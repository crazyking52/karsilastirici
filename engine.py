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

    def compare(self, ref_files, ref_column, comp_files, comp_column):
        """
        ref_files: [(dosya_yolu, sayfa_adı), ...] listesi
        ref_column: referans sütun adı
        comp_files: [(dosya_yolu, sayfa_adı), ...] listesi
        comp_column: karşılaştırma sütun adı
        """
        ref_values = set()
        ref_source_map = {}

        for file_path, sheet_name in ref_files:
            ref_df = self.load_file(file_path, sheet_name)
            if ref_column not in ref_df.columns:
                raise ValueError(
                    f"'{ref_column}' sütunu '{Path(file_path).name}' dosyasında bulunamadı.\n"
                    f"Mevcut sütunlar: {', '.join(ref_df.columns.astype(str))}"
                )
            values = set(ref_df[ref_column].dropna().astype(str).str.strip())
            ref_values.update(values)
            for v in values:
                if v not in ref_source_map:
                    ref_source_map[v] = []
                ref_source_map[v].append(Path(file_path).name)

        all_comp_values = set()
        source_map = {}

        for file_path, sheet_name in comp_files:
            comp_df = self.load_file(file_path, sheet_name)
            if comp_column not in comp_df.columns:
                raise ValueError(
                    f"'{comp_column}' sütunu '{Path(file_path).name}' dosyasında bulunamadı.\n"
                    f"Mevcut sütunlar: {', '.join(comp_df.columns.astype(str))}"
                )
            values = set(comp_df[comp_column].dropna().astype(str).str.strip())
            all_comp_values.update(values)
            for v in values:
                if v not in source_map:
                    source_map[v] = []
                source_map[v].append(Path(file_path).name)

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
            "stats": {
                "ref_total": total,
                "comp_total": len(all_comp_values),
                "match_count": len(matches),
                "only_ref_count": len(only_in_ref),
                "only_comp_count": len(only_in_comp),
                "match_percentage": match_pct,
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
