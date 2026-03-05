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
                on_progress=None, ignore_case=False, strip_spaces=False):
        total_rows = self._count_total_rows(ref_files) + self._count_total_rows(comp_files)
        checked = 0

        def _normalize(val):
            if strip_spaces:
                val = " ".join(val.split())
            if ignore_case:
                val = val.replace("İ", "i").replace("ı", "i")
                val = val.lower()
            return val

        ref_values = set()
        ref_display = {}
        ref_source_map = {}
        ref_duplicates = {}

        for file_path, sheet_name in ref_files:
            ref_df = self.load_file(file_path, sheet_name)
            fname = Path(file_path).name
            present_cols = [c for c in ref_columns if c in ref_df.columns]
            if not present_cols:
                row_count = len(ref_df)
                checked += row_count
                if on_progress:
                    on_progress(checked, total_rows, "Referans okunuyor...")
                continue
            row_count = len(ref_df)
            for col in present_cols:
                col_values = ref_df[col].dropna().astype(str).str.strip()
                counts = col_values.value_counts()
                for val, cnt in counts.items():
                    if cnt > 1:
                        key = _normalize(val)
                        if key not in ref_duplicates:
                            ref_duplicates[key] = []
                        ref_duplicates[key].append(f"{fname} → {col}: {cnt} kez")
                for raw in col_values:
                    norm = _normalize(raw)
                    ref_values.add(norm)
                    if norm not in ref_display:
                        ref_display[norm] = raw
                    if norm not in ref_source_map:
                        ref_source_map[norm] = []
                    source = f"{fname} → {col}"
                    if source not in ref_source_map[norm]:
                        ref_source_map[norm].append(source)
            checked += row_count
            if on_progress:
                on_progress(checked, total_rows, "Referans okunuyor...")

        all_comp_values = set()
        comp_display = {}
        source_map = {}
        comp_duplicates = {}

        for file_path, sheet_name in comp_files:
            comp_df = self.load_file(file_path, sheet_name)
            fname = Path(file_path).name
            present_cols = [c for c in comp_columns if c in comp_df.columns]
            if not present_cols:
                row_count = len(comp_df)
                checked += row_count
                if on_progress:
                    on_progress(checked, total_rows, "Karşılaştırılıyor...")
                continue
            row_count = len(comp_df)
            for col in present_cols:
                col_values = comp_df[col].dropna().astype(str).str.strip()
                counts = col_values.value_counts()
                for val, cnt in counts.items():
                    if cnt > 1:
                        key = _normalize(val)
                        if key not in comp_duplicates:
                            comp_duplicates[key] = []
                        comp_duplicates[key].append(f"{fname} → {col}: {cnt} kez")
                for raw in col_values:
                    norm = _normalize(raw)
                    all_comp_values.add(norm)
                    if norm not in comp_display:
                        comp_display[norm] = raw
                    if norm not in source_map:
                        source_map[norm] = []
                    source = f"{fname} → {col}"
                    if source not in source_map[norm]:
                        source_map[norm].append(source)
            checked += row_count
            if on_progress:
                on_progress(checked, total_rows, "Karşılaştırılıyor...")

        match_keys = ref_values & all_comp_values
        only_ref_keys = ref_values - all_comp_values
        only_comp_keys = all_comp_values - ref_values

        display = {**ref_display, **comp_display}
        matches = sorted(display.get(k, k) for k in match_keys)
        only_in_ref = sorted(display.get(k, k) for k in only_ref_keys)
        only_in_comp = sorted(display.get(k, k) for k in only_comp_keys)

        total = len(ref_values)
        match_pct = round(len(match_keys) / total * 100, 1) if total > 0 else 0

        return {
            "matches": matches,
            "only_in_reference": only_in_ref,
            "only_in_comparison": only_in_comp,
            "ref_source_map": ref_source_map,
            "source_map": source_map,
            "ref_duplicates": ref_duplicates,
            "comp_duplicates": comp_duplicates,
            "stats": {
                "ref_total": total,
                "comp_total": len(all_comp_values),
                "match_count": len(match_keys),
                "only_ref_count": len(only_ref_keys),
                "only_comp_count": len(only_comp_keys),
                "match_percentage": match_pct,
                "ref_dup_count": len(ref_duplicates),
                "comp_dup_count": len(comp_duplicates),
            },
        }

    def detailed_compare(self, ref_files, comp_files, key_columns,
                          detail_columns, on_progress=None,
                          ignore_case=False, strip_spaces=False):
        def _norm(val):
            val = str(val).strip()
            if val == "nan":
                return ""
            if strip_spaces:
                val = " ".join(val.split())
            if ignore_case:
                val = val.replace("İ", "i").replace("ı", "i").lower()
            return val

        ref_rows = {}
        total_rows = (self._count_total_rows(ref_files)
                      + self._count_total_rows(comp_files))
        checked = 0

        for file_path, sheet_name in ref_files:
            df = self.load_file(file_path, sheet_name)
            fname = Path(file_path).name
            for _, row in df.iterrows():
                key = tuple(_norm(row.get(c, "")) for c in key_columns)
                if any(k == "" for k in key):
                    continue
                detail = {}
                for c in detail_columns:
                    raw = str(row.get(c, ""))
                    detail[c] = raw if raw != "nan" else ""
                if key not in ref_rows:
                    ref_rows[key] = {"detail": detail, "file": fname}
            checked += len(df)
            if on_progress:
                on_progress(checked, total_rows, "Referans okunuyor...")

        differences = []
        matched_keys = set()
        only_comp_keys = []

        for file_path, sheet_name in comp_files:
            df = self.load_file(file_path, sheet_name)
            fname = Path(file_path).name
            for _, row in df.iterrows():
                key = tuple(_norm(row.get(c, "")) for c in key_columns)
                if any(k == "" for k in key):
                    continue

                if key in ref_rows:
                    matched_keys.add(key)
                    ref_detail = ref_rows[key]["detail"]
                    diffs = []
                    for c in detail_columns:
                        raw_comp = str(row.get(c, ""))
                        raw_comp = raw_comp if raw_comp != "nan" else ""
                        raw_ref = ref_detail.get(c, "")
                        norm_ref = _norm(raw_ref)
                        norm_comp = _norm(raw_comp)
                        if norm_ref != norm_comp:
                            diffs.append({
                                "column": c,
                                "ref_value": raw_ref,
                                "comp_value": raw_comp,
                            })
                    if diffs:
                        key_display = dict(zip(key_columns, key))
                        differences.append({
                            "key": key_display,
                            "ref_file": ref_rows[key]["file"],
                            "comp_file": fname,
                            "diffs": diffs,
                        })
                else:
                    key_display = dict(zip(key_columns, key))
                    only_comp_keys.append({"key": key_display, "file": fname})

            checked += len(df)
            if on_progress:
                on_progress(checked, total_rows, "Karşılaştırılıyor...")

        only_ref_keys = []
        for key, info in ref_rows.items():
            if key not in matched_keys:
                key_display = dict(zip(key_columns, key))
                only_ref_keys.append({"key": key_display, "file": info["file"]})

        return {
            "differences": differences,
            "only_in_ref": only_ref_keys,
            "only_in_comp": only_comp_keys,
            "stats": {
                "total_matched": len(matched_keys),
                "total_with_diffs": len(differences),
                "total_identical": len(matched_keys) - len(differences),
                "only_ref_count": len(only_ref_keys),
                "only_comp_count": len(only_comp_keys),
            },
        }

    def _build_dataframes(self, results):
        frames = {}

        if results["matches"]:
            frames["matches"] = pd.DataFrame({"Kayıt": results["matches"]})

        no_match_rows = []
        for v in results["only_in_reference"]:
            ref_src = ", ".join(results["ref_source_map"].get(v, ["?"]))
            no_match_rows.append({"Kayıt": v, "Durum": f"Referansta var ({ref_src}), karşılaştırmada yok"})
        for v in results["only_in_comparison"]:
            sources = ", ".join(results["source_map"].get(v, ["?"]))
            no_match_rows.append({"Kayıt": v, "Durum": f"Karşılaştırmada var ({sources}), referansta yok"})
        if no_match_rows:
            frames["no_match"] = pd.DataFrame(no_match_rows)

        if results["only_in_reference"]:
            frames["only_ref"] = pd.DataFrame(
                {"Kayıt": results["only_in_reference"]}
            )

        if results["only_in_comparison"]:
            frames["only_comp"] = pd.DataFrame({
                "Kayıt": results["only_in_comparison"],
                "Kaynak Dosya": [
                    ", ".join(results["source_map"].get(v, ["?"]))
                    for v in results["only_in_comparison"]
                ],
            })

        dup_rows = []
        for val, details in results["ref_duplicates"].items():
            for d in details:
                dup_rows.append({"Kayıt": val, "Konum": d, "Taraf": "Referans"})
        for val, details in results["comp_duplicates"].items():
            for d in details:
                dup_rows.append({"Kayıt": val, "Konum": d, "Taraf": "Karşılaştırma"})
        if dup_rows:
            frames["duplicates"] = pd.DataFrame(dup_rows)

        stats = results["stats"]
        frames["stats"] = pd.DataFrame({
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
        })

        return frames

    def export_results(self, results, output_path, csv_encoding="utf-8-sig"):
        suffix = Path(output_path).suffix.lower()
        frames = self._build_dataframes(results)

        if suffix == ".csv":
            self._export_csv(frames, output_path, encoding=csv_encoding)
        else:
            self._export_excel(frames, output_path)

    def _export_excel(self, frames, output_path):
        sheet_map = {
            "matches": "Eşleşenler",
            "no_match": "Eşleşmeyenler",
            "only_ref": "Yalnızca Referans",
            "only_comp": "Yalnızca Karşılaştırma",
            "duplicates": "Tekrar Edenler",
            "stats": "İstatistikler",
        }
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for key, sheet_name in sheet_map.items():
                if key in frames:
                    frames[key].to_excel(writer, sheet_name=sheet_name, index=False)

    def _export_csv(self, frames, output_path, encoding="utf-8-sig"):
        base = Path(output_path)
        stem = base.stem
        parent = base.parent

        name_map = {
            "matches": "eslesen",
            "no_match": "eslesmeyenler",
            "only_ref": "yalnizca_referans",
            "only_comp": "yalnizca_karsilastirma",
            "duplicates": "tekrar_edenler",
            "stats": "istatistikler",
        }

        for key, suffix in name_map.items():
            if key in frames:
                csv_path = parent / f"{stem}_{suffix}.csv"
                frames[key].to_csv(csv_path, index=False, encoding=encoding)
