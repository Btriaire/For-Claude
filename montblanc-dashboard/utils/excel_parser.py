import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional


class ExcelParser:
    def __init__(self, filepath: str):
        self.filepath = filepath
        self.workbook = pd.ExcelFile(filepath)
        self.sheets: Dict[str, pd.DataFrame] = {}
        self._load_sheets()

    def _load_sheets(self):
        for sheet_name in self.workbook.sheet_names:
            try:
                df = pd.read_excel(self.filepath, sheet_name=sheet_name)
                df = df.dropna(how="all")
                df.columns = [str(c).strip() for c in df.columns]
                self.sheets[sheet_name] = df
            except Exception:
                pass

    def get_structure(self) -> Dict:
        structure = {}
        for sheet_name, df in self.sheets.items():
            columns = []
            for col in df.columns:
                col_data = df[col].dropna()
                if len(col_data) == 0:
                    continue
                numeric_ratio = pd.to_numeric(col_data, errors="coerce").notna().sum() / len(col_data)
                is_numeric = numeric_ratio > 0.5
                sample_values = [str(v) for v in col_data.head(5).tolist()]
                columns.append({
                    "name": str(col),
                    "is_numeric": bool(is_numeric),
                    "sample": sample_values,
                })
            if columns:
                structure[sheet_name] = {"columns": columns, "rows": len(df)}
        return structure

    def extract_kpi_data(self, kpi_configs: List[Dict]) -> List[Dict]:
        results = []
        for kpi in kpi_configs:
            sheet_name = kpi.get("sheet")
            value_col = kpi.get("value_column")
            category_col = kpi.get("category_column")

            if not sheet_name or not value_col:
                continue
            if sheet_name not in self.sheets:
                continue

            df = self.sheets[sheet_name].copy()
            if value_col not in df.columns:
                continue

            df[value_col] = pd.to_numeric(df[value_col], errors="coerce")
            df = df.dropna(subset=[value_col])
            total = float(df[value_col].sum())

            breakdown: List[Dict] = []
            if category_col and category_col in df.columns:
                grouped = (
                    df.groupby(category_col, dropna=True)[value_col]
                    .sum()
                    .reset_index()
                    .sort_values(value_col, ascending=False)
                )
                for _, row in grouped.iterrows():
                    val = float(row[value_col])
                    if val > 0 and pd.notna(row[category_col]):
                        breakdown.append({"label": str(row[category_col]), "value": val})

                # Merge tail into "Other" if more than 7 segments
                if len(breakdown) > 7:
                    top = breakdown[:6]
                    other_val = sum(b["value"] for b in breakdown[6:])
                    top.append({"label": "Other", "value": other_val})
                    breakdown = top

            results.append({
                "id": kpi.get("id", ""),
                "label": kpi.get("label", value_col),
                "total": total,
                "format": kpi.get("format", "number"),
                "breakdown": breakdown,
                "target": kpi.get("target"),
                "color_index": kpi.get("color_index", 0),
            })
        return results
