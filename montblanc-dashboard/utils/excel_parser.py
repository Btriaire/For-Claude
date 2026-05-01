import pandas as pd
from typing import Dict, List

CATEGORICAL_HINTS = [
    "status", "type", "phase", "category", "priority", "team",
    "department", "region", "country", "stage", "state",
    "clinical", "operational", "educational", "statut", "type",
]


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
                n_unique = col_data.nunique()
                sample_values = [str(v) for v in col_data.head(5).tolist()]
                columns.append({
                    "name": str(col),
                    "is_numeric": bool(is_numeric),
                    "n_unique": int(n_unique),
                    "sample": sample_values,
                })
            if columns:
                structure[sheet_name] = {"columns": columns, "rows": len(df)}
        return structure

    def get_suggestions(self) -> List[Dict]:
        """Auto-detect categorical columns and suggest KPIs."""
        suggestions = []
        kpi_id = 0

        for sheet_name, df in self.sheets.items():
            n_rows = len(df)

            # Always suggest Total count
            suggestions.append({
                "id": f"kpi_{kpi_id}",
                "label": "Total Projects",
                "sheet": sheet_name,
                "value_column": None,
                "category_column": None,
                "aggregation": "count",
                "format": "number",
                "color_index": kpi_id,
                "_suggested": True,
            })
            kpi_id += 1

            # Find categorical columns (low cardinality text columns)
            for col in df.columns:
                col_data = df[col].dropna()
                if len(col_data) == 0:
                    continue

                is_numeric = pd.to_numeric(col_data, errors="coerce").notna().sum() / len(col_data) > 0.5
                if is_numeric:
                    continue

                n_unique = col_data.nunique()
                col_lower = col.lower().strip()

                is_hint = any(h in col_lower for h in CATEGORICAL_HINTS)
                is_low_cardinality = 2 <= n_unique <= 15

                if is_hint or is_low_cardinality:
                    label = f"Projects by {col}"
                    suggestions.append({
                        "id": f"kpi_{kpi_id}",
                        "label": label,
                        "sheet": sheet_name,
                        "value_column": None,
                        "category_column": col,
                        "aggregation": "count",
                        "format": "number",
                        "color_index": kpi_id,
                        "_suggested": True,
                    })
                    kpi_id += 1

            # Suggest numeric columns as sum KPIs
            for col in df.columns:
                col_data = df[col].dropna()
                if len(col_data) == 0:
                    continue
                is_numeric = pd.to_numeric(col_data, errors="coerce").notna().sum() / len(col_data) > 0.8
                if is_numeric:
                    suggestions.append({
                        "id": f"kpi_{kpi_id}",
                        "label": col,
                        "sheet": sheet_name,
                        "value_column": col,
                        "category_column": None,
                        "aggregation": "sum",
                        "format": "number",
                        "color_index": kpi_id,
                        "_suggested": True,
                    })
                    kpi_id += 1

            # Only process first sheet with data
            if suggestions:
                break

        return suggestions[:8]

    def extract_kpi_data(self, kpi_configs: List[Dict]) -> List[Dict]:
        results = []
        for kpi in kpi_configs:
            sheet_name   = kpi.get("sheet")
            value_col    = kpi.get("value_column")
            category_col = kpi.get("category_column")
            aggregation  = kpi.get("aggregation", "sum")

            if not sheet_name or sheet_name not in self.sheets:
                continue

            df = self.sheets[sheet_name].copy()

            if aggregation == "count":
                total = float(len(df))
                breakdown: List[Dict] = []
                if category_col and category_col in df.columns:
                    grouped = df[category_col].dropna().value_counts()
                    for label, count in grouped.items():
                        breakdown.append({"label": str(label), "value": float(count)})
            else:
                if not value_col or value_col not in df.columns:
                    continue
                df[value_col] = pd.to_numeric(df[value_col], errors="coerce")
                df = df.dropna(subset=[value_col])
                total = float(df[value_col].sum())
                breakdown = []
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

            # Cap at 7 + Other
            if len(breakdown) > 7:
                top = breakdown[:6]
                other_val = sum(b["value"] for b in breakdown[6:])
                top.append({"label": "Other", "value": other_val})
                breakdown = top

            results.append({
                "id": kpi.get("id", ""),
                "label": kpi.get("label", value_col or "Count"),
                "total": total,
                "format": kpi.get("format", "number"),
                "breakdown": breakdown,
                "color_index": kpi.get("color_index", 0),
            })
        return results
