"""Import and auto-analyze price data from Excel files (.xls/.xlsx).

Handles different table formats across cities by auto-detecting:
- Header row (looks for known column names)
- Price columns (含规费价格, 含税价, 除税价, etc.)
- Labor vs material rows (by unit = 工日)
- Price ranges (e.g., '219.00-266.00') and single values
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from .models import Meta, MonthlyPrice, PriceData, PriceItem, parse_price_range

# Known header keywords for auto-detection
HEADER_MARKERS = {"名称", "编码", "序号", "计量单位"}
PRICE_COL_PRIORITY = [
    "含规费价格",
    "含税价",
    "除税价",
    "不含规费价格",
    "单价",
    "价格",
]


@dataclass
class TableStructure:
    """Detected table structure."""
    header_row: int
    headers: list[str]
    name_col: str
    price_col: str
    code_col: str | None
    unit_col: str | None
    spec_col: str | None


def analyze_table(df: pd.DataFrame) -> TableStructure:
    """Auto-detect table structure by scanning for header patterns.

    Returns a TableStructure describing the detected layout.
    Raises ValueError if structure cannot be determined.
    """
    # Find header row
    header_row = None
    for i in range(min(15, len(df))):
        row_values = {str(v).strip() for v in df.iloc[i].values if pd.notna(v)}
        if len(row_values & HEADER_MARKERS) >= 2:
            header_row = i
            break

    if header_row is None:
        raise ValueError("无法自动识别表头行（未找到'名称'、'编码'等列标记）")

    headers = [str(v).strip() for v in df.iloc[header_row].values]

    # Find name column
    name_col = None
    for h in headers:
        if h == "名称":
            name_col = h
            break
    if name_col is None:
        raise ValueError(f"表头中缺少'名称'列。检测到的列: {headers}")

    # Find price column (by priority)
    price_col = None
    for keyword in PRICE_COL_PRIORITY:
        for h in headers:
            if keyword in h:
                price_col = h
                break
        if price_col:
            break
    if price_col is None:
        raise ValueError(f"无法找到价格列。检测到的列: {headers}")

    # Optional columns
    code_col = next((h for h in headers if "编码" in h), None)
    unit_col = next((h for h in headers if "计量单位" in h or h == "单位"), None)
    spec_col = next((h for h in headers if h == "规格"), None)

    return TableStructure(
        header_row=header_row,
        headers=headers,
        name_col=name_col,
        price_col=price_col,
        code_col=code_col,
        unit_col=unit_col,
        spec_col=spec_col,
    )


def extract_month_from_title(df: pd.DataFrame) -> str | None:
    """Extract month from title row (e.g., '2026年3月信息价' -> '2026-03')."""
    for i in range(min(5, len(df))):
        for val in df.iloc[i].values:
            if pd.isna(val):
                continue
            match = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月", str(val))
            if match:
                year, month = int(match.group(1)), int(match.group(2))
                return f"{year}-{month:02d}"
    return None


def analyze_excel(excel_path: str | Path) -> dict:
    """Analyze an Excel file and return a summary of its structure.

    Returns a dict describing all sheets, their table structures,
    detected labor types, material types, and price columns.
    This is used by the AI agent to understand what data is available.
    """
    excel_path = Path(excel_path)
    engine = "xlrd" if excel_path.suffix.lower() == ".xls" else "openpyxl"
    xls = pd.ExcelFile(excel_path, engine=engine)

    result = {
        "file": str(excel_path),
        "sheets": [],
    }

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        month = extract_month_from_title(df)

        sheet_info = {
            "name": sheet_name,
            "rows": len(df),
            "month": month,
            "labor_types": [],
            "material_types": [],
            "structure": None,
        }

        try:
            structure = analyze_table(df)
            sheet_info["structure"] = {
                "header_row": structure.header_row,
                "columns": structure.headers,
                "price_column": structure.price_col,
            }

            # Extract data rows
            df_data = df.iloc[structure.header_row + 1:].copy()
            df_data.columns = structure.headers

            for _, row in df_data.iterrows():
                name = str(row.get(structure.name_col, "")).strip()
                if not name or name in ("nan", "NaN", "None", ""):
                    continue

                unit = str(row.get(structure.unit_col, "")).strip() if structure.unit_col else ""
                price_str = str(row.get(structure.price_col, "")).strip()

                if price_str in ("", "nan", "NaN", "None", "-"):
                    continue

                entry = {"name": name, "unit": unit, "price": price_str}
                if structure.code_col:
                    entry["code"] = str(row.get(structure.code_col, "")).strip()
                if structure.spec_col:
                    spec = str(row.get(structure.spec_col, "")).strip()
                    if spec and spec not in ("nan", "NaN", "None"):
                        entry["spec"] = spec

                # Also collect all available price columns for this row
                for h in structure.headers:
                    if h == structure.price_col:
                        continue
                    for kw in PRICE_COL_PRIORITY:
                        if kw in h:
                            alt_price = str(row.get(h, "")).strip()
                            if alt_price and alt_price not in ("", "nan", "NaN", "None", "-"):
                                entry[h] = alt_price
                            break

                if unit == "工日":
                    sheet_info["labor_types"].append(entry)
                else:
                    sheet_info["material_types"].append(entry)

        except ValueError as e:
            sheet_info["error"] = str(e)

        result["sheets"].append(sheet_info)

    return result


def import_excel(
    excel_path: str | Path,
    city: str,
    month: str | None = None,
    sheet_name: str | None = None,
    types: list[str] | None = None,
    category: str | None = None,
    price_col: str | None = None,
    source: str = "",
) -> PriceData:
    """Import price data from a single Excel file.

    Auto-analyzes the table structure. If `types` is None, extracts ALL entries.
    If `category` is "labor", only extracts rows with unit="工日".

    Args:
        excel_path: Path to .xls or .xlsx file.
        city: City name (e.g., "广州").
        month: Month string. Auto-detected from title if None.
        sheet_name: Sheet to read. None = first sheet.
        types: Optional list of specific type names to filter.
        category: Optional "labor" or "material" filter.
        price_col: Override auto-detected price column.
        source: Data source description.

    Returns:
        PriceData with extracted prices.
    """
    excel_path = Path(excel_path)
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel 文件不存在: {excel_path}")

    engine = "xlrd" if excel_path.suffix.lower() == ".xls" else "openpyxl"
    xls = pd.ExcelFile(excel_path, engine=engine)

    # --sheet all: import all sheets and merge
    if sheet_name and sheet_name.lower() == "all":
        merged: PriceData | None = None
        for sn in xls.sheet_names:
            print(f"\n[import] Sheet: {sn}")
            try:
                part = import_excel(
                    excel_path=excel_path,
                    city=city,
                    month=month,
                    sheet_name=sn,
                    types=types,
                    category=category,
                    price_col=price_col,
                    source=source,
                )
                # Tag each item with its source sheet
                for item in part.prices:
                    item.sheet = sn
                if merged is None:
                    merged = part
                else:
                    merged.merge(part)
            except ValueError as e:
                print(f"  [skip] {sn}: {e}")
        if merged is None:
            raise ValueError("所有 Sheet 均无法解析")
        return merged

    if sheet_name is None:
        sheet_name = xls.sheet_names[0]
    elif sheet_name not in xls.sheet_names:
        available = ", ".join(xls.sheet_names)
        raise ValueError(f"Sheet '{sheet_name}' 不存在。可用: {available}")

    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

    # Auto-detect month
    if month is None:
        month = extract_month_from_title(df)
        if month is None:
            raise ValueError("无法从文件中识别月份，请通过 --period 指定")

    # Auto-analyze table structure
    structure = analyze_table(df)
    print(f"  [分析] Sheet '{sheet_name}': 表头行={structure.header_row}, 价格列='{structure.price_col}'")

    # Allow override of price column
    if price_col:
        matching = [h for h in structure.headers if price_col in h]
        if matching:
            structure.price_col = matching[0]
        else:
            print(f"  [warning] 未找到匹配 '{price_col}' 的列，使用自动检测的 '{structure.price_col}'")

    # Set up data frame with headers
    df_data = df.iloc[structure.header_row + 1:].copy()
    df_data.columns = structure.headers

    # Build PriceData
    meta = Meta(city=city, period=month, source=source)
    price_data = PriceData(meta=meta)
    types_set = {t.strip() for t in types} if types else None

    extracted_count = 0
    skipped_count = 0

    for _, row in df_data.iterrows():
        name = str(row.get(structure.name_col, "")).strip()
        if not name or name in ("nan", "NaN", "None", ""):
            continue

        unit = str(row.get(structure.unit_col, "")).strip() if structure.unit_col else ""
        row_category = "labor" if unit == "工日" else "material"

        # Filter by category
        if category and row_category != category:
            continue

        # Filter by type names
        if types_set and name not in types_set:
            continue

        # Parse price — try primary column first, then fallback to others
        price_str = str(row.get(structure.price_col, "")).strip()
        parsed = False
        p_min = p_max = 0.0

        if price_str and price_str not in ("", "nan", "NaN", "None", "-"):
            try:
                p_min, p_max = parse_price_range(price_str)
                parsed = True
            except ValueError:
                pass

        # Fallback: try other price columns in priority order
        if not parsed:
            for kw in PRICE_COL_PRIORITY:
                for h in structure.headers:
                    if kw in h and h != structure.price_col:
                        alt_str = str(row.get(h, "")).strip()
                        if alt_str and alt_str not in ("", "nan", "NaN", "None", "-"):
                            try:
                                p_min, p_max = parse_price_range(alt_str)
                                parsed = True
                                break
                            except ValueError:
                                pass
                if parsed:
                    break

        if not parsed:
            skipped_count += 1
            continue

        code = str(row.get(structure.code_col, "")).strip() if structure.code_col else ""
        spec = str(row.get(structure.spec_col, "")).strip() if structure.spec_col else ""
        if spec and spec not in ("nan", "NaN"):
            display_name = f"{name} {spec}" if spec else name
        else:
            display_name = name

        item = price_data.find_item(display_name)
        if item is None:
            item = PriceItem(name=display_name, code=code, category=row_category, unit=unit, sheet=sheet_name)
            price_data.prices.append(item)

        item.add_month(MonthlyPrice.from_range(month, p_min, p_max))
        extracted_count += 1

    print(f"  [结果] 提取 {extracted_count} 条记录, 跳过 {skipped_count} 条")

    if not price_data.prices:
        filter_desc = []
        if types_set:
            filter_desc.append(f"工种={types_set}")
        if category:
            filter_desc.append(f"类别={category}")
        hint = f" (过滤条件: {', '.join(filter_desc)})" if filter_desc else ""
        print(f"  [warning] 未提取到任何数据{hint}")

    return price_data


def import_multiple_excels(
    excel_paths: list[str | Path],
    city: str,
    sheet_name: str | None = None,
    types: list[str] | None = None,
    category: str | None = None,
    price_col: str | None = None,
    source: str = "",
) -> PriceData:
    """Import and merge price data from multiple Excel files (one per month).

    Returns merged PriceData with all months.
    """
    if not excel_paths:
        raise ValueError("未提供 Excel 文件路径")

    merged: PriceData | None = None

    for path in sorted(excel_paths, key=lambda p: str(p)):
        print(f"[import] 正在处理: {Path(path).name}")
        data = import_excel(
            excel_path=path,
            city=city,
            sheet_name=sheet_name,
            types=types,
            category=category,
            price_col=price_col,
            source=source,
        )
        if merged is None:
            merged = data
        else:
            merged.merge(data)

    if merged is None:
        raise ValueError("没有成功导入任何数据")

    # Update period to cover all months
    all_months = sorted({mp.month for item in merged.prices for mp in item.monthly})
    if len(all_months) > 1:
        merged.meta.period = f"{all_months[0]}~{all_months[-1]}"
    elif all_months:
        merged.meta.period = all_months[0]

    return merged
