"""CLI entry point for the engineering price collector skill.

Sub-commands:
    analyze  - Analyze Excel file structure (auto-detect format)
    import   - Import price data from Excel file(s)
    scrape   - Scrape price data from official websites
    chart    - Generate price trend charts from collected data
    query    - Query stored price data
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

DEFAULT_DATA_DIR = (
    "/data/workspace/skill-data"
    if Path("/data/workspace/skill-data").exists()
    else str(Path(__file__).resolve().parents[2] / ".." / ".." / "skill-data")
)


def parse_period(period: str) -> tuple[int, list[int]]:
    """Parse period string into (year, [months]).

    Supported formats:
        "2025-03"           -> (2025, [3])
        "2025-01~2025-12"   -> (2025, [1,2,...,12])
        "2025-06~2025-09"   -> (2025, [6,7,8,9])
    """
    # Range format
    match = re.match(r"(\d{4})-(\d{1,2})~(\d{4})-(\d{1,2})", period)
    if match:
        y1, m1, y2, m2 = int(match.group(1)), int(match.group(2)), int(match.group(3)), int(match.group(4))
        if y1 != y2:
            raise ValueError(f"跨年区间暂不支持: {period}")
        return y1, list(range(m1, m2 + 1))

    # Single month format
    match = re.match(r"(\d{4})-(\d{1,2})$", period)
    if match:
        return int(match.group(1)), [int(match.group(2))]

    raise ValueError(f"无法解析时间段: {period}。格式示例: 2025-03 或 2025-01~2025-12")


def cmd_analyze(args: argparse.Namespace) -> None:
    """Analyze Excel file structure."""
    from .excel_import import analyze_excel

    result = analyze_excel(args.excel)

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        total_labor = 0
        total_material = 0
        print(f"\n文件: {result['file']}")
        for sheet in result["sheets"]:
            print(f"\n{'='*70}")
            print(f"Sheet: {sheet['name']} ({sheet['rows']} 行)")
            if sheet.get("month"):
                print(f"月份: {sheet['month']}")
            if sheet.get("error"):
                print(f"错误: {sheet['error']}")
                continue
            if sheet.get("structure"):
                cols = sheet["structure"]["columns"]
                print(f"列结构: {' | '.join(cols)}")
                print(f"主价格列: {sheet['structure']['price_column']}")

            labor = sheet.get("labor_types", [])
            material = sheet.get("material_types", [])
            total_labor += len(labor)
            total_material += len(material)

            if labor:
                print(f"\n  ── 人工类型 ({len(labor)} 个) ──")
                for i, entry in enumerate(labor, 1):
                    code = f"  [{entry.get('code','')}]" if entry.get("code") else ""
                    alt_prices = ""
                    for k, v in entry.items():
                        if k not in ("name", "unit", "price", "code", "spec"):
                            alt_prices += f"  {k}={v}"
                    print(f"  {i:>4}. {entry['name']}{code}: {entry['price']} 元/{entry['unit']}{alt_prices}")

            if material:
                print(f"\n  ── 材料类型 ({len(material)} 个) ──")
                for i, entry in enumerate(material, 1):
                    code = f"  [{entry.get('code','')}]" if entry.get("code") else ""
                    spec = f" ({entry.get('spec','')})" if entry.get("spec") else ""
                    unit = entry.get("unit", "")
                    alt_prices = ""
                    for k, v in entry.items():
                        if k not in ("name", "unit", "price", "code", "spec"):
                            alt_prices += f"  {k}={v}"
                    print(f"  {i:>4}. {entry['name']}{spec}{code} [{unit}]: {entry['price']}{alt_prices}")

        print(f"\n{'='*70}")
        print(f"汇总: {len(result['sheets'])} 个 Sheet, 人工 {total_labor} 条, 材料 {total_material} 条, 合计 {total_labor + total_material} 条")


def cmd_import(args: argparse.Namespace) -> None:
    """Import data from Excel file(s)."""
    from .excel_import import import_excel, import_multiple_excels
    from .models import get_data_dir

    types = [t.strip() for t in args.types.split(",")] if args.types else None
    output_dir = get_data_dir(args.output, args.city)

    excel_paths = args.excel if isinstance(args.excel, list) else [args.excel]

    if len(excel_paths) == 1:
        data = import_excel(
            excel_path=excel_paths[0],
            city=args.city,
            month=args.period if args.period and "~" not in args.period else None,
            sheet_name=args.sheet,
            types=types,
            category=args.category,
            price_col=args.price_col,
        )
    else:
        data = import_multiple_excels(
            excel_paths=excel_paths,
            city=args.city,
            sheet_name=args.sheet,
            types=types,
            category=args.category,
            price_col=args.price_col,
        )

    path = data.save(output_dir)
    print(f"\n[完成] 数据已保存到: {path}")
    print(json.dumps(data.to_dict(), ensure_ascii=False, indent=2))


def cmd_scrape(args: argparse.Namespace) -> None:
    """Scrape data from official websites."""
    from .models import get_data_dir
    from .scraper import scrape_prices

    year, months = parse_period(args.period)
    types = [t.strip() for t in args.types.split(",")] if args.types else None
    output_dir = get_data_dir(args.output, args.city)

    data = scrape_prices(
        city=args.city,
        year=year,
        months=months,
        types=types,
        category=args.category,
    )

    path = data.save(output_dir)
    print(f"\n[完成] 数据已保存到: {path}")
    print(json.dumps(data.to_dict(), ensure_ascii=False, indent=2))


def cmd_chart(args: argparse.Namespace) -> None:
    """Generate charts from stored data."""
    from .models import PriceData, get_data_dir
    from .visualization import export_xlsx, generate_charts_by_sheet

    data_dir = get_data_dir(args.output, args.city)
    json_path = data_dir / "prices.json"

    if not json_path.exists():
        print(f"[error] 未找到数据文件: {json_path}")
        print("请先使用 import 或 scrape 命令采集数据")
        sys.exit(1)

    data = PriceData.load(json_path)
    types = [t.strip() for t in args.types.split(",")] if args.types else None

    # Generate charts grouped by sheet
    print(f"\n[chart] 按 Sheet 分组生成图表...")
    paths = generate_charts_by_sheet(data, data_dir, types=types)
    print(f"\n[完成] 共生成 {len(paths)} 张图表")

    # Export xlsx grouped by sheet
    xlsx_path = export_xlsx(data, data_dir, types=types)
    print(f"[完成] Excel 已生成: {xlsx_path}")


def cmd_query(args: argparse.Namespace) -> None:
    """Query stored data."""
    from .models import PriceData, get_data_dir

    data_dir = get_data_dir(args.output, args.city)
    json_path = data_dir / "prices.json"

    if not json_path.exists():
        print(f"[error] 未找到数据文件: {json_path}")
        sys.exit(1)

    data = PriceData.load(json_path)
    types = {t.strip() for t in args.types.split(",")} if args.types else None

    if args.json:
        filtered = data.to_dict()
        if types:
            filtered["prices"] = [p for p in filtered["prices"] if p["name"] in types]
        print(json.dumps(filtered, ensure_ascii=False, indent=2))
    else:
        print(f"\n城市: {data.meta.city}")
        print(f"时段: {data.meta.period}")
        print(f"来源: {data.meta.source}")
        print(f"采集时间: {data.meta.collected_at}")
        print(f"\n{'名称':<15} {'类型':<8} {'单位':<6} {'月份数':<6} {'最新价格'}")
        print("-" * 60)
        for item in data.prices:
            if types and item.name not in types:
                continue
            cat = "人工" if item.category == "labor" else "材料"
            n_months = len(item.monthly)
            latest = f"{item.monthly[-1].price_avg:.2f}" if item.monthly else "-"
            print(f"{item.name:<15} {cat:<8} {item.unit:<6} {n_months:<6} {latest}")


def cmd_run(args: argparse.Namespace) -> None:
    """One-shot workflow: [scrape →] parse → charts + xlsx → report."""
    from .excel_import import import_excel, import_multiple_excels
    from .models import get_data_dir
    from .visualization import export_xlsx, generate_charts_by_sheet

    types = [t.strip() for t in args.types.split(",")] if args.types else None
    output_dir = get_data_dir(args.output, args.city)
    sheet = args.sheet or "all"

    if args.excel:
        # ── 有 Excel：直接分析 ──
        excel_paths = args.excel
        if len(excel_paths) == 1:
            data = import_excel(
                excel_path=excel_paths[0],
                city=args.city,
                month=args.period if args.period and "~" not in args.period else None,
                sheet_name=sheet,
                types=types,
                category=args.category,
            )
        else:
            data = import_multiple_excels(
                excel_paths=excel_paths,
                city=args.city,
                sheet_name=sheet,
                types=types,
                category=args.category,
            )
    else:
        # ── 无 Excel：自动爬取 ──
        from .scraper import scrape_prices
        year, months = parse_period(args.period)
        data = scrape_prices(
            city=args.city,
            year=year,
            months=months,
            types=types,
            category=args.category,
            data_dir=args.output,
        )

    json_path = data.save(output_dir)

    # ── Step 2: 生成图表 + Excel ──
    chart_paths = generate_charts_by_sheet(data, output_dir, types=types)
    xlsx_path = export_xlsx(data, output_dir, types=types)

    # ── Step 3: 报告 ──
    from collections import Counter
    sheet_counts = Counter(it.sheet for it in data.prices)
    labor = [it for it in data.prices if it.category == "labor"]
    material = [it for it in data.prices if it.category == "material"]

    print(f"\n{'='*60}")
    print(f"  工程信息价采集报告")
    print(f"{'='*60}")
    print(f"  城市: {data.meta.city}")
    print(f"  时段: {data.meta.period}")
    print(f"  采集时间: {data.meta.collected_at}")
    source_desc = f"{len(args.excel)} 个文件" if args.excel else f"自动抓取 ({data.meta.source})"
    print(f"  数据来源: {source_desc}")
    print(f"")
    print(f"  Sheet 明细 ({len(sheet_counts)} 个):")
    for sn, cnt in sheet_counts.most_common():
        print(f"    - {sn}: {cnt} 条")
    print(f"")
    print(f"  人工: {len(labor)} 个工种")
    for it in labor:
        latest = it.monthly[-1] if it.monthly else None
        price_str = f"{latest.price_avg:.0f} 元/工日 ({latest.price_min:.0f}~{latest.price_max:.0f})" if latest else "-"
        print(f"    - {it.name}: {price_str}")
    print(f"")
    print(f"  材料: {len(material)} 种")
    print(f"")
    print(f"  输出文件:")
    print(f"    JSON:  {json_path}")
    print(f"    Excel: {xlsx_path}")
    print(f"    图表:  {len(chart_paths)} 张")
    for p in chart_paths:
        print(f"      {p}")
    print(f"{'='*60}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="工程信息价采集工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # === run (one-shot workflow) ===
    p_run = subparsers.add_parser("run", help="抓取/分析 → 图表 → 报告")
    p_run.add_argument("--city", required=True, help="城市名称")
    p_run.add_argument("--period", required=True, help="时间段 (如: 2026-03 或 2025-01~2025-12)")
    p_run.add_argument("--excel", nargs="+", help="已有 Excel 文件路径（不传则自动爬取）")
    p_run.add_argument("--types", help="逗号分隔的工种/材料名称（可选）")
    p_run.add_argument("--category", choices=["labor", "material"], help="按类别过滤")
    p_run.add_argument("--sheet", help="Sheet 名称（默认 all）")
    p_run.add_argument("--output", default=DEFAULT_DATA_DIR, help="输出目录")

    # === analyze ===
    p_analyze = subparsers.add_parser("analyze", help="分析 Excel 文件结构")
    p_analyze.add_argument("--excel", required=True, help="Excel 文件路径")
    p_analyze.add_argument("--json", action="store_true", help="输出 JSON 格式")

    # === import ===
    p_import = subparsers.add_parser("import", help="从 Excel 文件导入数据")
    p_import.add_argument("--city", required=True, help="城市名称")
    p_import.add_argument("--excel", required=True, nargs="+", help="Excel 文件路径（支持多个）")
    p_import.add_argument("--period", help="月份 (如: 2026-03)，默认从文件自动识别")
    p_import.add_argument("--types", help="逗号分隔的工种/材料名称（可选，不填则提取全部）")
    p_import.add_argument("--category", choices=["labor", "material"], help="按类别过滤")
    p_import.add_argument("--sheet", help="Sheet 名称（默认第一个 Sheet）")
    p_import.add_argument("--price-col", help="价格列名称（默认自动检测）")
    p_import.add_argument("--output", default=DEFAULT_DATA_DIR, help="输出目录")

    # === scrape ===
    p_scrape = subparsers.add_parser("scrape", help="从造价信息网爬取数据")
    p_scrape.add_argument("--city", required=True, help="城市名称")
    p_scrape.add_argument("--period", required=True, help="时间段 (如: 2025-01~2025-12)")
    p_scrape.add_argument("--types", help="逗号分隔的工种/材料名称（可选）")
    p_scrape.add_argument("--category", choices=["labor", "material"], help="按类别过滤")
    p_scrape.add_argument("--output", default=DEFAULT_DATA_DIR, help="输出目录")

    # === chart ===
    p_chart = subparsers.add_parser("chart", help="生成价格走势图")
    p_chart.add_argument("--city", required=True, help="城市名称")
    p_chart.add_argument("--types", help="逗号分隔的工种/材料名称（可选）")
    p_chart.add_argument("--output", default=DEFAULT_DATA_DIR, help="数据目录")

    # === query ===
    p_query = subparsers.add_parser("query", help="查询已采集的数据")
    p_query.add_argument("--city", required=True, help="城市名称")
    p_query.add_argument("--types", help="逗号分隔的工种/材料名称（可选）")
    p_query.add_argument("--json", action="store_true", help="输出 JSON 格式")
    p_query.add_argument("--output", default=DEFAULT_DATA_DIR, help="数据目录")

    args = parser.parse_args()

    commands = {
        "run": cmd_run,
        "analyze": cmd_analyze,
        "import": cmd_import,
        "scrape": cmd_scrape,
        "chart": cmd_chart,
        "query": cmd_query,
    }

    commands[args.command](args)


if __name__ == "__main__":
    main()
