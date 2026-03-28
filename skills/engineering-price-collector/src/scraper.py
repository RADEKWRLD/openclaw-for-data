"""Scraper orchestration — downloads Excel files then imports them."""

from __future__ import annotations

from pathlib import Path

from .cities import get_scraper
from .excel_import import import_excel
from .models import PriceData


def scrape_prices(
    city: str,
    year: int,
    months: list[int],
    types: list[str] | None = None,
    category: str | None = None,
    data_dir: str | Path = "/data/workspace/skill-data",
) -> PriceData:
    """Scrape price data: download Excel files from official sites, then import.

    Flow:
    1. Use city plugin to find Excel download links on the official site
    2. Download Excel files to skill-data/downloads/<city>/
    3. Import each downloaded file using excel_import
    4. Merge all months into one PriceData

    Args:
        city: City name (e.g., "广州").
        year: Target year.
        months: List of month numbers (1-12).
        types: Optional type name filter.
        category: Optional "labor" or "material" filter.
        data_dir: Base data directory.

    Returns:
        PriceData with all scraped prices.
    """
    scraper = get_scraper(city)
    download_dir = Path(data_dir) / "downloads" / city

    print(f"[scrape] 城市: {scraper.CITY_NAME}")
    print(f"[scrape] 数据源: {scraper.SOURCE_NAME}")
    print(f"[scrape] 目标: {year}年 {[f'{m}月' for m in months]}")
    print(f"[scrape] 下载目录: {download_dir}")

    # Step 1: Download Excel files
    downloaded = scraper.download_excels(year, months, download_dir)

    if not downloaded:
        months_str = "、".join(f"{m}月" for m in months)
        raise RuntimeError(
            f"自动下载失败。请按以下步骤操作：\n"
            f"\n"
            f"Step 1: 用 web_search 搜索下载地址\n"
            f"   搜索关键词: \"{city} {year}年{months_str} 工程造价信息价 下载\"\n"
            f"\n"
            f"Step 2: 用浏览器工具打开搜索到的页面，找到 Excel 下载链接并下载\n"
            f"   参考网站: {scraper.SOURCE_NAME} {scraper.BASE_URL}\n"
            f"   保存到: {download_dir}/\n"
            f"\n"
            f"Step 3: 下载完成后重新运行:\n"
            f"   run --city {city} --period {year}-{months[0]:02d} --excel {download_dir}/<文件名>.xls"
        )

    print(f"\n[scrape] 已下载 {len(downloaded)} 个文件，开始解析...")

    # Step 2: Import each downloaded file
    merged: PriceData | None = None

    for month_str, file_path in sorted(downloaded.items()):
        print(f"\n[import] {month_str}: {file_path.name}")
        try:
            data = import_excel(
                excel_path=file_path,
                city=city,
                month=month_str,
                sheet_name="all",
                types=types,
                category=category,
                source=scraper.SOURCE_NAME,
            )
            if merged is None:
                merged = data
            else:
                merged.merge(data)
        except Exception as e:
            print(f"  [error] 解析失败: {e}")

    if merged is None:
        raise RuntimeError("所有文件解析均失败")

    # Update period
    all_months = sorted({mp.month for item in merged.prices for mp in item.monthly})
    if len(all_months) > 1:
        merged.meta.period = f"{all_months[0]}~{all_months[-1]}"
    elif all_months:
        merged.meta.period = all_months[0]

    merged.meta.source = scraper.SOURCE_NAME

    total_items = len(merged.prices)
    total_records = sum(len(item.monthly) for item in merged.prices)
    print(f"\n[scrape] 完成: {total_items} 个品类, {total_records} 条月度记录")

    return merged
