"""Generate price trend charts and Excel exports from collected data.

Charts are generated per-sheet with labor and material items separated.
Materials are intelligently grouped by category (钢材类, 管材管件, 苗木花卉, etc.)
rather than blindly paginated.
"""

from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path

import matplotlib
matplotlib.use("Agg")

import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties, fontManager

from .models import PriceData, PriceItem

MAX_ITEMS_PER_CHART = 20

# ─── Material categorization rules ───
# Order matters: first match wins
_MATERIAL_CATEGORIES = [
    (r'钢筋|钢管|钢板|钢丝|钢缆|钢模|钢桩|钢轨|角钢|工字钢|H型钢|扁钢|钢带|铁丝|焊条|焊丝|槽钢|型钢|卷板|钢绞|圆钢|预应力', '钢材类'),
    (r'混凝土|砼|C\d+', '混凝土类'),
    (r'水泥|硅酸盐', '水泥类'),
    (r'砂浆|干混', '砂浆类'),
    (r'沥青', '沥青类'),
    (r'砂|碎石|石屑|石灰|矿粉|石粉|片石|块石|卵石|石子|道碴|粉煤灰', '砂石料'),
    (r'管|套筒|接头|DN\d|弯头|三通|四通|法兰|卡箍|喉箍', '管材管件'),
    (r'泵|阀|闸|蝶|止回|减压|排气|球阀', '泵阀类'),
    (r'电缆|电线|母线|铜芯|BV|屏蔽|导线|接线盒', '电线电缆'),
    (r'桥架|线槽', '桥架类'),
    (r'挖掘|推土|压路|起重|履带|装载|铲运|摊铺|搅拌|振动|打桩|吊车|吊装|发电|空压|卷扬|切割|钻|破碎|碾压|铣刨|洒布|抛丸|凿岩|潜孔|旋挖|注浆|喷射|台车|拌合|载重|自卸|翻斗|电梯|压缩机|汽车|运输车|拖车', '机械设备'),
    (r'雪松|湿地松|龙柏|榕|桂|杜鹃|茉莉|罗汉|木棉|紫荆|樟|兰|苗|花|棕|蕨|藤|灌木|乔木|地被|种子|松|柏|梅|桃|荔|芒|椰|棕榈|紫薇|凤凰|叶|山茶|红枫|腊肠|朴|栾|楠|扶桑|假连翘|九里香|福建茶|银杏|丹桂|鹅掌|勒杜鹃|洋紫荆|含笑|水杉|红锥|落羽|合欢|木荷|枫香|无患子|鸡蛋花|蒲葵|红千层|麻楝|美人蕉|蒲桃|荷花|睡莲|麦冬|鸢尾|翠芦莉|马尼拉|结缕|狗牙根|黑麦|百喜|铁冬青|海桐|毛杜鹃|芒果|大叶|小叶|红花|白花|黄花|黄金|草坪|草皮|竹', '苗木花卉'),
    (r'信号灯|标志|交通|标线|标牌|反光|护栏|隔离|波形|防撞', '交通设施'),
    (r'灯|照明|LED|投光', '照明类'),
    (r'防水|卷材', '防水材料'),
    (r'防火门|木质|钢质', '门窗类'),
    (r'橡胶|支座|聚四氟|胶管|密封|止水|胶条|胶圈', '橡胶密封'),
    (r'PE管|PVC|HDPE|PPR|聚乙烯|聚氯乙烯|塑料|尼龙|PP[管件]', '塑料制品'),
    (r'铸铁|球墨', '铸铁制品'),
    (r'井盖|窨井|检查井|雨水篦|集水', '市政配件'),
    (r'螺丝|螺栓|螺母|锚栓|铆钉|膨胀', '紧固件'),
    (r'土工|无纺布|滤布|网格|麻袋|草袋|稻草|纤维|编织', '辅助材料'),
    (r'油漆|涂料|脱漆|汽油|柴油|润滑|机油', '油料涂料'),
    (r'砖|瓦|花岗|大理|板岩', '砖石材料'),
    (r'木方|板材|胶合板|纤维板|模板|原木|方木', '木材类'),
    (r'涵|盖板|预制|构件|桩基|墩', '预制构件'),
    (r'警示|安全', '安全设施'),
    (r'监测|仪器|仪表', '仪器仪表'),
]


def _classify_material(name: str) -> str:
    """Classify a material item into a category by name pattern matching."""
    for pattern, category in _MATERIAL_CATEGORIES:
        if re.search(pattern, name):
            return category
    return "其他"
COLORS = ["#e74c3c", "#3498db", "#2ecc71", "#f39c12", "#9b59b6",
          "#1abc9c", "#e67e22", "#34495e", "#c0392b", "#2980b9",
          "#27ae60", "#d35400", "#8e44ad", "#16a085", "#f1c40f",
          "#7f8c8d", "#2c3e50", "#e91e63", "#00bcd4", "#ff9800"]


def _find_cjk_font() -> FontProperties | None:
    preferred = [
        "Noto Sans CJK SC", "Noto Sans SC", "PingFang SC", "Heiti SC",
        "SimHei", "Microsoft YaHei", "WenQuanYi Micro Hei",
        "Source Han Sans SC", "Arial Unicode MS",
    ]
    available_names = {f.name for f in fontManager.ttflist}
    for name in preferred:
        if name in available_names:
            return FontProperties(family=name)
    for f in fontManager.ttflist:
        if "CJK" in f.name or "SC" in f.name:
            return FontProperties(family=f.name)
    return None


def _safe_filename(name: str) -> str:
    """Convert sheet name to safe filename (replace commas, spaces)."""
    return name.replace("，", "").replace(",", "").replace(" ", "").replace("/", "_")


def _plot_trend(
    items: list[PriceItem],
    title: str,
    output_path: Path,
    font_prop: FontProperties | None,
    show_range: bool = True,
) -> Path:
    """Plot a single trend chart for a list of items."""
    fig, ax = plt.subplots(figsize=(14, 7))

    for i, item in enumerate(items):
        if not item.monthly:
            continue
        months = [m.month for m in item.monthly]
        avgs = [m.price_avg for m in item.monthly]
        color = COLORS[i % len(COLORS)]

        ax.plot(months, avgs, marker="o", color=color, linewidth=2,
                markersize=5, label=item.name)

        if show_range:
            mins = [m.price_min for m in item.monthly]
            maxs = [m.price_max for m in item.monthly]
            if any(mn != mx for mn, mx in zip(mins, maxs)):
                ax.fill_between(months, mins, maxs, alpha=0.12, color=color)

        if avgs:
            ax.annotate(f"{avgs[-1]:.0f}", xy=(months[-1], avgs[-1]),
                        textcoords="offset points", xytext=(8, 0),
                        fontsize=8, color=color, fontproperties=font_prop)

    ax.set_title(title, fontsize=14, fontweight="bold", fontproperties=font_prop)
    ax.set_xlabel("月份", fontsize=11, fontproperties=font_prop)

    unit = items[0].unit if items else ""
    ax.set_ylabel(f"价格 (元/{unit})" if unit else "价格 (元)",
                  fontsize=11, fontproperties=font_prop)

    plt.xticks(rotation=45, ha="right", fontproperties=font_prop)
    plt.yticks(fontproperties=font_prop)

    # Legend: outside right if many items, inside if few
    if len(items) > 10:
        ax.legend(fontsize=8, prop=font_prop, framealpha=0.9,
                  loc="upper left", bbox_to_anchor=(1.02, 1), borderaxespad=0)
    else:
        ax.legend(fontsize=9, prop=font_prop, framealpha=0.9, loc="best")

    ax.grid(True, alpha=0.3, linestyle="--")
    ax.set_axisbelow(True)

    plt.tight_layout()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return output_path


def _plot_comparison(
    items: list[PriceItem],
    month: str,
    title: str,
    output_path: Path,
    font_prop: FontProperties | None,
) -> Path | None:
    """Plot a bar chart comparing items for a specific month."""
    names, avgs, mins, maxs = [], [], [], []
    for item in items:
        for mp in item.monthly:
            if mp.month == month:
                names.append(item.name)
                avgs.append(mp.price_avg)
                mins.append(mp.price_min)
                maxs.append(mp.price_max)
                break

    if not names:
        return None

    width = max(10, len(names) * 0.8)
    fig, ax = plt.subplots(figsize=(width, 6))

    bar_colors = [COLORS[i % len(COLORS)] for i in range(len(names))]
    bars = ax.bar(names, avgs, color=bar_colors, alpha=0.8, edgecolor="white", linewidth=1.2)

    yerr_low = [a - mn for a, mn in zip(avgs, mins)]
    yerr_high = [mx - a for a, mx in zip(avgs, maxs)]
    ax.errorbar(names, avgs, yerr=[yerr_low, yerr_high], fmt="none", ecolor="gray", capsize=4)

    for bar, avg in zip(bars, avgs):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 2,
                f"{avg:.0f}", ha="center", fontsize=9, fontweight="bold",
                fontproperties=font_prop)

    ax.set_title(title, fontsize=14, fontweight="bold", fontproperties=font_prop)
    unit = items[0].unit if items else ""
    ax.set_ylabel(f"价格 (元/{unit})" if unit else "价格 (元)",
                  fontsize=11, fontproperties=font_prop)

    plt.xticks(rotation=45, ha="right", fontproperties=font_prop, fontsize=9)
    plt.yticks(fontproperties=font_prop)
    ax.grid(True, axis="y", alpha=0.3, linestyle="--")
    ax.set_axisbelow(True)
    plt.tight_layout()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return output_path


def _plot_category_summary(
    items: list[PriceItem],
    title: str,
    output_path: Path,
    font_prop: FontProperties | None,
) -> Path:
    """Plot a bar chart showing average price per material category.

    Groups items by _classify_material, computes category-level avg,
    and draws one bar per category.
    """
    by_cat: dict[str, list[float]] = defaultdict(list)
    for it in items:
        cat = _classify_material(it.name)
        if it.monthly:
            by_cat[cat].append(it.monthly[-1].price_avg)

    # Sort by count descending
    cat_data = sorted(by_cat.items(), key=lambda x: -len(x[1]))

    names = []
    avgs = []
    counts = []
    for cat, prices in cat_data:
        names.append(f"{cat}\n({len(prices)}种)")
        avgs.append(sum(prices) / len(prices))
        counts.append(len(prices))

    if not names:
        return output_path

    fig, ax = plt.subplots(figsize=(max(10, len(names) * 0.9), 6))
    bar_colors = [COLORS[i % len(COLORS)] for i in range(len(names))]
    bars = ax.bar(names, avgs, color=bar_colors, alpha=0.85, edgecolor="white", linewidth=1.2)

    for bar, avg in zip(bars, avgs):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 2,
                f"{avg:.0f}", ha="center", fontsize=9, fontweight="bold",
                fontproperties=font_prop)

    ax.set_title(title, fontsize=14, fontweight="bold", fontproperties=font_prop)
    ax.set_ylabel("分类均价 (元)", fontsize=11, fontproperties=font_prop)
    plt.xticks(rotation=45, ha="right", fontproperties=font_prop, fontsize=9)
    plt.yticks(fontproperties=font_prop)
    ax.grid(True, axis="y", alpha=0.3, linestyle="--")
    ax.set_axisbelow(True)
    plt.tight_layout()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return output_path


def generate_charts_by_sheet(
    data: PriceData,
    output_dir: str | Path,
    types: list[str] | None = None,
) -> list[Path]:
    """Generate charts grouped by source sheet.

    Per sheet, generates at most 2 charts:
    - <sheet>_人工.png — labor price comparison bar chart
    - <sheet>_材料分类.png — material category summary bar chart

    Total ≤ ~10 charts for a typical 5-sheet file.
    """
    output_dir = Path(output_dir)
    charts_dir = output_dir / "charts"
    charts_dir.mkdir(parents=True, exist_ok=True)
    font_prop = _find_cjk_font()

    items = data.prices
    if types:
        items = [item for item in items if item.name in types]

    by_sheet: dict[str, list[PriceItem]] = defaultdict(list)
    for item in items:
        by_sheet[item.sheet or "未分类"].append(item)

    generated: list[Path] = []
    city = data.meta.city
    period = data.meta.period

    for sheet_name, sheet_items in by_sheet.items():
        safe_name = _safe_filename(sheet_name)
        labor = [it for it in sheet_items if it.category == "labor"]
        material = [it for it in sheet_items if it.category == "material"]

        # --- 1 labor chart: bar comparison ---
        if labor:
            latest = labor[0].monthly[-1].month if labor[0].monthly else None
            if latest:
                path = _plot_comparison(
                    labor, latest,
                    title=f"{city} {sheet_name} - 人工价格 ({latest})",
                    output_path=charts_dir / f"{safe_name}_人工.png",
                    font_prop=font_prop,
                )
                if path:
                    generated.append(path)
                    print(f"  [chart] {path.name} ({len(labor)} 个工种)")

        # --- 1 material chart: category summary ---
        if material:
            path = _plot_category_summary(
                material,
                title=f"{city} {sheet_name} - 材料分类均价 ({period})",
                output_path=charts_dir / f"{safe_name}_材料分类.png",
                font_prop=font_prop,
            )
            generated.append(path)
            print(f"  [chart] {path.name} ({len(material)} 种材料)")

    # --- Overview for --types ---
    if types and len(items) <= MAX_ITEMS_PER_CHART:
        path = _plot_trend(
            items,
            title=f"{city} 指定品类价格走势 ({period})",
            output_path=charts_dir / "overview.png",
            font_prop=font_prop,
        )
        generated.append(path)
        print(f"  [chart] overview.png ({len(items)} 个品类)")

    return generated


# Keep the old functions for backward compatibility / simple use cases
def generate_trend_chart(
    data: PriceData,
    output_dir: str | Path,
    types: list[str] | None = None,
    show_range: bool = True,
    title: str | None = None,
) -> Path:
    """Generate a single trend chart (legacy, for small datasets)."""
    output_dir = Path(output_dir)
    charts_dir = output_dir / "charts"
    font_prop = _find_cjk_font()

    items = data.prices
    if types:
        items = [item for item in items if item.name in types]
    if not items:
        raise ValueError("没有可绘制的数据")

    chart_title = title or f"{data.meta.city} 工程信息价走势 ({data.meta.period})"
    return _plot_trend(items, chart_title, charts_dir / "trend.png", font_prop, show_range)


def generate_comparison_chart(
    data: PriceData,
    output_dir: str | Path,
    month: str,
    types: list[str] | None = None,
    title: str | None = None,
) -> Path:
    """Generate a single comparison chart (legacy)."""
    output_dir = Path(output_dir)
    charts_dir = output_dir / "charts"
    font_prop = _find_cjk_font()

    items = data.prices
    if types:
        items = [item for item in items if item.name in types]

    chart_title = title or f"{data.meta.city} {month} 工种价格对比"
    result = _plot_comparison(items, month, chart_title, charts_dir / f"comparison_{month}.png", font_prop)
    if result is None:
        raise ValueError(f"没有 {month} 月份的数据")
    return result


# ─── Excel export ───

def _auto_col_width(ws, df) -> None:
    for col_idx, col in enumerate(df.columns, 1):
        max_len = max(len(str(col)), df[col].astype(str).str.len().max())
        letter = ws.cell(row=1, column=col_idx).column_letter
        ws.column_dimensions[letter].width = min(max_len + 4, 35)


def export_xlsx(
    data: PriceData,
    output_dir: str | Path,
    types: list[str] | None = None,
) -> Path:
    """Export price data to a multi-sheet Excel, grouped by source sheet.

    Output sheets:
    - One sheet per source (建筑安装, 市政, etc.) with all items
    - 人工汇总: All labor items across sheets
    - 信息: Metadata
    """
    import pandas as pd

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    items = data.prices
    if types:
        items = [item for item in items if item.name in types]
    if not items:
        raise ValueError("没有可导出的数据")

    all_months = sorted({mp.month for item in items for mp in item.monthly})

    def build_rows(item_list):
        rows = []
        for item in item_list:
            row: dict = {
                "序号": len(rows) + 1,
                "名称": item.name,
                "编码": item.code,
                "类型": "人工" if item.category == "labor" else "材料",
                "单位": item.unit,
            }
            month_map = {mp.month: mp for mp in item.monthly}
            for m in all_months:
                mp = month_map.get(m)
                if mp:
                    row[f"{m}_最低价"] = mp.price_min
                    row[f"{m}_最高价"] = mp.price_max
                    row[f"{m}_均价"] = mp.price_avg
                else:
                    row[f"{m}_最低价"] = None
                    row[f"{m}_最高价"] = None
                    row[f"{m}_均价"] = None
            rows.append(row)
        return rows

    # Group by sheet
    by_sheet: dict[str, list[PriceItem]] = defaultdict(list)
    for item in items:
        by_sheet[item.sheet or "未分类"].append(item)

    labor_items = [it for it in items if it.category == "labor"]
    material_items = [it for it in items if it.category == "material"]

    output_path = output_dir / f"prices_{data.meta.city}_{data.meta.period}.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Per-sheet tabs
        for sheet_name, sheet_items in by_sheet.items():
            # Excel sheet name max 31 chars
            tab_name = sheet_name[:31]
            df = pd.DataFrame(build_rows(sheet_items))
            df.to_excel(writer, sheet_name=tab_name, index=False)
            _auto_col_width(writer.sheets[tab_name], df)

        # Labor summary
        if labor_items:
            df_labor = pd.DataFrame(build_rows(labor_items))
            df_labor.to_excel(writer, sheet_name="人工汇总", index=False)
            _auto_col_width(writer.sheets["人工汇总"], df_labor)

        # Metadata
        sheets_summary = []
        for sn, si in by_sheet.items():
            labor_count = sum(1 for it in si if it.category == "labor")
            mat_count = sum(1 for it in si if it.category == "material")
            sheets_summary.append(f"{sn}: 人工{labor_count} 材料{mat_count}")

        meta_rows = [
            {"项目": "城市", "值": data.meta.city},
            {"项目": "时段", "值": data.meta.period},
            {"项目": "来源", "值": data.meta.source},
            {"项目": "采集时间", "值": data.meta.collected_at},
            {"项目": "人工条目数", "值": len(labor_items)},
            {"项目": "材料条目数", "值": len(material_items)},
            {"项目": "总条目数", "值": len(items)},
            {"项目": "Sheet数", "值": len(by_sheet)},
            {"项目": "月份数", "值": len(all_months)},
            {"项目": "月份列表", "值": ", ".join(all_months)},
            {"项目": "Sheet明细", "值": " | ".join(sheets_summary)},
        ]
        df_meta = pd.DataFrame(meta_rows)
        df_meta.to_excel(writer, sheet_name="信息", index=False)
        _auto_col_width(writer.sheets["信息"], df_meta)

    return output_path
