# 工程信息价采集工具 (Engineering Price Collector)

## Description

从各省市造价信息网**自动下载**工程信息价 Excel 文件，智能分析表格结构，提取人工费和材料价格，按 Sheet 分组生成走势图和 Excel 汇总表。

## Workflow

```
用户输入(城市+时间段)
  → scrape: AI 自动从官网下载 Excel → downloads/<city>/
  → import: 自动分析表格结构 → prices.json
  → chart: 按 Sheet 分组生成走势图 + Excel 汇总表
```

## Usage

```bash
SKILL=/data/workspace/skills/engineering-price-collector/run.sh

# 自动爬取（下载 Excel → 解析 → 存 JSON）
bash $SKILL scrape --city "广州" --period "2025-01~2025-12"

# 分析已有 Excel 文件结构
bash $SKILL analyze --excel /data/workspace/skill-data/downloads/广州/2026-03_信息价.xls

# 手动导入（所有 Sheet 全量导入）
bash $SKILL import --city "广州" --sheet all \
  --excel /data/workspace/skill-data/downloads/广州/2026-03_信息价.xls

# 生成分组图表 + Excel
bash $SKILL chart --city "广州"

# 只看特定工种
bash $SKILL chart --city "广州" --types "钢筋工,模板工,普工,一般抹灰工"

# 查询数据
bash $SKILL query --city "广州" --json
```

## Sub-commands

| Command | Description |
|---------|-------------|
| `scrape` | 从官网自动下载 Excel 到 downloads/ → 解析 → 存 JSON |
| `analyze` | 分析 Excel 表格结构（Sheet、列、人工/材料类型） |
| `import` | 从本地 Excel 导入（`--sheet all` 导入所有 Sheet） |
| `chart`  | 按 Sheet 分组生成走势图（人工/材料分开，材料自动分页） |
| `query`  | 查询已存储的数据 |

## Output Structure

```
skill-data/
├── downloads/<city>/                   # AI 自动下载的原始 Excel
│   ├── 2025-01_信息价.xls
│   └── 2025-02_信息价.xls
└── prices/<city>/
    ├── prices.json                     # 全量结构化数据
    ├── prices_<city>_<period>.xlsx     # Excel 汇总（按原始 Sheet 分 Tab）
    └── charts/
        ├── 建筑安装_人工_trend.png
        ├── 建筑安装_人工_comparison_2026-03.png
        ├── 建筑安装_材料_trend.png
        ├── 建筑安装_材料_trend_p2.png   # 材料超 20 条自动分页
        ├── 市政_人工_trend.png
        ├── 市政_材料_trend.png
        └── ...
```

## Adding New City Support

在 `src/cities/` 下创建新文件，继承 `BaseCityScraper`，实现：
- `get_download_page_url(year, month)` — 信息价下载页 URL
- `find_excel_links(html, year, month)` — 从页面中解析 Excel 下载链接

然后在 `src/cities/__init__.py` 中注册。
