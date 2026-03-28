# 工程信息价采集工具 (Engineering Price Collector)

## Description

从各省市造价信息网自动下载工程信息价 Excel 文件，智能分析表格结构，提取人工费和材料价格，按 Sheet 分组生成走势图和 Excel 汇总表。

## Quick Start — AI 只需执行一条命令

```bash
bash /data/workspace/skills/engineering-price-collector/run.sh run \
  --city "广州" \
  --excel /data/workspace/skill-data/downloads/广州/2026-03_信息价.xls \
  --output /data/workspace/skill-data
```

一条命令完成全部流程：**导入全部 Sheet → 智能分类 → 生成图表 + Excel → 输出报告**

### 可选参数

```bash
# 只看人工
--category labor

# 只看指定工种
--types "钢筋工,模板工,普工,一般抹灰工"

# 多个月份文件一起导入
--excel 1月.xls 2月.xls 3月.xls
```

## Output

```
skill-data/
├── downloads/<city>/                   # AI 自动下载 / 用户放入的原始 Excel
└── prices/<city>/
    ├── prices.json                     # 全量结构化数据
    ├── prices_<city>_<period>.xlsx     # 多 Sheet Excel 汇总
    └── charts/                         # ≤10 张图表
        ├── <sheet>_人工.png            # 人工价格柱状图
        └── <sheet>_材料分类.png        # 材料分类均价柱状图
```

## All Commands

| Command | Description |
|---------|-------------|
| `run`     | **一键执行**: 导入 → 图表 → xlsx → 报告 |
| `scrape`  | 从官网自动下载 Excel → 导入 → 报告 |
| `analyze` | 分析 Excel 结构（调试用） |
| `import`  | 仅导入数据 |
| `chart`   | 仅生成图表 |
| `query`   | 查询已有数据 |
