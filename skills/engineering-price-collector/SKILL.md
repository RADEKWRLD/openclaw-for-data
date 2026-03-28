---
name: engineering-price-collector
description: "采集各省市工程造价信息价（人工费、材料价格）。用于：用户询问工程信息价、造价信息、人工费、材料价格、价格走势时触发。自动从造价信息网下载 Excel 或解析已有文件，生成 JSON + 图表 + Excel 汇总报告。"
metadata:
  {"openclaw": {"emoji": "🏗️", "requires": {"bins": ["python3"]}}}
---

# 工程信息价采集工具

## 使用流程

### 1. 获取数据

如果用户没有提供 Excel 文件，先用 web_search 搜索并用 browser 下载：

- 搜索关键词：「<城市> <年月> 工程造价信息价 下载」
- 下载 Excel 文件，**保存到 `{baseDir}/../../skill-data/downloads/<城市>/` 目录**（容器内路径：`/data/workspace/skill-data/downloads/<城市>/`）

### 2. 运行分析

```bash
bash {baseDir}/run.sh run \
  --city "<城市>" \
  --period "<时间段>" \
  --excel /data/workspace/skill-data/downloads/<城市>/<文件名>.xls
```

无 `--excel` 时工具会尝试自动爬取，失败后会提示搜索下载。

### 参数

| 参数 | 必填 | 说明 |
|------|------|------|
| `--city` | 是 | 城市名称，如「广州」 |
| `--period` | 是 | 时间段，如 `2026-03` 或 `2025-01~2025-12` |
| `--excel` | 否 | Excel 文件路径，支持多个文件空格分隔 |
| `--types` | 否 | 过滤工种，如 `钢筋工,模板工,普工` |
| `--category` | 否 | `labor`(人工) 或 `material`(材料) |

### 输出

所有输出保存在 `/data/workspace/skill-data/prices/<城市>/`：

- `prices.json` — 全量结构化数据
- `prices_<城市>_<时段>.xlsx` — 多 Sheet Excel 汇总表
- `charts/<Sheet>_人工.png` — 人工价格柱状图
- `charts/<Sheet>_材料分类.png` — 材料分类均价图

### 示例

```bash
# 单月
bash {baseDir}/run.sh run --city "广州" --period "2026-03" \
  --excel /data/workspace/skill-data/downloads/广州/2026年3月信息价.xls

# 多月
bash {baseDir}/run.sh run --city "广州" --period "2025-01~2025-12" \
  --excel /data/workspace/skill-data/downloads/广州/1月.xls \
         /data/workspace/skill-data/downloads/广州/2月.xls

# 只看人工
bash {baseDir}/run.sh run --city "广州" --period "2026-03" --category labor \
  --excel /data/workspace/skill-data/downloads/广州/2026年3月信息价.xls
```
