# 工程信息价采集工具 (Engineering Price Collector)

## AI 调用方式

### 方式一：全自动（无 Excel，工具自己爬）

```bash
bash /data/workspace/skills/engineering-price-collector/run.sh run \
  --city "广州" --period "2025-01~2025-12"
```

工具自动从造价信息网下载 Excel → 解析 → 图表 → 报告

### 方式二：AI 先用浏览器下载，再分析

1. AI 用浏览器访问造价信息网，下载 Excel 到 `skill-data/downloads/<城市>/`
2. 运行分析：

```bash
bash /data/workspace/skills/engineering-price-collector/run.sh run \
  --city "广州" --period "2026-03" \
  --excel /data/workspace/skill-data/downloads/广州/2026年3月信息价.xls
```

### 参数

| 参数 | 必填 | 说明 |
|------|------|------|
| `--city` | 是 | 城市名称 |
| `--period` | 是 | 时间段，如 `2026-03` 或 `2025-01~2025-12` |
| `--excel` | 否 | 已有 Excel 路径（不传则自动爬取），支持多个文件 |
| `--types` | 否 | 过滤工种，如 `钢筋工,模板工,普工` |
| `--category` | 否 | `labor`(人工) 或 `material`(材料) |

## Output

```
skill-data/
├── downloads/<city>/               # 原始 Excel（自动下载或 AI 放入）
└── prices/<city>/
    ├── prices.json                  # 全量 JSON
    ├── prices_<city>_<period>.xlsx  # 多 Sheet Excel 汇总
    └── charts/                      # ~10 张图表
        ├── <sheet>_人工.png         # 人工价格柱状图
        └── <sheet>_材料分类.png     # 材料分类均价图
```

## 自动爬取失败时

如果自动爬取报错，AI 应使用**内置浏览器工具**手动完成下载：

1. 用浏览器工具搜索「<城市> 工程造价 信息价 <年月>」
2. 访问对应的造价信息网站，找到 Excel 下载链接
3. 下载文件到 `skill-data/downloads/<城市>/`
4. 用 `--excel` 参数重新运行 `run` 命令

> **注意**: AI 运行环境内置了浏览器工具（Browser Tool），可以直接访问网页、点击下载链接。不要让用户手动操作。

## 数据源

| 城市 | 网站 |
|------|------|
| 广州 | https://gc.gzcc.gov.cn |

新增城市：在 `src/cities/` 下添加插件
