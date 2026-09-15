---
name: asin-pricepulse
description: ASIN PricePulse 价格脉搏 — 批量抓取 Amazon 17 个站点的 ASIN 本地原生价格,生成 HTML 洞察报告 + Excel 明细。当用户请求跨站点价格对比、竞品 ASIN 价格情报、品类价盘调研时激活。与用户 VPN 状态无关。AST 内部工具。
version: 1.3
---

# ASIN PricePulse 价格脉搏

> 批量抓取 Amazon 17 个站点的 ASIN 本地原生价格,生成交互报告。

## When to activate

- Batch ASIN price lookup across Amazon marketplaces
- Cross-market pricing comparison
- Competitive category pricing / brand price analysis
- Any request mentioning "竞品价格" / "ASIN 价格" / "价格对比"

## Pre-flight reminder (ALWAYS show before running)

```
🔍 PricePulse 使用前提醒:
• 抓取的是本地访客挂牌价,不是权威定价数据源
• 每个市场自动切换本地邮编 + 币种
• 与你 VPN 状态无关(cookie 注入)
• 仅限 AST 内部使用

预计耗时: ≤30→1-2min | 30-100→3-8min | 100-300→15-40min | 300+→慢跑
```

## Invocation

```bash
pricepulse --markets US --input /path/to/asins.xlsx --output /path/to/out --title "Report Title" --subtitle "2026.5.1-7.31" --yes
```

If `pricepulse` is not on PATH: `python -m pricepulse.cli ...`

### Key flags
- `--markets, -m` — comma-separated: US,UK,DE,FR,IT,ES,NL,SE,PL,JP,CA,AU,AE,SA,SG,IN,MX,BR,TR
- `--input, -i` — .xlsx/.csv/.txt with ASIN column (auto-detects Search Rank, Purchase Rank, Keyword Searches)
- `--asins, -a` — or inline comma-separated ASINs
- `--delay 1.0 --workers 3` — for large batches (>100)
- `--title` / `--subtitle` — report header

### Common market presets
- EU5 → `DE,FR,IT,ES,UK`
- NA → `US,CA,MX`
- APAC → `JP,SG,IN,AU`

## Output

- `asin_detail_*.xlsx` — styled Excel (Segoe UI, Indigo header, frozen pane)
- `竞品ASIN分析报告_*.html` — interactive report:
  - Price Tier Cards (clickable → filters volume chart)
  - ASIN Count vertical bar chart (clickable)
  - Brand Avg Price Top 12 (clickable)
  - Brand Bubble Cloud Top 20 (clickable)
  - Search Volume by ASIN Top 15 (dynamically filtered)
  - Detail table with sort + filter (Select All toggle)

## Brand extraction (byline-first)

1. Byline anchor ("Visit the X Store")
2. Store URL → Product table → JSON → Title first word (last resort)

## Price banding

3 equal-width tiers: Entry / Mid-tier / Premium, based on (max-min)/3.

## Do NOT
- Do NOT call for single ASIN — use fetch_external_page
- Do NOT invoke without pre-flight reminder
- Do NOT share raw output externally
