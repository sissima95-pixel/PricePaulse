---
name: asin-pricepulse
description: ASIN PricePulse (alias PP / pp) 价格脉搏 — 批量抓取 Amazon 17 个站点的 ASIN 本地原生价格,生成 HTML 洞察报告(价格段 / 品牌份额 / 竞争者分析 / Search Volume 交互图)+ Excel 明细。当用户说 "用 pp"、"pricepulse"、"竞品价格"、"ASIN 价格"、"价格对比" 时激活。与用户 VPN 状态无关。AST 内部工具。
version: 1.4
---

# ASIN PricePulse 价格脉搏 (PP)

## Aliases

**PricePulse** = **pricepulse** = **PP** = **pp**. "用 pp 帮我..." / "请 pp 处理..." → activate this skill.

## When to activate

- Batch ASIN price lookup across Amazon marketplaces
- Cross-market pricing comparison / competitive category pricing
- Brand landscape / price tier analysis for a category
- Any mention of "竞品价格" / "ASIN 价格" / "价格对比" / "pp"

## Pre-flight reminder (ALWAYS show before running)

```
PricePulse 使用前提醒:
• 抓取的是本地访客挂牌价快照,不是权威定价数据源
• 每个市场自动切换本地邮编 + 币种,与 VPN 状态无关
• 仅限 AST 内部使用
预计耗时: ≤30 ASIN → 1-2 min | 30-100 → 3-8 min | 100-300 → 15-40 min | 300+ → 用慢速模式
```

## Invocation

```bash
pricepulse --markets US --input /path/asins.xlsx --output /path/out --title "..." --subtitle "2026.5.1-7.31" --yes
```
If not on PATH: `python -m pricepulse.cli ...`

Flags: `--markets/-m` (US,UK,DE,FR,IT,ES,NL,SE,PL,JP,CA,AU,AE,SA,SG,IN,MX,BR,TR) · `--input/-i` (.xlsx/.csv/.txt, auto-detects Search Rank / Purchase Rank / Keyword Searches) · `--asins/-a` inline · `--delay 1.0 --workers 3` for 300+ ("慢速模式") · `--title` / `--subtitle`

Presets: EU5 = DE,FR,IT,ES,UK · NA = US,CA,MX · APAC = JP,SG,IN,AU

### Custom price tiers
If the user specifies tiers (e.g. "0-40, 40-80, 80+"), monkey-patch `pricepulse.reporter._price_bands` to return `[(label, count, lo, hi), ...]` with those bounds before calling `write_html`. Otherwise default = tercile split.

## Output

- `asin_detail_*.xlsx` — styled Excel
- `竞品ASIN分析报告_*.html` — sections:
  1. Price Tier Overview (3 cards) + Insight A. PRICE OVERVIEW
  2. Distribution — ASIN count by tier (clickable) + Brand Avg Price Top 12 (clickable)
  3. Brand Share — donut by Search Volume + donut by ASIN Count + Insight B. BRAND LANDSCAPE + C. INDIVIDUAL PLAYER ANALYSIS (top 3 brands)
  4. Brand Share by Price Tier — Top 5 per tier, color dot table, expandable ASIN sub-table
  5. Search Volume by ASIN Top 15 — brand bubble cloud (click to filter) + Copy ASINs
  6. ASIN Detail — sort/filter, Copy ASINs, Export Excel
- Insights are EN paragraph + CN italic paragraph, no emoji.

## Brand extraction (v1.4)
Priority: po-brand table → "Brand Name" row → byline "Visit the X Store" → JSON → store URL (demoted: sponsored-store false positives) → title first word (last resort). 80+ descriptive words blocklisted (Portable, Wireless, Neck...). Auto retry pass for OK results with empty brand. "Unknown" always sorted last / folded into OTHERS.

## Price banding
Default 3 tiers by tercile (33rd/67th pct), bounds rounded to nice numbers → balanced distribution. Python + JS share the same bounds (TIER_BOUNDS JSON).

## Do NOT
- Single ASIN → use fetch_external_page instead
- Skip the pre-flight reminder
- Share raw output externally (AST internal only)
