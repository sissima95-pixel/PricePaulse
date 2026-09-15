---
name: asin-pricepulse
description: ASIN PricePulse — batch fetch Amazon ASIN prices across 17 marketplaces. Generates interactive HTML report + Excel detail.
version: 1.3
agent: claude-code
---

# ASIN PricePulse — CLI for Claude Code / Kiro

## Install
```bash
git clone https://github.com/sissima95-pixel/PricePulse.git && cd PricePulse && pip install -e .
```

## Usage
```bash
pricepulse --markets US --input asins.xlsx --output ./out --title "Category Report" --subtitle "2026.5-7" --yes
```

## Key flags
- `--markets` — US,UK,DE,FR,IT,ES,NL,SE,PL,JP,CA,AU,AE,SA,SG,IN,MX,BR,TR
- `--input` — .xlsx/.csv/.txt (auto-detects Search Rank, Purchase Rank, Keyword Searches columns)
- `--delay 1.0 --workers 3` — for large batches
- `--title` / `--subtitle` — report header text

## Output
- `asin_detail_*.xlsx` — styled Excel
- `竞品ASIN分析报告_*.html` — interactive report with price tiers, brand charts, volume chart, detail table
