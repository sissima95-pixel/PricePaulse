---
name: asin-pricepulse
description: ASIN PricePulse (alias PP / pp) — batch fetch Amazon ASIN prices across 17 marketplaces; interactive HTML report + Excel detail. Activate on "pp", "pricepulse", "竞品价格", "ASIN 价格".
version: 1.4
agent: claude-code
---

# ASIN PricePulse (PP) — CLI for Claude Code / Kiro / Codex

Aliases: PricePulse = pricepulse = PP = pp.

## Install (from zip)
Unzip `PricePulse-vX.zip`, then in the `PricePulse/` folder: `pip install -e .`
Verify: `pricepulse --version`

## Usage
```bash
pricepulse --markets US --input asins.xlsx --output ./out --title "US Personal Fans" --subtitle "2026.5.1-7.31" --yes
```
- `--markets` US,UK,DE,FR,IT,ES,NL,SE,PL,JP,CA,AU,AE,SA,SG,IN,MX,BR,TR (EU5 = DE,FR,IT,ES,UK)
- `--input` .xlsx/.csv/.txt; auto-detects Search Rank / Purchase Rank / Keyword Searches columns
- `--asins B0XXX,B0YYY` inline instead of a file
- 300+ ASINs → add `--delay 1.0 --workers 3`

## Output
- `asin_detail_*.xlsx`
- `竞品ASIN分析报告_*.html`: price tiers · distribution · brand share donuts · brand-by-tier (Top 5, expandable ASINs) · Search Volume Top 15 with brand bubble filter · detail table (sort / filter / Copy ASINs / Export Excel) · EN+CN insights A/B/C

## Notes
- Pre-flight: prices are a local-visitor list-price snapshot; VPN-independent; AST internal only.
- Custom tiers ("0-40, 40-80, 80+"): patch `pricepulse.reporter._price_bands` before `write_html`.
