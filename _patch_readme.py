s = open('README.md', encoding='utf-8').read()
old = "| 版本 | 日期 | 变更 |\n|------|------|------|\n"
new = old + (
"| 1.4.0 | 2026-09-16 | 品牌份额环形图(by Search Volume / by ASIN Count);Brand Share by Price Tier(每档 Top 5,彩色圆点表格,可展开 ASIN 明细子表);"
"Insight 改为专业格式 A/B/C 三段(英文 + 中文斜体,无 emoji),C 段逐个分析 Top 3 品牌;品牌气泡云移至 Search Volume 区块;Copy ASINs / Export Excel 按钮;"
"品牌识别新增 po-brand + Brand Name 表格模式(最高优先级),Store URL 降级,自动重试空品牌;Unknown 统一排最后;价格段改为分位数(tercile)智能分档,Python/JS 边界同步;别名 PP |\n"
"| 1.3.0 | 2026-08-11 | PriceLens 改名 PricePulse(包名 / CLI / skill 全部更新) |\n"
)
assert old in s
s = s.replace(old, new, 1)
# also refresh the report-content section headline list
s = s.replace("### HTML 报告包含", "### HTML 报告包含(v1.4)", 1)
open('README.md', 'w', encoding='utf-8').write(s)
print("README updated")
