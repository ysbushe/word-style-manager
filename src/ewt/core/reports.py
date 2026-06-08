"""HTML reports for batch style operations."""

from __future__ import annotations

from datetime import datetime
from html import escape
from pathlib import Path


def generate_html_report(report_path, title, rows, summary=None):
    report_path = Path(report_path)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    summary = summary or {}
    summary_html = "".join(
        f"<div class='metric'><span>{escape(str(key))}</span><strong>{escape(str(value))}</strong></div>"
        for key, value in summary.items()
    )
    table_rows = "\n".join(
        "<tr>" + "".join(f"<td>{escape(str(cell))}</td>" for cell in row) + "</tr>"
        for row in rows
    )
    html = f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>{escape(title)}</title>
<style>
*{{box-sizing:border-box}}
body{{font-family:"Microsoft YaHei UI",Arial,sans-serif;background:#f4f7fb;color:#1d2939;margin:0;padding:36px}}
.page{{max-width:1100px;margin:auto}}
.head{{background:#155e75;color:white;padding:26px 30px;border-radius:10px 10px 0 0}}
h1{{font-size:24px;margin:0}}.time{{opacity:.8;margin-top:7px;font-size:13px}}
.card{{background:white;border:1px solid #dfe6ee;padding:24px 28px;margin-bottom:18px}}
.metrics{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px}}
.metric{{background:#f7fafc;border-left:4px solid #0e7490;padding:12px}}
.metric span{{display:block;color:#667085;font-size:12px}}.metric strong{{font-size:18px}}
table{{border-collapse:collapse;width:100%;font-size:13px}}
td{{padding:10px 12px;border-bottom:1px solid #e5eaf0;vertical-align:top}}
tr:nth-child(even){{background:#f8fafc}}td:first-child{{font-weight:600;width:190px;color:#344054}}
.foot{{color:#667085;font-size:12px;text-align:center}}
</style>
</head>
<body><div class="page">
<div class="head"><h1>{escape(title)}</h1><div class="time">生成时间：{datetime.now():%Y-%m-%d %H:%M:%S}</div></div>
<div class="card metrics">{summary_html}</div>
<div class="card"><table><tbody>{table_rows}</tbody></table></div>
<div class="foot">Word 样式管理器</div>
</div></body></html>"""
    report_path.write_text(html, encoding="utf-8")
    return str(report_path)
