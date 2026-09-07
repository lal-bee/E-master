"""P0-09 静态 HTML 展示：把 ValidationDisplay 渲染为可本地打开的页面。"""

from __future__ import annotations

import html
import json

from ui.display import ValidationDisplay


def render_html(display: ValidationDisplay) -> str:
    """渲染无第三方依赖的 HTML 页面。"""
    summary = display.summary
    issues_json = json.dumps(
        [
            {
                "file": display.file_name,
                "seq": error.sequence_number,
                "sheet": error.sheet_name,
                "row": error.row_number,
                "column_letter": error.column_letter,
                "column_number": error.column_number,
                "field": error.field_name,
                "cell": error.cell_reference,
                "type": error.error_type,
                "code": error.error_code,
                "actual": error.actual_value,
                "expected": error.expected_value,
                "reason": error.reason,
                "source": error.source,
            }
            for error in display.errors
        ],
        ensure_ascii=False,
    ).replace("</", "<\\/")

    select_options = _render_select_options(display)
    return _TEMPLATE.format(
        file_name=html.escape(display.file_name),
        total=summary.total_checks,
        passed=summary.pass_count,
        errors=summary.error_count,
        skipped=summary.skipped_empty_row_count,
        format_errors=summary.format_error_count,
        data_not_found=summary.data_not_found_count,
        data_error=summary.data_error_count,
        type_options=select_options["type"],
        sheet_options=select_options["sheet"],
        field_options=select_options["field"],
        code_options=select_options["code"],
        issues_json=issues_json,
        all_passed_message=(
            '<div class="pass-message">全部通过，无错误。</div>'
            if display.all_passed
            else ""
        ),
    )


def _render_select_options(display: ValidationDisplay) -> dict[str, str]:
    def render(label: str, options: tuple[str, ...]) -> str:
        values = "".join(
            f'<option value="{html.escape(value, quote=True)}">{html.escape(value)}</option>'
            for value in options
        )
        return f'<option value="">全部{label}</option>{values}'

    return {
        "type": render("错误类型", display.error_type_options),
        "sheet": render("Sheet", display.sheet_options),
        "field": render("字段", display.field_options),
        "code": render("错误编码", display.error_code_options),
    }


_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>核验结果 - {file_name}</title>
<style>
body{{font-family:"Microsoft YaHei",Arial,sans-serif;margin:24px;color:#222;}}
h1{{font-size:20px;}}
.file{{color:#555;margin:8px 0 16px;}}
.cards{{display:flex;gap:12px;flex-wrap:wrap;}}
.card{{border:1px solid #ddd;border-radius:6px;padding:10px 16px;min-width:120px;}}
.card .num{{font-size:24px;font-weight:bold;}}
.filters{{display:flex;gap:12px;flex-wrap:wrap;margin:18px 0;}}
table{{border-collapse:collapse;width:100%;font-size:13px;}}
th,td{{border:1px solid #ddd;padding:6px 8px;text-align:left;white-space:nowrap;}}
th{{background:#f5f5f5;}}
.reason{{white-space:normal;min-width:220px;}}
tr.clickable{{cursor:pointer;}}
.pass-message{{color:#1a7f37;font-size:16px;margin:16px 0;}}
#detail{{border:1px solid #ddd;border-radius:6px;padding:12px 16px;margin-top:16px;display:none;}}
#detail h3{{margin-top:0;}}
</style>
</head>
<body>
<h1>Excel 内容比对与数据质量核验系统</h1>
<div class="file">核验文件：{file_name}</div>
<div class="cards">
  <div class="card"><div>总数</div><div class="num">{total}</div></div>
  <div class="card"><div>通过</div><div class="num">{passed}</div></div>
  <div class="card"><div>错误</div><div class="num">{errors}</div></div>
  <div class="card"><div>跳过空行</div><div class="num">{skipped}</div></div>
</div>
<h2>错误分类</h2>
<div class="cards">
  <div class="card"><div>格式错误</div><div class="num">{format_errors}</div></div>
  <div class="card"><div>数据不存在</div><div class="num">{data_not_found}</div></div>
  <div class="card"><div>数据错误</div><div class="num">{data_error}</div></div>
</div>
{all_passed_message}
<h2>错误明细</h2>
<div class="filters">
  <label>错误类型 <select id="filterType">{type_options}</select></label>
  <label>Sheet <select id="filterSheet">{sheet_options}</select></label>
  <label>字段 <select id="filterField">{field_options}</select></label>
  <label>错误编码 <select id="filterCode">{code_options}</select></label>
</div>
<table id="issueTable">
<thead>
<tr>
<th>序号</th><th>Sheet</th><th>行号</th><th>列</th><th>字段</th>
<th>单元格</th><th>错误类型</th><th>错误编码</th><th>实际值</th>
<th>期望值</th><th>错误原因</th><th>来源</th>
</tr>
</thead>
<tbody></tbody>
</table>
<div id="detail"></div>
<script>
const issues = {issues_json};
const filters = ["type", "sheet", "field", "code"];
function matches(item) {{
  return filters.every(function(key) {{
    const value = document.getElementById("filter" + key.charAt(0).toUpperCase() + key.slice(1)).value;
    return value === "" || item[key] === value;
  }});
}}
function renderTable() {{
  const rows = issues.filter(matches);
  const tbody = document.querySelector("#issueTable tbody");
  tbody.innerHTML = "";
  rows.forEach(function(item) {{
    const tr = document.createElement("tr");
    tr.className = "clickable";
    tr.dataset.seq = item.seq;
    [item.seq, item.sheet, item.row, item.column_letter, item.field, item.cell,
     item.type, item.code, item.actual, item.expected, item.reason, item.source]
      .forEach(function(value) {{
        const td = document.createElement("td");
        td.textContent = value === null || value === undefined ? "" : String(value);
        if (value === item.reason) td.className = "reason";
        tr.appendChild(td);
      }});
    tr.addEventListener("click", function() {{ showDetail(item); }});
    tbody.appendChild(tr);
  }});
}}
function showDetail(item) {{
  const detail = document.getElementById("detail");
  detail.style.display = "block";
  detail.innerHTML =
    "<h3>错误详情：<strong>" + item.sheet + "!" + item.cell + "</strong></h3>" +
    "<p>文件：" + item.file + "</p>" +
    "<p>Sheet：" + item.sheet + "</p>" +
    "<p>物理行号：" + item.row + "</p>" +
    "<p>列号：" + item.column_number + "；列字母：" + item.column_letter +
    "；字段：" + item.field + "；单元格：" + item.cell + "</p>" +
    "<p>实际值：" + item.actual + "</p>" +
    "<p>期望值：" + item.expected + "</p>" +
    "<p>错误类型：" + item.type + "；错误编码：" + item.code + "</p>" +
    "<p>错误原因：" + item.reason + "</p>" +
    "<p>来源模块：" + item.source + "</p>";
}}
filters.forEach(function(key) {{
  document.getElementById("filter" + key.charAt(0).toUpperCase() + key.slice(1))
    .addEventListener("change", renderTable);
}});
renderTable();
</script>
</body>
</html>
"""
