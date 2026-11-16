
from __future__ import annotations

import difflib
import html
import io
import re
from copy import copy
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.utils import get_column_letter

FILL_CHANGED = PatternFill("solid", fgColor="FFF2CC")
FILL_ADDED = PatternFill("solid", fgColor="C6EFCE")
FILL_REMOVED = PatternFill("solid", fgColor="FFC7CE")
FILL_HEADER = PatternFill("solid", fgColor="1F4E78")
FONT_HEADER = Font(color="FFFFFF", bold=True)
FONT_REMOVED = Font(color="C00000", bold=True, strike=True)


def printable(value):
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
    return str(value)


def comparable(value):
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    return value


def _tokens(text):
    return re.findall(r"\s+|[A-Za-z0-9_.]+|[^\w\s]", str(text))


def make_rich_diff(old_value, new_value):
    old_t = _tokens(printable(old_value))
    new_t = _tokens(printable(new_value))
    sm = difflib.SequenceMatcher(a=old_t, b=new_t, autojunk=False)
    rich = CellRichText()
    normal = InlineFont(color="000000")
    added = InlineFont(color="0000FF", b=True)
    removed = InlineFont(color="C00000", b=True, strike=True)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            text = "".join(new_t[j1:j2])
            if text:
                rich.append(TextBlock(normal, text))
        elif tag == "insert":
            text = "".join(new_t[j1:j2])
            if text:
                rich.append(TextBlock(added, text))
        elif tag == "delete":
            text = "".join(old_t[i1:i2])
            if text:
                rich.append(TextBlock(removed, text))
        elif tag == "replace":
            old_text = "".join(old_t[i1:i2])
            new_text = "".join(new_t[j1:j2])
            if old_text:
                rich.append(TextBlock(removed, old_text))
            if new_text:
                rich.append(TextBlock(added, new_text))
    return rich


def make_html_inline_diff(old_value, new_value):
    old_t = _tokens(printable(old_value))
    new_t = _tokens(printable(new_value))
    sm = difflib.SequenceMatcher(a=old_t, b=new_t, autojunk=False)
    parts = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        old_text = html.escape("".join(old_t[i1:i2]))
        new_text = html.escape("".join(new_t[j1:j2]))
        if tag == "equal":
            parts.append(new_text)
        elif tag == "insert":
            parts.append(f"<ins>{new_text}</ins>")
        elif tag == "delete":
            parts.append(f"<del>{old_text}</del>")
        else:
            if old_text:
                parts.append(f"<del>{old_text}</del>")
            if new_text:
                parts.append(f"<ins>{new_text}</ins>")
    return "".join(parts)


def add_summary_sheet(wb, records, new_name, old_name):
    title = "Diff_Summary"
    if title in wb.sheetnames:
        del wb[title]
    ws = wb.create_sheet(title, 0)
    headers = ["Status", "Sheet", "Cell", "Row", "Column", "Column Header", "OLD value", "NEW value"]
    ws.append(["Excel color-diff report"])
    ws.append([f"NEW: {new_name}"])
    ws.append([f"OLD: {old_name}"])
    ws.append([])
    ws.append(headers)
    for cell in ws[5]:
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = Alignment(horizontal="center")
    for record in records:
        ws.append([
            record["status"], record["sheet"], record["cell"], record["row"],
            record["column"], record["header"], record["old"], record["new"]
        ])
        fill = FILL_CHANGED if record["status"] == "CHANGED" else (
            FILL_ADDED if record["status"] == "ADDED" else FILL_REMOVED
        )
        for cell in ws[ws.max_row]:
            cell.fill = copy(fill)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        if record["status"] == "REMOVED":
            ws.cell(ws.max_row, 7).font = copy(FONT_REMOVED)
    for i, width in enumerate([12, 22, 10, 9, 9, 30, 45, 45], 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A6"
    ws.auto_filter.ref = f"A5:H{max(5, ws.max_row)}"


def make_html_report(records, new_name, old_name):
    counts = {key: sum(r["status"] == key for r in records) for key in ("CHANGED", "ADDED", "REMOVED")}
    rows = []
    for r in records:
        cls = r["status"].lower()
        detail = (
            make_html_inline_diff(r["old"], r["new"]) if r["status"] == "CHANGED"
            else f"<ins>{html.escape(r['new'])}</ins>" if r["status"] == "ADDED"
            else f"<del>{html.escape(r['old'])}</del>"
        )
        rows.append(
            "<tr class='%s'><td>%s</td><td>%s</td><td>%s</td><td>%s</td>"
            "<td>%s</td><td class='value'>%s</td><td class='value'>%s</td><td class='value'>%s</td></tr>" % (
                cls, html.escape(r["status"]), html.escape(r["sheet"]),
                html.escape(r["cell"]), html.escape(r["header"]),
                html.escape(str(r["row"])), html.escape(r["old"]),
                html.escape(r["new"]), detail
            )
        )
    if not rows:
        rows.append("<tr><td colspan='8' class='none'>No value differences found.</td></tr>")
    return f"""<!doctype html>
<html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>
<title>Excel Diff Report</title>
<style>
body{{font-family:Arial,sans-serif;margin:24px;background:#f5f7fa;color:#172033}}
.card{{background:white;border-radius:10px;padding:18px;box-shadow:0 2px 10px #0001;margin-bottom:16px}}
h1{{margin:0 0 12px;color:#1f4e78}} .files{{line-height:1.7;word-break:break-all}}
.stats{{display:flex;gap:12px;flex-wrap:wrap}} .stat{{padding:10px 16px;border-radius:8px;font-weight:700}}
.changed{{background:#fff2cc}} .added{{background:#c6efce}} .removed{{background:#ffc7ce}}
.controls{{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:12px}} input,select{{padding:9px;border:1px solid #aab4c3;border-radius:6px}}
table{{border-collapse:collapse;width:100%;background:white;font-size:13px}} th{{position:sticky;top:0;background:#1f4e78;color:white;text-align:left;padding:9px}}
td{{border:1px solid #d8dee8;padding:7px;vertical-align:top}} td.value{{white-space:pre-wrap;word-break:break-word;max-width:360px}}
tr.changed{{background:#fff9e6}} tr.added{{background:#eaf7ed}} tr.removed{{background:#fdebec}}
tr.removed td:nth-child(6){{color:#c00000;text-decoration:line-through}}
ins{{background:#9ff4df;color:#0000ff;font-weight:700;text-decoration:none}}
del{{background:#ffc7ce;color:#c00000;font-weight:700;text-decoration:line-through}}
.none{{text-align:center;padding:24px}}
</style></head><body>
<div class='card'><h1>Excel Color Diff Report</h1><div class='files'><b>NEW:</b> {html.escape(new_name)}<br><b>OLD:</b> {html.escape(old_name)}</div></div>
<div class='card stats'><div class='stat changed'>Changed: {counts['CHANGED']}</div><div class='stat added'>Added: {counts['ADDED']}</div><div class='stat removed'>Removed: {counts['REMOVED']}</div><div class='stat'>Total: {len(records)}</div></div>
<div class='card'><div class='controls'><input id='q' placeholder='Search sheet, cell, header or value'><select id='status'><option value=''>All statuses</option><option>CHANGED</option><option>ADDED</option><option>REMOVED</option></select></div>
<table id='diff'><thead><tr><th>Status</th><th>Sheet</th><th>Cell</th><th>Column Header</th><th>Row</th><th>OLD value</th><th>NEW value</th><th>Inline change detail</th></tr></thead><tbody>{''.join(rows)}</tbody></table></div>
<script>
const q=document.getElementById('q'), st=document.getElementById('status');
function filter(){{let text=q.value.toLowerCase(),status=st.value;document.querySelectorAll('#diff tbody tr').forEach(r=>{{let ok=r.innerText.toLowerCase().includes(text)&&(!status||r.cells[0]?.innerText===status);r.style.display=ok?'':'none';}})}}
q.addEventListener('input',filter);st.addEventListener('change',filter);
</script></body></html>"""


def compare_oit_workbooks(new_bytes: bytes, old_bytes: bytes, new_name: str, old_name: str, progress_callback=None):
    """Return (Excel bytes, HTML string, records, counts) for uploaded NEW and OLD workbooks."""
    if progress_callback:
        progress_callback(3, "Loading NEW OIT workbook...")
    new_wb = load_workbook(io.BytesIO(new_bytes), data_only=False)
    if progress_callback:
        progress_callback(8, "Loading OLD OIT workbook...")
    old_wb = load_workbook(io.BytesIO(old_bytes), data_only=False)

    sheet_names = list(dict.fromkeys(new_wb.sheetnames + old_wb.sheetnames))
    dimensions = {}
    total = 0
    for name in sheet_names:
        nws = new_wb[name] if name in new_wb.sheetnames else None
        ows = old_wb[name] if name in old_wb.sheetnames else None
        max_row = max(nws.max_row if nws else 0, ows.max_row if ows else 0)
        max_col = max(nws.max_column if nws else 0, ows.max_column if ows else 0)
        dimensions[name] = (max_row, max_col)
        total += max_row * max_col

    records = []
    processed = 0
    for sheet_name in sheet_names:
        nws = new_wb[sheet_name] if sheet_name in new_wb.sheetnames else None
        ows = old_wb[sheet_name] if sheet_name in old_wb.sheetnames else None
        max_row, max_col = dimensions[sheet_name]
        header_row = 2 if max_row >= 2 else 1
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                new_cell = nws.cell(r, c) if nws else None
                old_cell = ows.cell(r, c) if ows else None
                nv = new_cell.value if new_cell else None
                ov = old_cell.value if old_cell else None
                if comparable(nv) != comparable(ov):
                    if comparable(ov) == "" and comparable(nv) != "":
                        status, fill = "ADDED", FILL_ADDED
                    elif comparable(nv) == "" and comparable(ov) != "":
                        status, fill = "REMOVED", FILL_REMOVED
                    else:
                        status, fill = "CHANGED", FILL_CHANGED
                    if new_cell is not None:
                        new_cell.fill = copy(fill)
                        if status == "CHANGED" and isinstance(ov, str) and isinstance(nv, str):
                            new_cell.value = make_rich_diff(ov, nv)
                            new_cell.alignment = Alignment(
                                horizontal=new_cell.alignment.horizontal,
                                vertical=new_cell.alignment.vertical or "top",
                                wrap_text=True,
                            )
                        elif status == "CHANGED":
                            changed_font = copy(new_cell.font)
                            changed_font.color = "0000FF"
                            changed_font.bold = True
                            new_cell.font = changed_font
                        elif status == "ADDED":
                            added_font = copy(new_cell.font)
                            added_font.color = "0000FF"
                            added_font.bold = True
                            new_cell.font = added_font
                        elif status == "REMOVED" and old_cell is not None:
                            new_cell.value = ov
                            base_font = copy(old_cell.font)
                            base_font.color = "C00000"
                            base_font.bold = True
                            base_font.strike = True
                            new_cell.font = base_font
                    coord = f"{get_column_letter(c)}{r}"
                    header_source = nws if nws else ows
                    header = header_source.cell(header_row, c).value if (header_source and r != header_row) else nv
                    records.append({
                        "status": status, "sheet": sheet_name, "cell": coord,
                        "row": r, "column": c, "header": printable(header),
                        "old": printable(ov), "new": printable(nv),
                    })
                processed += 1
                if progress_callback and (processed % max(1, total // 100) == 0 or processed == total):
                    pct = 10 + int(processed / max(1, total) * 70)
                    progress_callback(min(80, pct), f"Comparing workbook cells: {processed:,} / {total:,}")

    if progress_callback:
        progress_callback(84, "Adding Diff_Summary worksheet...")
    add_summary_sheet(new_wb, records, new_name, old_name)

    if progress_callback:
        progress_callback(90, "Generating HTML diff report...")
    html_report = make_html_report(records, new_name, old_name)

    if progress_callback:
        progress_callback(94, "Saving Excel diff workbook...")
    excel_buffer = io.BytesIO()
    new_wb.save(excel_buffer)
    new_wb.close()
    old_wb.close()

    counts = {key: sum(r["status"] == key for r in records) for key in ("CHANGED", "ADDED", "REMOVED")}
    if progress_callback:
        progress_callback(100, "OIT comparison completed.")
    return excel_buffer.getvalue(), html_report, records, counts
