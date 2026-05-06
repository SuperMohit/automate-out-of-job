#!/usr/bin/env python3
"""
generate_audit_report.py — Generic audit report renderer.

Reads a findings.json file and produces a multi-sheet styled Excel workbook
matching the audit_report.xlsx reference style.

Usage:
    python3 generate_audit_report.py findings.json
    python3 generate_audit_report.py findings.json --output my_report.xlsx
    python3 generate_audit_report.py --empty-template empty_template.xlsx

Schema for findings.json:
    {
      "metadata": {
        "repo_name": "<repo>",
        "audit_date": "YYYY-MM-DD",
        "auditor": "Claude <model>"
      },
      "findings": [
        {"id":"S1","category":"Security","severity":"CRITICAL","priority":"P0",
         "title":"...","files":"path/to/file:line","description":"...","fix":"..."}
      ]
    }

Categories supported (any other string is rendered with default colour):
  Security, Production, Container, CI/CD, Code Quality,
  Observability, Azure, Additional, SOC2

Severities: CRITICAL, HIGH, MEDIUM, LOW
Priorities: P0, P1, P2, P3

For SOC2 findings, use IDs starting with "SOC2-<TSC>" (e.g. SOC2-CC6.1).
The renderer adds a Trust Service Criteria column on the SOC2 sheet,
extracting the TSC from the prefix between "SOC2-" and the first "." or end.
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.stderr.write("ERROR: openpyxl is required. Install with: pip install openpyxl\n")
    sys.exit(1)


# ── colour palette (matches audit_report.xlsx reference) ──────────────────
CLR_HEADER_BG   = "1F3864"
CLR_HEADER_FG   = "FFFFFF"
CLR_CRITICAL_BG = "C00000"
CLR_HIGH_BG     = "FF6600"
CLR_MEDIUM_BG   = "FFCC00"
CLR_LOW_BG      = "92D050"
CLR_LOW_FG      = "000000"
CLR_WHITE_FG    = "FFFFFF"
CLR_SUBTITLE_BG = "D6E4F0"
CLR_SUBTITLE_FG = "444444"

CATEGORY_COLORS = {
    "Security":      "FFD7D7",
    "Production":    "DDEEFF",
    "Container":     "E8F5E9",
    "CI/CD":         "FFF3E0",
    "Code Quality":  "F3E5F5",
    "Observability": "E0F7FA",
    "Azure":         "FFF8E1",
    "Additional":    "F5F5F5",
    "SOC2":          "EDE7F6",
}

PRIORITY_ROADMAP_COLORS = {
    "P0": "C00000",
    "P1": "FF6600",
    "P2": "FFCC00",
    "P3": "92D050",
}

PRIORITY_ORDER = {"P0": 0, "P1": 1, "P2": 2, "P3": 3}
SEV_ORDER      = {"CRITICAL": 0, "HIGH": 1, "MEDIUM": 2, "LOW": 3}

# Order of category sheets in the workbook
CATEGORY_SHEETS = [
    ("Security",      "Security"),
    ("Production",    "Production"),
    ("Container",     "Container"),
    ("CI/CD",         "CI-CD"),
    ("Code Quality",  "Code Quality"),
    ("Observability", "Observability"),
    ("Azure",         "Azure"),
    ("Additional",    "Additional"),
    ("SOC2",          "SOC2"),
]

# SOC2 Trust Service Criteria labels
TSC_MAP = {
    "CC6": "CC6 — Logical & Physical Access",
    "CC7": "CC7 — System Operations",
    "CC8": "CC8 — Change Management",
    "CC9": "CC9 — Risk Mitigation",
    "A1":  "A1 — Availability",
    "PI1": "PI1 — Processing Integrity",
    "C1":  "C1 — Confidentiality",
    "P1":  "P — Privacy",
}


# ── style helpers ─────────────────────────────────────────────────────────
def _fill(c):    return PatternFill("solid", fgColor=c)
def _font(bold=False, color="000000", size=11):
    return Font(bold=bold, color=color, size=size, name="Calibri")
def _align(wrap=True, v="top", h="left"):
    return Alignment(wrap_text=wrap, vertical=v, horizontal=h)
def _border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def _header_style(cell, text, size=11):
    cell.value = text
    cell.fill = _fill(CLR_HEADER_BG)
    cell.font = _font(bold=True, color=CLR_HEADER_FG, size=size)
    cell.alignment = _align(v="center", h="center")
    cell.border = _border()


def _severity_style(cell, severity):
    bg = {"CRITICAL": CLR_CRITICAL_BG, "HIGH": CLR_HIGH_BG,
          "MEDIUM": CLR_MEDIUM_BG, "LOW": CLR_LOW_BG}.get(severity, "FFFFFF")
    fg = CLR_WHITE_FG if severity in ("CRITICAL", "HIGH") else CLR_LOW_FG
    cell.value = severity
    cell.fill = _fill(bg)
    cell.font = _font(bold=True, color=fg)
    cell.alignment = _align(v="center", h="center")
    cell.border = _border()


def _data_style(cell, text, bold=False, bg=None, h="left"):
    cell.value = text
    if bg:
        cell.fill = _fill(bg)
    cell.font = _font(bold=bold)
    cell.alignment = _align(h=h)
    cell.border = _border()


def _add_findings_sheet(wb, sheet_name, findings):
    ws = wb.create_sheet(sheet_name)
    headers = ["ID", "Category", "Severity", "Priority", "Title",
               "Affected File(s)", "Description", "Recommended Fix"]
    for col, h in enumerate(headers, 1):
        _header_style(ws.cell(row=1, column=col), h)
    ws.row_dimensions[1].height = 28
    for col, w in enumerate([6, 14, 10, 8, 42, 38, 60, 52], 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    for row_idx, f in enumerate(findings, 2):
        cat_bg = CATEGORY_COLORS.get(f["category"], "F5F5F5")
        _data_style(ws.cell(row=row_idx, column=1), f.get("id", ""), h="center")
        _data_style(ws.cell(row=row_idx, column=2), f.get("category", ""), bg=cat_bg)
        _severity_style(ws.cell(row=row_idx, column=3), f.get("severity", ""))
        _data_style(ws.cell(row=row_idx, column=4), f.get("priority", ""), h="center")
        _data_style(ws.cell(row=row_idx, column=5), f.get("title", ""), bold=True)
        _data_style(ws.cell(row=row_idx, column=6), f.get("files", ""))
        _data_style(ws.cell(row=row_idx, column=7), f.get("description", ""))
        _data_style(ws.cell(row=row_idx, column=8), f.get("fix", ""))
        ws.row_dimensions[row_idx].height = 60

    ws.freeze_panes = "A2"
    return ws


def _add_soc2_tsc_column(ws, findings):
    """Add the 'Trust Service Criteria' column on the SOC2 sheet."""
    col_i = 9
    ws.column_dimensions[get_column_letter(col_i)].width = 30
    _header_style(ws.cell(row=1, column=col_i), "Trust Service Criteria (TSC)")
    for row_idx, f in enumerate(findings, 2):
        id_ = f.get("id", "")
        # Extract TSC key from "SOC2-CC6.1" -> "CC6"
        tsc_key = id_.replace("SOC2-", "").split(".")[0]
        tsc_label = TSC_MAP.get(tsc_key, tsc_key)
        _data_style(ws.cell(row=row_idx, column=col_i), tsc_label, bg="EDE7F6")


def _build_executive_summary(wb, metadata, findings):
    sev_counts = Counter(f.get("severity", "") for f in findings)
    cat_counts = Counter(f.get("category", "") for f in findings)

    ws = wb.create_sheet("Executive Summary")
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 4
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 12

    repo_name  = metadata.get("repo_name", "<unknown repo>")
    audit_date = metadata.get("audit_date", "")
    auditor    = metadata.get("auditor", "Claude")
    total      = len(findings)

    c = ws["A1"]
    c.value = f"Repository Audit Report — {repo_name}"
    c.fill = _fill(CLR_HEADER_BG)
    c.font = _font(bold=True, color=CLR_HEADER_FG, size=14)
    c.alignment = _align(v="center", h="left")
    ws.merge_cells("A1:E1")
    ws.row_dimensions[1].height = 30

    c = ws["A2"]
    c.value = f"Audit Date: {audit_date}  |  Total Findings: {total}  |  Auditor: {auditor}"
    c.fill = _fill(CLR_SUBTITLE_BG)
    c.font = _font(color=CLR_SUBTITLE_FG, size=10)
    c.alignment = _align(v="center", h="left")
    ws.merge_cells("A2:E2")
    ws.row_dimensions[2].height = 20

    for col, txt in zip(["A", "B"], ["Severity", "Count"]):
        _header_style(ws[f"{col}4"], txt)
    for col, txt in zip(["D", "E"], ["Category", "Findings"]):
        _header_style(ws[f"{col}4"], txt)

    sev_rows = [
        ("CRITICAL", sev_counts.get("CRITICAL", 0), CLR_CRITICAL_BG, CLR_WHITE_FG),
        ("HIGH",     sev_counts.get("HIGH", 0),     CLR_HIGH_BG,     CLR_WHITE_FG),
        ("MEDIUM",   sev_counts.get("MEDIUM", 0),   CLR_MEDIUM_BG,   CLR_LOW_FG),
        ("LOW",      sev_counts.get("LOW", 0),       CLR_LOW_BG,      CLR_LOW_FG),
    ]
    for i, (sev, cnt, bg, fg) in enumerate(sev_rows, 5):
        c = ws[f"A{i}"]
        c.value = sev
        c.fill = _fill(bg)
        c.font = _font(bold=True, color=fg)
        c.alignment = _align(v="center", h="center")
        c.border = _border()
        c2 = ws[f"B{i}"]
        c2.value = cnt
        c2.alignment = _align(v="center", h="center")
        c2.border = _border()

    sorted_cats = sorted(cat_counts.items())
    for i, (cat, cnt) in enumerate(sorted_cats, 5):
        cat_bg = CATEGORY_COLORS.get(cat, "F5F5F5")
        c = ws[f"D{i}"]
        c.value = cat
        c.fill = _fill(cat_bg)
        c.alignment = _align(v="center", h="left")
        c.border = _border()
        c2 = ws[f"E{i}"]
        c2.value = cnt
        c2.alignment = _align(v="center", h="center")
        c2.border = _border()

    p_row = 5 + max(len(sorted_cats), 4) + 1
    _header_style(ws[f"A{p_row}"], "Priority")
    _header_style(ws[f"B{p_row}"], "Meaning")
    ws.merge_cells(f"B{p_row}:E{p_row}")

    priorities = [
        ("P0", "Fix immediately — critical security or data breach risk"),
        ("P1", "Fix in current sprint — production-readiness blockers"),
        ("P2", "Fix in next sprint — significant reliability/security improvements"),
        ("P3", "Backlog — quality of life, observability, Azure maturity"),
    ]
    for i, (p, meaning) in enumerate(priorities, p_row + 1):
        c = ws[f"A{i}"]
        c.value = p
        c.alignment = _align(v="center", h="center")
        c.border = _border()
        c2 = ws[f"B{i}"]
        c2.value = meaning
        c2.alignment = _align(v="center", h="left")
        c2.border = _border()
        ws.merge_cells(f"B{i}:E{i}")


def _build_priority_roadmap(wb, findings):
    roadmap = sorted(
        findings,
        key=lambda f: (
            PRIORITY_ORDER.get(f.get("priority", "P3"), 99),
            SEV_ORDER.get(f.get("severity", "LOW"), 99),
        ),
    )
    ws = wb.create_sheet("Priority Roadmap")
    headers = ["Priority", "ID", "Category", "Severity", "Title", "Recommended Fix"]
    widths  = [10, 6, 14, 10, 48, 60]
    for col, (h, w) in enumerate(zip(headers, widths), 1):
        _header_style(ws.cell(row=1, column=col), h)
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 28

    for row_idx, f in enumerate(roadmap, 2):
        pri = f.get("priority", "")
        cat = f.get("category", "")
        sev = f.get("severity", "")
        cat_bg = CATEGORY_COLORS.get(cat, "F5F5F5")

        pri_bg = PRIORITY_ROADMAP_COLORS.get(pri, "FFFFFF")
        pri_fg = CLR_WHITE_FG if pri in ("P0", "P1") else CLR_LOW_FG
        c = ws.cell(row=row_idx, column=1)
        c.value = pri
        c.fill = _fill(pri_bg)
        c.font = _font(bold=True, color=pri_fg)
        c.alignment = _align(v="center", h="center")
        c.border = _border()

        _data_style(ws.cell(row=row_idx, column=2), f.get("id", ""), h="center")
        _data_style(ws.cell(row=row_idx, column=3), cat, bg=cat_bg)
        _severity_style(ws.cell(row=row_idx, column=4), sev)
        _data_style(ws.cell(row=row_idx, column=5), f.get("title", ""), bold=True)
        _data_style(ws.cell(row=row_idx, column=6), f.get("fix", ""))
        ws.row_dimensions[row_idx].height = 60

    ws.freeze_panes = "A2"


def _validate_findings(findings):
    """Soft validation. Print warnings; do not abort."""
    valid_sev  = set(SEV_ORDER)
    valid_pri  = set(PRIORITY_ORDER)
    valid_cats = set(c for c, _ in CATEGORY_SHEETS)
    issues = []
    seen_ids = set()
    for i, f in enumerate(findings):
        loc = f"finding[{i}] id={f.get('id', '?')}"
        for k in ("id", "category", "severity", "priority", "title", "files", "description", "fix"):
            if k not in f:
                issues.append(f"{loc}: missing field '{k}'")
        if f.get("severity") and f["severity"] not in valid_sev:
            issues.append(f"{loc}: invalid severity '{f['severity']}'")
        if f.get("priority") and f["priority"] not in valid_pri:
            issues.append(f"{loc}: invalid priority '{f['priority']}'")
        if f.get("category") and f["category"] not in valid_cats:
            issues.append(f"{loc}: unknown category '{f['category']}' "
                          f"(no per-category sheet will be created)")
        fid = f.get("id")
        if fid in seen_ids:
            issues.append(f"{loc}: duplicate id '{fid}'")
        seen_ids.add(fid)
    if issues:
        sys.stderr.write("Validation warnings:\n")
        for msg in issues:
            sys.stderr.write(f"  - {msg}\n")


def render(findings_data, output_path):
    metadata = findings_data.get("metadata", {})
    findings = findings_data.get("findings", [])

    _validate_findings(findings)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _build_executive_summary(wb, metadata, findings)
    _add_findings_sheet(wb, "All Findings", findings)

    for cat_name, sheet_name in CATEGORY_SHEETS:
        cat_rows = [f for f in findings if f.get("category") == cat_name]
        if not cat_rows:
            continue
        ws_cat = _add_findings_sheet(wb, sheet_name, cat_rows)
        if sheet_name == "SOC2":
            _add_soc2_tsc_column(ws_cat, cat_rows)

    _build_priority_roadmap(wb, findings)

    wb.save(output_path)
    return wb


def render_empty_template(output_path):
    """Produce an empty workbook with all sheets and headers but no rows."""
    placeholder = {
        "metadata": {
            "repo_name": "<repo-name>",
            "audit_date": "YYYY-MM-DD",
            "auditor": "Claude <model>",
        },
        "findings": [],
    }
    return render(placeholder, output_path)


def main():
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    p.add_argument("findings_json", nargs="?",
                   help="Path to findings.json")
    p.add_argument("-o", "--output", default=None,
                   help="Output xlsx path (default: audit_report_<repo_name>.xlsx)")
    p.add_argument("--empty-template", metavar="PATH",
                   help="Generate an empty styled template (no findings) and exit")
    args = p.parse_args()

    if args.empty_template:
        render_empty_template(args.empty_template)
        print(f"Wrote empty template: {args.empty_template}")
        return

    if not args.findings_json:
        p.error("findings_json is required (or use --empty-template)")

    findings_path = Path(args.findings_json)
    if not findings_path.exists():
        sys.stderr.write(f"ERROR: {findings_path} not found\n")
        sys.exit(1)

    with open(findings_path) as f:
        data = json.load(f)

    if args.output:
        output_path = args.output
    else:
        repo = data.get("metadata", {}).get("repo_name", "report")
        safe = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in repo)
        output_path = f"audit_report_{safe}.xlsx"

    wb = render(data, output_path)
    sheets = wb.sheetnames
    total = len(data.get("findings", []))
    print(f"Saved: {output_path}  ({total} findings across {len(sheets)} sheets)")


if __name__ == "__main__":
    main()
