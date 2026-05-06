"""
Shared utilities for the automated CDR pipeline.
"""

from docx import Document as _Document

_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def detect_compliance_heading(template_path: str) -> str:
    """
    Auto-detect the heading in the template that precedes the largest
    multi-column table (>= 3 columns, most rows).  This identifies the
    section used for the compliance checklist (admin + data tables).
    Returns the heading text, or "" if detection fails.
    """
    try:
        W = f"{{{_NS}}}"
        doc = _Document(template_path)
        h_ids = {s.style_id for s in doc.styles
                 if s.name and s.name.startswith("Heading")}
        cur_h = ""
        best_rows = 0
        result = ""
        for e in doc.element.body:
            if e.tag == f"{W}p":
                pPr = e.find(f"{W}pPr")
                if pPr is None:
                    continue
                ps = pPr.find(f"{W}pStyle")
                if ps is not None and ps.get(f"{W}val", "") in h_ids:
                    cur_h = "".join(t.text or "" for t in e.iter(f"{W}t")).strip()
            elif e.tag == f"{W}tbl" and cur_h:
                rows = e.findall(f"{W}tr")
                if rows:
                    r0_cells = rows[0].findall(f"{W}tc")
                    if len(r0_cells) >= 3 and len(rows) > best_rows:
                        best_rows = len(rows)
                        result = cur_h
        return result
    except Exception:
        return ""
