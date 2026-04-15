"""
post_processor.py
-----------------
Post-processing passes applied to the filled output document before saving.
All behaviour is driven by the "post_processing" block in prompts.json —
nothing is hardcoded in this module.
"""

import json
import os

from docx import Document
from docx.oxml.ns import qn

_PROMPTS_FILE = os.path.join(os.path.dirname(__file__), "..", "prompts.json")
with open(_PROMPTS_FILE, "r", encoding="utf-8") as _f:
    _CFG: dict = json.load(_f).get("post_processing", {})

_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _para_style_id(elem) -> str:
    pPr = elem.find(f"{{{_NS}}}pPr")
    if pPr is None:
        return ""
    pStyle = pPr.find(f"{{{_NS}}}pStyle")
    if pStyle is None:
        return ""
    return pStyle.get(f"{{{_NS}}}val", "")


def _para_text(elem) -> str:
    return "".join(t.text or "" for t in elem.iter(f"{{{_NS}}}t")).strip()


def _heading_style_ids(doc: Document) -> set:
    return {s.style_id for s in doc.styles if s.name and s.name.startswith("Heading")}


# ---------------------------------------------------------------------------
# Pass 1 — Remove paragraphs with specific styles (e.g. "Guidance")
# ---------------------------------------------------------------------------

def remove_styled_paragraphs(doc: Document) -> None:
    """Remove all body-level paragraphs whose style name is in the configured list."""
    targets: list[str] = _CFG.get("styles_to_remove", [])
    if not targets:
        return

    id_to_name = {s.style_id: (s.name or "") for s in doc.styles}
    body = doc.element.body
    to_remove = [
        e for e in list(body)
        if e.tag == f"{{{_NS}}}p" and id_to_name.get(_para_style_id(e), "") in targets
    ]
    for e in to_remove:
        body.remove(e)


# ---------------------------------------------------------------------------
# Pass 2 — Remove template instruction bracket markers  (< / >)
# ---------------------------------------------------------------------------

def remove_template_instructions(doc: Document) -> None:
    """
    Remove body-level paragraphs whose full text is exactly '<' or '>'.
    These are the opening/closing markers of template instruction blocks.
    """
    if not _CFG.get("remove_template_instructions", False):
        return

    body = doc.element.body
    to_remove = [
        e for e in list(body)
        if e.tag == f"{{{_NS}}}p" and _para_text(e) in ("<", ">")
    ]
    for e in to_remove:
        body.remove(e)


# ---------------------------------------------------------------------------
# Pass 3 — Remove empty table rows
# ---------------------------------------------------------------------------

def remove_empty_table_rows(doc: Document) -> None:
    """Remove rows (skipping the header row) where every cell is empty."""
    if not _CFG.get("remove_empty_table_rows", False):
        return

    ns = _NS
    for tbl in doc.element.body.iter(f"{{{ns}}}tbl"):
        rows = tbl.findall(f"{{{ns}}}tr")
        if len(rows) <= 1:
            continue
        for tr in rows[1:]:
            cells = tr.findall(f".//{{{ns}}}tc")
            if cells and all(
                not "".join(t.text or "" for t in tc.iter(f"{{{ns}}}t")).strip()
                for tc in cells
            ):
                tbl.remove(tr)


# ---------------------------------------------------------------------------
# Pass 4 — Set all text colour to black
# ---------------------------------------------------------------------------

def set_all_text_black(doc: Document) -> None:
    """Strip explicit colour runs so all text renders as black."""
    if not _CFG.get("set_all_text_black", False):
        return

    for rPr in doc.element.body.iter(f"{{{_NS}}}rPr"):
        color = rPr.find(f"{{{_NS}}}color")
        if color is not None:
            rPr.remove(color)


# ---------------------------------------------------------------------------
# Pass 5 — Sort tables alphabetically under specified headings
# ---------------------------------------------------------------------------

def sort_tables_alphabetically(doc: Document) -> None:
    """
    For each heading listed in CFG["sort_table_alphabetically_under_headings"],
    find the first table that follows it and sort its rows alphabetically by
    the first column (header row is kept in place).
    """
    targets: list[str] = _CFG.get("sort_table_alphabetically_under_headings", [])
    if not targets:
        return

    h_ids = _heading_style_ids(doc)
    body = list(doc.element.body)
    targets_lower = [t.strip().lower() for t in targets]

    for i, elem in enumerate(body):
        if elem.tag != f"{{{_NS}}}p" or _para_style_id(elem) not in h_ids:
            continue
        if _para_text(elem).strip().lower() not in targets_lower:
            continue

        # Find first <w:tbl> after this heading (before next heading)
        for nxt in body[i + 1:]:
            if nxt.tag == f"{{{_NS}}}p" and _para_style_id(nxt) in h_ids:
                break
            if nxt.tag == f"{{{_NS}}}tbl":
                _sort_table(nxt)
                break


def _sort_table(tbl_elem) -> None:
    rows = tbl_elem.findall(f"{{{_NS}}}tr")
    if len(rows) <= 1:
        return

    header, *data_rows = rows

    def first_col_key(tr):
        tc = tr.find(f"{{{_NS}}}tc")
        if tc is None:
            return ""
        return "".join(t.text or "" for t in tc.iter(f"{{{_NS}}}t")).strip().lower()

    data_rows.sort(key=first_col_key)

    for r in rows:
        tbl_elem.remove(r)
    tbl_elem.append(header)
    for r in data_rows:
        tbl_elem.append(r)


# ---------------------------------------------------------------------------
# Pass 6 — Remove specified sections (heading + content)
# ---------------------------------------------------------------------------

def remove_sections(doc: Document) -> None:
    """
    Remove each heading listed in CFG["remove_sections"] together with all
    body elements that follow it (up to the next heading).
    """
    targets: list[str] = _CFG.get("remove_sections", [])
    if not targets:
        return

    h_ids = _heading_style_ids(doc)
    targets_lower = [t.strip().lower() for t in targets]
    ns = _NS
    body = doc.element.body

    for target in targets_lower:
        children = list(body)
        for i, elem in enumerate(children):
            if elem.tag == f"{{{ns}}}p" and _para_style_id(elem) in h_ids:
                if _para_text(elem).strip().lower() == target:
                    to_remove = [elem]
                    for nxt in children[i + 1:]:
                        if nxt.tag == f"{{{ns}}}p" and _para_style_id(nxt) in h_ids:
                            break
                        to_remove.append(nxt)
                    for e in to_remove:
                        body.remove(e)
                    break


# ---------------------------------------------------------------------------
# Pass 7 — Fill placeholders throughout document
# ---------------------------------------------------------------------------

def fill_header_footer_placeholders(doc: Document, product_name: str | None) -> None:
    """
    Replace every placeholder listed in CFG["header_footer_placeholders"]
    with *product_name* throughout the entire document: body paragraphs,
    body tables (including cover page), and all headers/footers.

    Uses lxml directly so it works on both native python-docx elements and
    raw XML tables inserted at the zip level. Also handles placeholders that
    Word has split across multiple <w:r> runs.
    """
    if not product_name:
        return

    placeholders: list[str] = _CFG.get("header_footer_placeholders", [])
    if not placeholders:
        return

    def _replace_in_para_elem(p_elem) -> None:
        run_elems = p_elem.findall('.//' + qn('w:r'))
        if not run_elems:
            return
        for placeholder in placeholders:
            while True:
                texts = [''.join(t.text or '' for t in r.findall('.//' + qn('w:t')))
                         for r in run_elems]
                full = ''.join(texts)
                if placeholder not in full:
                    break
                ph_start = full.index(placeholder)
                ph_end   = ph_start + len(placeholder)
                offsets = [0]
                for t in texts[:-1]:
                    offsets.append(offsets[-1] + len(t))
                first_ri = next(i for i, off in enumerate(offsets)
                                if off + len(texts[i]) > ph_start)
                last_ri  = next(i for i in range(len(run_elems) - 1, -1, -1)
                                if offsets[i] < ph_end)
                pre  = texts[first_ri][:ph_start - offsets[first_ri]]
                post = texts[last_ri][ph_end - offsets[last_ri]:]
                new_text = pre + product_name + post
                # Write into first run's <w:t>
                wt_list = run_elems[first_ri].findall('.//' + qn('w:t'))
                if wt_list:
                    wt_list[0].text = new_text
                    for wt in wt_list[1:]:
                        wt.getparent().remove(wt)
                else:
                    from lxml import etree as _etree
                    wt = _etree.SubElement(run_elems[first_ri], qn('w:t'))
                    wt.text = new_text
                # Clear text in remaining involved runs
                for ri in range(first_ri + 1, last_ri + 1):
                    for wt in run_elems[ri].findall('.//' + qn('w:t')):
                        wt.text = ''

    def _replace_in_xml_elem(container_elem) -> None:
        """Iterate every <w:p> inside *container_elem* and replace placeholders."""
        for p_elem in container_elem.iter(qn('w:p')):
            _replace_in_para_elem(p_elem)

    # Document body
    _replace_in_xml_elem(doc.element.body)

    # Headers and footers
    for section in doc.sections:
        for part in (
            section.header, section.even_page_header, section.first_page_header,
            section.footer, section.even_page_footer, section.first_page_footer,
        ):
            try:
                _replace_in_xml_elem(part._element)
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def apply_all(doc: Document, product_name: str | None = None) -> None:
    """Run every configured post-processing pass on *doc* in-place."""
    remove_styled_paragraphs(doc)
    remove_template_instructions(doc)
    remove_empty_table_rows(doc)
    set_all_text_black(doc)
    sort_tables_alphabetically(doc)
    remove_sections(doc)
    fill_header_footer_placeholders(doc, product_name)
