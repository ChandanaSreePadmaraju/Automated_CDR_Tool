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
    Remove template instruction blocks delimited by '<' ... '>' paragraphs.

    A block opens when a paragraph's text is exactly '<' or starts with '< '
    and closes when a paragraph's text is exactly '>' or starts with '> '.
    Every paragraph in the block (both markers inclusive) is removed.
    This handles multi-paragraph instruction notes such as:
        < note:
        The Product identification RX.Y ...
        >
    """
    if not _CFG.get("remove_template_instructions", False):
        return

    body = doc.element.body
    to_remove = []
    in_block = False

    for e in list(body):
        if e.tag != f"{{{_NS}}}p":
            continue
        text = _para_text(e)
        if not in_block:
            if text == "<" or text.startswith("< "):
                in_block = True
                to_remove.append(e)
                # Single-line block: e.g. '< DELETE THIS >'; close immediately
                if text.endswith(">") and len(text) > 1:
                    in_block = False
        else:
            to_remove.append(e)
            if text == ">" or text.startswith("> "):
                in_block = False

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
    """Force all text to render as black.

    Strategy:
    1. Remove any explicit <w:color> element (was just deleting the override).
    2. Insert <w:color w:val="000000"/> so the run's colour is always explicit
       black regardless of what a named style (e.g. GuidanceChar) would apply.
    3. Remove character-style references (w:rStyle) whose style name contains
       'Guidance', so those styles can no longer override colour/font.
    """
    if not _CFG.get("set_all_text_black", False):
        return

    # Build set of rStyle ids that are Guidance character styles
    guidance_rStyle_ids: set[str] = set()
    for s in doc.styles:
        if s.name and "guidance" in s.name.lower() and s.type.name == "CHARACTER":
            guidance_rStyle_ids.add(s.style_id)

    from lxml import etree as _etree

    for rPr in doc.element.body.iter(f"{{{_NS}}}rPr"):
        # 1. Remove GuidanceChar rStyle
        rStyle = rPr.find(f"{{{_NS}}}rStyle")
        if rStyle is not None and rStyle.get(f"{{{_NS}}}val", "") in guidance_rStyle_ids:
            rPr.remove(rStyle)

        # 2. Remove existing color element (might be non-black)
        color = rPr.find(f"{{{_NS}}}color")
        if color is not None:
            rPr.remove(color)

        # 3. Insert explicit black color as first child so it takes effect
        black = _etree.Element(f"{{{_NS}}}color")
        black.set(f"{{{_NS}}}val", "000000")
        rPr.insert(0, black)


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

def remove_preamble_before_first_heading(doc: Document) -> None:
    """
    Remove visible-text paragraphs and duplicate page-break paragraphs that
    appear BEFORE the first heading (the template instruction page).
    Tables (cover page) are always preserved.
    Only the LAST page-break paragraph before the first heading is kept so
    that the first content section still starts on a fresh page without an
    extra blank page between the cover and the first heading.
    """
    if not _CFG.get("remove_preamble_before_first_heading", False):
        return

    h_ids = _heading_style_ids(doc)
    body = doc.element.body

    def _has_page_break(p_elem) -> bool:
        for br in p_elem.iter(f"{{{_NS}}}br"):
            if br.get(f"{{{_NS}}}type") == "page":
                return True
        pPr = p_elem.find(f"{{{_NS}}}pPr")
        if pPr is not None and pPr.find(f"{{{_NS}}}sectPr") is not None:
            return True
        return False

    # Collect all paragraph elements before the first heading
    pre_heading_paras: list = []
    for elem in list(body):
        if elem.tag == f"{{{_NS}}}p" and _para_style_id(elem) in h_ids:
            break
        if elem.tag == f"{{{_NS}}}p":
            pre_heading_paras.append(elem)

    # Find the FIRST and LAST page-break paragraphs.
    # The template has: COVER → [break1] → instruction page → [break2] → TOC → [break3] → content.
    # After removing the instruction page, we want to keep break1 (cover→TOC) and
    # break3 (TOC→content). Any intermediate page-break paragraphs are removed.
    pb_indices = [i for i, e in enumerate(pre_heading_paras) if _has_page_break(e)]
    keep_pb = set()
    if pb_indices:
        keep_pb.add(pb_indices[0])   # first: cover → TOC
        keep_pb.add(pb_indices[-1])  # last:  TOC  → content

    to_remove = []
    for i, elem in enumerate(pre_heading_paras):
        text = ''.join(t.text or '' for t in elem.iter(f"{{{_NS}}}t")).strip()
        is_pb = _has_page_break(elem)
        if i in keep_pb:
            continue  # keep first and last page breaks
        if text or is_pb:
            to_remove.append(elem)

    for e in to_remove:
        body.remove(e)


def strip_inline_angle_brackets(doc: Document) -> None:
    """
    Remove standalone '<' and '>' runs from every body paragraph.
    Template placeholders like  <Philips Medical Systems Nederland B.V.>
    have the angle brackets as separate <w:r> runs; we keep the text
    between them and simply delete the bracket runs.
    """
    if not _CFG.get("strip_inline_angle_brackets", False):
        return

    for p_elem in doc.element.body.iter(f"{{{_NS}}}p"):
        for r in list(p_elem.iter(f"{{{_NS}}}r")):
            text = "".join(t.text or "" for t in r.iter(f"{{{_NS}}}t")).strip()
            if text in ("<", ">"):
                parent = r.getparent()
                if parent is not None:
                    parent.remove(r)


def remove_paragraphs_with_text(doc: Document) -> None:
    """
    Remove any body-level paragraph whose text contains one of the
    strings listed in CFG["remove_paragraphs_with_text"].
    """
    targets: list[str] = _CFG.get("remove_paragraphs_with_text", [])
    if not targets:
        return

    body = doc.element.body
    to_remove = []
    for elem in list(body):
        if elem.tag != f"{{{_NS}}}p":
            continue
        text = _para_text(elem)
        if any(t.lower() in text.lower() for t in targets):
            to_remove.append(elem)
    for e in to_remove:
        body.remove(e)


def inject_definitions_fixed_rows(doc: Document) -> None:
    """
    Ensure specific rows (e.g. CDR, ISO, IGT-S) always exist in the
    Definitions & abbreviations table, then sorting is handled by the
    existing sort_tables_alphabetically pass.
    The rows are defined in CFG["definitions_fixed_rows"] as
    [["abbrev", "definition"], ...].
    CFG["definitions_heading"] names the heading to look under.
    """
    fixed: list[list[str]] = _CFG.get("definitions_fixed_rows", [])
    heading_text: str       = _CFG.get("definitions_heading", "Definitions & abbreviations")
    if not fixed:
        return

    h_ids = _heading_style_ids(doc)
    body  = list(doc.element.body)

    # Find the definitions table
    tbl_elem = None
    for i, elem in enumerate(body):
        if (elem.tag == f"{{{_NS}}}p"
                and _para_style_id(elem) in h_ids
                and _para_text(elem).strip().lower() == heading_text.lower()):
            for nxt in body[i + 1:]:
                if nxt.tag == f"{{{_NS}}}p" and _para_style_id(nxt) in h_ids:
                    break
                if nxt.tag == f"{{{_NS}}}tbl":
                    tbl_elem = nxt
                    break
            break

    if tbl_elem is None:
        return

    rows = list(tbl_elem.findall(f"{{{_NS}}}tr"))
    if not rows:
        return

    # Collect existing first-column values (lower-cased) to detect duplicates
    existing_keys: set[str] = set()
    for tr in rows[1:]:  # skip header
        tc = tr.find(f"{{{_NS}}}tc")
        if tc is not None:
            existing_keys.add(
                "".join(t.text or "" for t in tc.iter(f"{{{_NS}}}t")).strip().lower()
            )

    # Use the last data row as a style template for new rows
    template_row = rows[-1]

    from copy import deepcopy
    from lxml import etree

    for abbrev, definition in fixed:
        if abbrev.lower() in existing_keys:
            continue  # already present

        # Clone the template row and overwrite cell texts
        new_row = deepcopy(template_row)
        cells   = new_row.findall(f"{{{_NS}}}tc")
        texts   = [abbrev, definition]
        for ci, cell in enumerate(cells[:len(texts)]):
            # Clear existing runs and set plain text
            for r in cell.findall(f".//{{{_NS}}}r"):
                parent = r.getparent()
                if parent is not None:
                    parent.remove(r)
            # Re-add a simple run with the text
            for p_elem in cell.findall(f"{{{_NS}}}p"):
                r_elem = etree.SubElement(p_elem, f"{{{_NS}}}r")
                t_elem = etree.SubElement(r_elem, f"{{{_NS}}}t")
                t_elem.text = texts[ci]
                t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                break

        tbl_elem.append(new_row)


def add_page_break_before_headings(doc: Document) -> None:
    """
    Add w:pageBreakBefore to every paragraph whose style matches an entry in
    CFG["page_break_before_headings"] (list of heading style names, e.g. ["Heading 1"]).
    The first heading in the document is skipped — the page before it is already
    produced by the cover-page page-break paragraph.
    """
    targets: list[str] = _CFG.get("page_break_before_headings", [])
    if not targets:
        return

    targets_lower = {t.strip().lower() for t in targets}
    # Build style-id → style-name map
    id_to_name: dict[str, str] = {
        s.style_id: s.name for s in doc.styles if s.name
    }

    from lxml import etree as _etree

    body = doc.element.body
    first_skipped = False
    for elem in body:
        if elem.tag != f"{{{_NS}}}p":
            continue
        style_id = _para_style_id(elem)
        style_name = id_to_name.get(style_id, "")
        if style_name.lower() not in targets_lower:
            continue
        if not first_skipped:
            first_skipped = True
            continue  # skip the very first matching heading (cover already breaks before it)

        pPr = elem.find(f"{{{_NS}}}pPr")
        if pPr is None:
            pPr = _etree.SubElement(elem, f"{{{_NS}}}pPr")
            elem.insert(0, pPr)
        # Only add if not already present
        if pPr.find(f"{{{_NS}}}pageBreakBefore") is None:
            pb = _etree.SubElement(pPr, f"{{{_NS}}}pageBreakBefore")
            # insert after pStyle if present
            ps = pPr.find(f"{{{_NS}}}pStyle")
            if ps is not None:
                ps.addnext(pb)


def apply_all(doc: Document, product_name: str | None = None) -> None:
    """Run every configured post-processing pass on *doc* in-place."""
    fill_header_footer_placeholders(doc, product_name)   # must run BEFORE bracket stripping
    remove_preamble_before_first_heading(doc)
    remove_styled_paragraphs(doc)
    remove_template_instructions(doc)
    strip_inline_angle_brackets(doc)
    remove_paragraphs_with_text(doc)
    remove_empty_table_rows(doc)
    set_all_text_black(doc)
    inject_definitions_fixed_rows(doc)
    sort_tables_alphabetically(doc)
    remove_sections(doc)
    # Note: add_page_break_before_headings is available but not called by default
    # because the template has no page breaks between content sections.
