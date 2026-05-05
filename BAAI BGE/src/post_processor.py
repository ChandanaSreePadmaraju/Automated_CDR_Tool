"""
post_processor.py
-----------------
Post-processing passes applied to the filled output document before saving.
All behaviour is driven by the "post_processing" block in prompts.json —
nothing is hardcoded in this module.
"""

import json
import os
import re
from copy import deepcopy

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

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
# Pass 3 — Remove paragraphs with specific styles (e.g. "Guidance")
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
# Pass 4 — Remove template instruction bracket markers  (< / >)
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
# Pass 7 — Remove empty table rows
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
# Pass 10 — Set all text colour to black
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

    # Build set of rStyle ids whose style name contains any of the
    # styles_to_remove names — these are character-style variants
    # (e.g. "GuidanceChar") that must also be stripped so they cannot
    # override the forced black colour.  Driven by prompts.json so no
    # style names are hardcoded here.
    char_style_prefixes: list[str] = [
        s.lower() for s in _CFG.get("styles_to_remove", [])
    ]
    guidance_rStyle_ids: set[str] = set()
    for s in doc.styles:
        if (
            s.name
            and s.type.name == "CHARACTER"
            and any(prefix in s.name.lower() for prefix in char_style_prefixes)
        ):
            guidance_rStyle_ids.add(s.style_id)

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
        black = etree.Element(f"{{{_NS}}}color")
        black.set(f"{{{_NS}}}val", "000000")
        rPr.insert(0, black)

    # Also apply to headers and footers — they are separate XML parts not in body
    for section in doc.sections:
        for part in (
            section.header, section.even_page_header, section.first_page_header,
            section.footer, section.even_page_footer, section.first_page_footer,
        ):
            try:
                for rPr in part._element.iter(f"{{{_NS}}}rPr"):
                    rStyle = rPr.find(f"{{{_NS}}}rStyle")
                    if rStyle is not None and rStyle.get(f"{{{_NS}}}val", "") in guidance_rStyle_ids:
                        rPr.remove(rStyle)
                    color = rPr.find(f"{{{_NS}}}color")
                    if color is not None:
                        rPr.remove(color)
                    black = etree.Element(f"{{{_NS}}}color")
                    black.set(f"{{{_NS}}}val", "000000")
                    rPr.insert(0, black)
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Pass 13 — Sort tables alphabetically under specified headings
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
# Pass 8 — Reduce font size for NOTE paragraphs
# ---------------------------------------------------------------------------

def set_note_text_size(doc: Document) -> None:
    """Reduce the font size of NOTE paragraphs and all following paragraphs
    that belong to the same NOTE block (until the next NOTE prefix, a heading,
    or an empty paragraph that is not a list item).

    Applies to body paragraphs only (not table cells — notes rarely appear
    inside tables).
    Prefixes and target size are read from prompts.json:
      note_text_prefixes   : list[str]  – e.g. ["NOTE"]
      note_text_size_half_pt: int       – half-points (16 = 8 pt)
    """
    prefixes  = [p.lower() for p in _CFG.get("note_text_prefixes", [])]
    size_val  = str(_CFG.get("note_text_size_half_pt", 18))
    if not prefixes:
        return

    W = f"{{{_NS}}}"
    h_ids = _heading_style_ids(doc)

    def _apply_size(p_elem) -> None:
        for r in p_elem.findall(f".//{W}r"):
            rPr = r.find(f"{W}rPr")
            if rPr is None:
                rPr = etree.Element(f"{W}rPr")
                r.insert(0, rPr)
            for tag in (f"{W}sz", f"{W}szCs"):
                for old in rPr.findall(tag):
                    rPr.remove(old)
            sz = etree.SubElement(rPr, f"{W}sz")
            sz.set(f"{W}val", size_val)
            szCs = etree.SubElement(rPr, f"{W}szCs")
            szCs.set(f"{W}val", size_val)

    # Process both body paragraphs and paragraphs inside table cells.
    # For body paragraphs: track NOTE blocks across consecutive paragraphs.
    # For table-cell paragraphs: apply size to the NOTE paragraph and all
    # following paragraphs in the same cell until an empty paragraph or heading.

    # --- Body-level paragraphs (block tracking) ---
    body_paras = [
        e for e in doc.element.body
        if e.tag == f"{W}p"
    ]

    in_note = False
    for p_elem in body_paras:
        p_text = "".join(t.text or "" for t in p_elem.iter(f"{W}t")).strip()
        style_id = _para_style_id(p_elem)

        if style_id in h_ids:
            in_note = False
            continue

        if any(p_text.lower().startswith(prefix) for prefix in prefixes):
            in_note = True

        if in_note:
            _apply_size(p_elem)
            if not p_text and style_id not in ("ListParagraph",):
                in_note = False

    # --- Table cell paragraphs ---
    for tbl in doc.element.body.iter(f"{W}tbl"):
        for tc in tbl.iter(f"{W}tc"):
            cell_paras = list(tc.findall(f"{W}p"))
            in_note = False
            for p_elem in cell_paras:
                p_text = "".join(t.text or "" for t in p_elem.iter(f"{W}t")).strip()
                style_id = _para_style_id(p_elem)

                if style_id in h_ids:
                    in_note = False
                    continue

                if any(p_text.lower().startswith(prefix) for prefix in prefixes):
                    in_note = True

                if in_note:
                    _apply_size(p_elem)
                    if not p_text and style_id not in ("ListParagraph",):
                        in_note = False


# ---------------------------------------------------------------------------
# Pass 9 — Strip superscript formatting from list-marker runs
# ---------------------------------------------------------------------------

_MARKER_RE = re.compile(r'^\d+[.\s]*$')


def strip_superscript_list_markers(doc: Document) -> None:
    """Remove <w:vertAlign val="superscript"/> only from LEADING numeric runs.

    A run is considered a leading list marker if:
      - Its text matches the bare-number pattern (e.g. "1", "1.", "2."), AND
      - All runs that appear before it in the same paragraph have no visible
        text (i.e. the superscript is the very first visible content).

    This preserves legitimate trailing footnote-reference superscripts (e.g.
    a small "1" appended after "Instructions for Use Azurion R3.0") while
    removing numbered-list counters that Word copied as superscripts.

    Controlled by prompts.json: "strip_superscript_list_markers": true
    """
    if not _CFG.get("strip_superscript_list_markers", False):
        return

    W = f"{{{_NS}}}"
    for p_elem in doc.element.body.iter(f"{W}p"):
        runs = p_elem.findall(f".//{W}r")
        for idx, r in enumerate(runs):
            rPr = r.find(f"{W}rPr")
            if rPr is None:
                continue
            vert = rPr.find(f"{W}vertAlign")
            if vert is None or vert.get(f"{W}val") != "superscript":
                continue
            t = r.find(f"{W}t")
            text = (t.text or "").strip() if t is not None else ""
            if not _MARKER_RE.match(text):
                continue
            # Only strip if no visible text exists in any earlier run
            preceding_text = "".join(
                (rt.text or "")
                for prev_r in runs[:idx]
                for rt in prev_r.findall(f"{W}t")
            ).strip()
            if not preceding_text:
                rPr.remove(vert)


# ---------------------------------------------------------------------------
# Pass 14 — Remove specified sections (heading + content)
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
                        # Never delete layout paragraphs (inline sectPr) —
                        # they carry the template's header/footer rId links.
                        if nxt.tag == f"{{{ns}}}p":
                            pPr = nxt.find(f"{{{ns}}}pPr")
                            if pPr is not None and pPr.find(f"{{{ns}}}sectPr") is not None:
                                continue
                        to_remove.append(nxt)
                    for e in to_remove:
                        body.remove(e)
                    break


def remove_headings_only(doc: Document) -> None:
    """
    Remove each heading listed in CFG["remove_headings_only"] together with
    any immediately following non-table paragraphs, but STOP (and keep) the
    first table or next heading encountered.  Use this instead of
    remove_sections when the heading's content (a table) must be preserved.
    """
    targets: list[str] = _CFG.get("remove_headings_only", [])
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
                        # Stop at any table — keep it
                        if nxt.tag == f"{{{ns}}}tbl":
                            break
                        # Stop at next heading — keep it
                        if nxt.tag == f"{{{ns}}}p" and _para_style_id(nxt) in h_ids:
                            break
                        # Skip layout paragraphs
                        if nxt.tag == f"{{{ns}}}p":
                            pPr = nxt.find(f"{{{ns}}}pPr")
                            if pPr is not None and pPr.find(f"{{{ns}}}sectPr") is not None:
                                continue
                        to_remove.append(nxt)
                    for e in to_remove:
                        body.remove(e)
                    break


# Pass 11a — Remove blank paragraphs immediately BEFORE any heading
def remove_blank_paragraph_before_headings(doc: Document) -> None:
    """Remove ALL consecutive empty paragraphs that appear immediately before a Heading.

    Keeps layout paragraphs (page-breaks / inline sectPr) and paragraphs
    with images intact.
    """
    if not _CFG.get("remove_blank_para_before_headings", False):
        return

    def _is_blank(elem) -> bool:
        if elem.tag != f"{{{_NS}}}p":
            return False
        pPr = elem.find(f"{{{_NS}}}pPr")
        if pPr is not None and pPr.find(f"{{{_NS}}}sectPr") is not None:
            return False  # layout paragraph — keep
        for br in elem.iter(f"{{{_NS}}}br"):
            if br.get(f"{{{_NS}}}type") == "page":
                return False  # page-break — keep
        text = ''.join(t.text or '' for t in elem.iter(f"{{{_NS}}}t")).strip()
        has_drawing = any(n.tag == qn('w:drawing') for n in elem.iter())
        return not text and not has_drawing

    h_ids = _heading_style_ids(doc)
    body_elem = doc.element.body

    changed = True
    while changed:
        changed = False
        children = list(body_elem)
        for i, elem in enumerate(children):
            if elem.tag != f"{{{_NS}}}p":
                continue
            if _para_style_id(elem) not in h_ids:
                continue
            # Remove all consecutive blank paragraphs directly before this heading
            j = i - 1
            while j >= 0 and _is_blank(children[j]):
                try:
                    body_elem.remove(children[j])
                    changed = True
                except Exception:
                    pass
                j -= 1


# Pass 11b
def remove_blank_paragraph_after_headings(doc: Document) -> None:
    """Remove ALL consecutive empty paragraphs that immediately follow a Heading.

    Keeps layout paragraphs (page-breaks / inline sectPr) and paragraphs
    with images intact.
    """
    if not _CFG.get("remove_blank_para_after_headings", False):
        return

    def _is_blank(elem) -> bool:
        if elem.tag != f"{{{_NS}}}p":
            return False
        pPr = elem.find(f"{{{_NS}}}pPr")
        if pPr is not None and pPr.find(f"{{{_NS}}}sectPr") is not None:
            return False  # layout paragraph — keep
        for br in elem.iter(f"{{{_NS}}}br"):
            if br.get(f"{{{_NS}}}type") == "page":
                return False  # page-break — keep
        text = ''.join(t.text or '' for t in elem.iter(f"{{{_NS}}}t")).strip()
        has_drawing = any(n.tag == qn('w:drawing') for n in elem.iter())
        return not text and not has_drawing

    h_ids = _heading_style_ids(doc)
    body_elem = doc.element.body

    # Iterate live over children; re-read list after each removal so indices
    # stay correct even though we're removing elements.
    changed = True
    while changed:
        changed = False
        children = list(body_elem)
        for i, elem in enumerate(children):
            if elem.tag != f"{{{_NS}}}p":
                continue
            if _para_style_id(elem) not in h_ids:
                continue
            # Remove all consecutive blank paragraphs directly after this heading
            j = i + 1
            while j < len(children) and _is_blank(children[j]):
                try:
                    body_elem.remove(children[j])
                    changed = True
                except Exception:
                    pass
                j += 1


# ---------------------------------------------------------------------------
# Pass 1 — Fill placeholders throughout document
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
                    wt = etree.SubElement(run_elems[first_ri], qn('w:t'))
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
# Public API — passes 2, 5, 6, 11, 12, 15
# ---------------------------------------------------------------------------

# Pass 2
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


# Pass 5
def strip_inline_angle_brackets(doc: Document) -> None:
    """
    Remove standalone '<' and '>' runs from every paragraph (including inside
    table cells), and strip inline '<...>' comment substrings.

    Handles three cases:
    1. Standalone bracket runs: a run whose only text is '<' or '>' — removed.
    2. Single-run inline comment: '<...>' fully within one run's text — stripped.
    3. Multi-run inline comment: '<' in one run and '>' in a later run of the
       same paragraph — all runs from the opening '<' to the closing '>' are
       cleaned so the bracketed span is removed.
    """
    if not _CFG.get("strip_inline_angle_brackets", False):
        return

    _inline_re = re.compile(r'<[^<>]+>')

    def _strip_para(p_elem) -> None:
        runs = list(p_elem.findall(f"{{{_NS}}}r"))

        # Pass A — standalone bracket runs and single-run inline comments
        for r in runs:
            t_elems = r.findall(f"{{{_NS}}}t")
            run_text = "".join(t.text or "" for t in t_elems).strip()
            if run_text in ("<", ">"):
                parent = r.getparent()
                if parent is not None:
                    parent.remove(r)
            elif _inline_re.search(run_text):
                cleaned = _inline_re.sub("", run_text).strip()
                if t_elems:
                    t_elems[0].text = cleaned
                    for t in t_elems[1:]:
                        t.text = ""

        # Pass B — multi-run inline comments (re-read after Pass A removals)
        runs = list(p_elem.findall(f"{{{_NS}}}r"))
        open_run_idx = None
        for idx, r in enumerate(runs):
            run_text = "".join(t.text or "" for t in r.findall(f"{{{_NS}}}t"))
            if open_run_idx is None:
                lt_pos = run_text.find("<")
                if lt_pos != -1:
                    # Check the rest of this run — if '>' is also here, no span needed
                    after = run_text[lt_pos:]
                    if ">" not in after:
                        open_run_idx = idx
                        # Truncate this run at the '<'
                        for t in r.findall(f"{{{_NS}}}t"):
                            t.text = (t.text or "")[:lt_pos] if lt_pos == 0 else run_text[:lt_pos]
                            break
            else:
                gt_pos = run_text.find(">")
                if gt_pos != -1:
                    # Keep text after '>'
                    remainder = run_text[gt_pos + 1:]
                    for t in r.findall(f"{{{_NS}}}t"):
                        t.text = remainder
                        break
                    # Remove runs strictly between open_run_idx and idx
                    for mid in runs[open_run_idx + 1 : idx]:
                        parent = mid.getparent()
                        if parent is not None:
                            parent.remove(mid)
                    open_run_idx = None
                else:
                    # Entirely inside comment span — blank it out
                    for t in r.findall(f"{{{_NS}}}t"):
                        t.text = ""

    for p_elem in doc.element.body.iter(f"{{{_NS}}}p"):
        _strip_para(p_elem)


# Pass 6
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


# Pass 12
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
    heading_text: str       = _CFG.get("definitions_heading", "")
    if not fixed or not heading_text:
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

    for abbrev, definition in fixed:
        if abbrev.lower() in existing_keys:
            continue  # already present

        # Clone the template row and overwrite cell texts
        new_row = deepcopy(template_row)
        cells   = new_row.findall(f"{{{_NS}}}tc")
        texts   = [abbrev, definition]
        for ci, cell in enumerate(cells[:len(texts)]):
            # Capture run properties (font/size/bold) from the first run
            # before clearing, so the new run inherits the same formatting.
            existing_runs = cell.findall(f".//{{{_NS}}}r")
            saved_rPr = None
            for r in existing_runs:
                rPr = r.find(f"{{{_NS}}}rPr")
                if rPr is not None:
                    saved_rPr = deepcopy(rPr)
                    break

            # Clear existing runs
            for r in existing_runs:
                parent = r.getparent()
                if parent is not None:
                    parent.remove(r)

            # Re-add a run with the same rPr so font matches the table
            for p_elem in cell.findall(f"{{{_NS}}}p"):
                r_elem = etree.SubElement(p_elem, f"{{{_NS}}}r")
                if saved_rPr is not None:
                    r_elem.insert(0, saved_rPr)
                t_elem = etree.SubElement(r_elem, f"{{{_NS}}}t")
                t_elem.text = texts[ci]
                t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                break

        tbl_elem.append(new_row)

# ---------------------------------------------------------------------------
# Pass 16b — Prepend header row to compliance checklist table
# ---------------------------------------------------------------------------

def prepend_compliance_table_header(doc: Document, template_path: str | None = None) -> None:
    """
    Prepend a column-header row to the compliance table.

    All formatting is read at runtime from the template document itself —
    nothing is hardcoded.  Strategy:
      1. Open the template doc and find its largest table under the configured
         heading (same search as for the output doc).
      2. Find the first row in that table that has actual text content and
         matches the expected column count — this is a section-header style
         row whose pPr/rPr/tcPr carry the correct font, size, spacing etc.
      3. Deepcopy that row wholesale: replace only the text in each cell.
      4. Adjust each cell's tcW width to match the actual widths found in the
         output table (which may differ because content was filled in).
      5. Prepend the new row and mark it as tblHeader.
    """
    cfg: dict = _CFG.get("compliance_table_header", {})
    if not cfg:
        return

    heading_text: str = cfg.get("heading", "")
    columns: list = cfg.get("columns", [])
    if not heading_text or not columns:
        return

    W = f"{{{_NS}}}"
    XML_NS = "http://www.w3.org/XML/1998/namespace"

    # ------------------------------------------------------------------
    # 1. Locate the compliance table in the OUTPUT doc
    # ------------------------------------------------------------------
    h_ids = _heading_style_ids(doc)
    body_children = list(doc.element.body)

    out_tbl = None
    best = 0
    for i, elem in enumerate(body_children):
        if elem.tag == f"{W}p" and _para_style_id(elem) in h_ids:
            if _para_text(elem).strip().lower() == heading_text.lower():
                for nxt in body_children[i + 1:]:
                    if nxt.tag == f"{W}p" and _para_style_id(nxt) in h_ids:
                        break
                    if nxt.tag == f"{W}tbl":
                        n = len(nxt.findall(f"{W}tr"))
                        if n > best:
                            best = n
                            out_tbl = nxt
                break

    if out_tbl is None:
        return

    out_rows = out_tbl.findall(f"{W}tr")
    if not out_rows:
        return

    # Idempotency
    if columns[0].lower() in "".join(
        t.text or "" for t in out_rows[0].iter(f"{W}t")
    ).strip().lower():
        return

    # Collect cell widths from the first row with len(columns) cells
    out_widths: list[tuple[str, str]] = []
    for r in out_rows:
        cells = r.findall(f"{W}tc")
        if len(cells) == len(columns):
            for c in cells:
                tcW = c.find(f"{W}tcPr/{W}tcW")
                out_widths.append((
                    tcW.get(f"{W}w", "0") if tcW is not None else "0",
                    tcW.get(f"{W}type", "dxa") if tcW is not None else "dxa",
                ))
            break
    while len(out_widths) < len(columns):
        out_widths.append(("0", "dxa"))

    # ------------------------------------------------------------------
    # 2. Find a style-donor row from the TEMPLATE doc
    # ------------------------------------------------------------------
    donor_row = None
    if template_path and os.path.isfile(template_path):
        tmpl_doc = Document(template_path)
        tmpl_h_ids = _heading_style_ids(tmpl_doc)
        tmpl_children = list(tmpl_doc.element.body)

        tmpl_tbl = None
        best_t = 0
        for i, elem in enumerate(tmpl_children):
            if elem.tag == f"{W}p" and _para_style_id(elem) in tmpl_h_ids:
                if _para_text(elem).strip().lower() == heading_text.lower():
                    for nxt in tmpl_children[i + 1:]:
                        if nxt.tag == f"{W}p" and _para_style_id(nxt) in tmpl_h_ids:
                            break
                        if nxt.tag == f"{W}tbl":
                            n = len(nxt.findall(f"{W}tr"))
                            if n > best_t:
                                best_t = n
                                tmpl_tbl = nxt
                    break

        if tmpl_tbl is not None:
            # Always use the first row that has text — these are section-header
            # rows (e.g. "4 | Risk analysis") and carry the correct bold/font
            # styling for a table header.  Data rows (4-col content rows) must
            # NOT be used as the donor even if their cell count matches,
            # because they carry content/data formatting, not header formatting.
            for r in tmpl_tbl.findall(f"{W}tr"):
                if "".join(t.text or "" for t in r.iter(f"{W}t")).strip():
                    donor_row = r
                    break

    # ------------------------------------------------------------------
    # 3. Build the header row
    #    If we have a donor row: deepcopy it and patch widths + text.
    #    If not (template unavailable): copy the first row of the output
    #    table — still better than hardcoding.
    # ------------------------------------------------------------------
    if donor_row is not None:
        new_row = deepcopy(donor_row)
    else:
        ref = next((r for r in out_rows if len(r.findall(f"{W}tc")) == len(columns)), out_rows[0])
        new_row = deepcopy(ref)

    # Mark as repeating header
    trPr = new_row.find(f"{W}trPr")
    if trPr is None:
        trPr = etree.Element(f"{W}trPr")
        new_row.insert(0, trPr)
    if trPr.find(f"{W}cantSplit") is None:
        trPr.insert(0, etree.Element(f"{W}cantSplit"))
    if trPr.find(f"{W}tblHeader") is None:
        etree.SubElement(trPr, f"{W}tblHeader")

    # Ensure there are exactly len(columns) cells; if donor had fewer
    # (e.g. 3-col row with gridSpan), we need to expand it.
    cells = new_row.findall(f"{W}tc")

    # Remove gridSpan so cells become independent
    for c in cells:
        tcPr = c.find(f"{W}tcPr")
        if tcPr is not None:
            gs = tcPr.find(f"{W}gridSpan")
            if gs is not None:
                tcPr.remove(gs)

    # Add missing cells by cloning the last cell
    while len(new_row.findall(f"{W}tc")) < len(columns):
        new_row.append(deepcopy(new_row.findall(f"{W}tc")[-1]))

    # Remove extra cells
    cells = new_row.findall(f"{W}tc")
    for extra in cells[len(columns):]:
        new_row.remove(extra)

    # Patch cell widths to match output table and write column text
    cells = new_row.findall(f"{W}tc")
    for ci, col_text in enumerate(columns):
        cell = cells[ci]

        # Update width
        w_val, w_type = out_widths[ci]
        tcPr = cell.find(f"{W}tcPr")
        if tcPr is None:
            tcPr = etree.SubElement(cell, f"{W}tcPr")
            cell.insert(0, tcPr)
        tcW = tcPr.find(f"{W}tcW")
        if tcW is None:
            tcW = etree.SubElement(tcPr, f"{W}tcW")
        tcW.set(f"{W}w", w_val)
        tcW.set(f"{W}type", w_type)

        # Replace text: keep only the first paragraph, clear runs, add one
        all_paras = cell.findall(f"{W}p")
        for extra in all_paras[1:]:
            cell.remove(extra)
        p = all_paras[0] if all_paras else etree.SubElement(cell, f"{W}p")

        for r_elem in list(p.findall(f"{W}r")):
            p.remove(r_elem)
        for child in list(p):
            if child.tag not in (f"{W}pPr", f"{W}r"):
                p.remove(child)

        # Derive rPr from pPr/rPr (always has full font info)
        pPr = p.find(f"{W}pPr")
        base_rPr = pPr.find(f"{W}rPr") if pPr is not None else None

        r_new = etree.SubElement(p, f"{W}r")
        new_rPr = deepcopy(base_rPr) if base_rPr is not None else etree.Element(f"{W}rPr")
        if new_rPr.find(f"{W}b") is None:
            new_rPr.insert(0, etree.Element(f"{W}b"))
        if new_rPr.find(f"{W}bCs") is None:
            etree.SubElement(new_rPr, f"{W}bCs")
        r_new.insert(0, new_rPr)

        t = etree.SubElement(r_new, f"{W}t")
        t.text = col_text
        t.set(f"{{{XML_NS}}}space", "preserve")

    # Apply gray shading to all header cells so the column header row is
    # visually distinct from the section-header data rows (which have fill=auto).
    for cell in new_row.findall(f"{W}tc"):
        tcPr = cell.find(f"{W}tcPr")
        if tcPr is None:
            tcPr = etree.SubElement(cell, f"{W}tcPr")
            cell.insert(0, tcPr)
        shd = tcPr.find(f"{W}shd")
        if shd is None:
            shd = etree.SubElement(tcPr, f"{W}shd")
        shd.set(f"{W}val", "clear")
        shd.set(f"{W}color", "auto")
        shd.set(f"{W}fill", "D9D9D9")

    # Prepend before the current first row
    out_tbl.insert(list(out_tbl).index(out_rows[0]), new_row)

    W = f"{{{_NS}}}"
    h_ids = _heading_style_ids(doc)
    body_children = list(doc.element.body)

    # Find the compliance table — the LARGEST table under the heading
    tbl_elem = None
    best_rows = 0
    for i, elem in enumerate(body_children):
        if elem.tag == f"{W}p" and _para_style_id(elem) in h_ids:
            if _para_text(elem).strip().lower() == heading_text.lower():
                for nxt in body_children[i + 1:]:
                    if nxt.tag == f"{W}p" and _para_style_id(nxt) in h_ids:
                        break
                    if nxt.tag == f"{W}tbl":
                        n = len(nxt.findall(f"{W}tr"))
                        if n > best_rows:
                            best_rows = n
                            tbl_elem = nxt
                break

    if tbl_elem is None:
        return

    rows = tbl_elem.findall(f"{W}tr")
    if not rows:
        return

    # Idempotency — skip if header already present
    if columns[0].lower() in "".join(
        t.text or "" for t in rows[0].iter(f"{W}t")
    ).strip().lower():
        return

    # Read actual column widths from the first row that has len(columns) cells
    ref_row = next((r for r in rows if len(r.findall(f"{W}tc")) == len(columns)), None)
    cell_widths: list[tuple[str, str]] = []
    if ref_row is not None:
        for c in ref_row.findall(f"{W}tc"):
            tcW = c.find(f"{W}tcPr/{W}tcW")
            cell_widths.append((
                tcW.get(f"{W}w", "0") if tcW is not None else "0",
                tcW.get(f"{W}type", "dxa") if tcW is not None else "dxa",
            ))
    while len(cell_widths) < len(columns):
        cell_widths.append(("0", "dxa"))

    # ------------------------------------------------------------------
    # Build the row from scratch using the template's section-header style.
    # Every attribute below was read directly from the template file's
    # compliance table section-header rows ("4 | Risk analysis", etc.).
    # ------------------------------------------------------------------
    XML_NS = "http://www.w3.org/XML/1998/namespace"

    new_row = etree.Element(f"{W}tr")
    trPr = etree.SubElement(new_row, f"{W}trPr")
    etree.SubElement(trPr, f"{W}cantSplit")
    etree.SubElement(trPr, f"{W}tblHeader")

    for ci, col_text in enumerate(columns):
        w_val, w_type = cell_widths[ci]

        tc = etree.SubElement(new_row, f"{W}tc")

        # ---- tcPr -------------------------------------------------------
        tcPr = etree.SubElement(tc, f"{W}tcPr")
        tcW_e = etree.SubElement(tcPr, f"{W}tcW")
        tcW_e.set(f"{W}w", w_val)
        tcW_e.set(f"{W}type", w_type)
        # Gray background on all header cells to distinguish from data rows
        shd = etree.SubElement(tcPr, f"{W}shd")
        shd.set(f"{W}val", "clear")
        shd.set(f"{W}color", "auto")
        shd.set(f"{W}fill", "D9D9D9")

        # ---- paragraph --------------------------------------------------
        p = etree.SubElement(tc, f"{W}p")

        pPr = etree.SubElement(p, f"{W}pPr")
        ps = etree.SubElement(pPr, f"{W}pStyle")
        ps.set(f"{W}val", "Default")
        etree.SubElement(pPr, f"{W}keepNext")
        etree.SubElement(pPr, f"{W}keepLines")
        sp = etree.SubElement(pPr, f"{W}spacing")
        sp.set(f"{W}before", "66")
        sp.set(f"{W}after", "54")
        if ci == len(columns) - 1:         # Verdict → centre
            jc = etree.SubElement(pPr, f"{W}jc")
            jc.set(f"{W}val", "center")

        # pPr/rPr  (paragraph-mark formatting — matches section-header rows)
        pRpr = etree.SubElement(pPr, f"{W}rPr")
        rf_p = etree.SubElement(pRpr, f"{W}rFonts")
        rf_p.set(f"{W}asciiTheme", "minorHAnsi")
        rf_p.set(f"{W}hAnsiTheme", "minorHAnsi")
        rf_p.set(f"{W}cstheme",    "minorHAnsi")
        etree.SubElement(pRpr, f"{W}b")
        etree.SubElement(pRpr, f"{W}bCs")
        col_p = etree.SubElement(pRpr, f"{W}color")
        col_p.set(f"{W}val", "auto")
        sz_p = etree.SubElement(pRpr, f"{W}sz");    sz_p.set(f"{W}val", "22")
        szc_p = etree.SubElement(pRpr, f"{W}szCs"); szc_p.set(f"{W}val", "22")
        lg_p = etree.SubElement(pRpr, f"{W}lang");  lg_p.set(f"{W}val", "en-US")

        # ---- run --------------------------------------------------------
        r_e = etree.SubElement(p, f"{W}r")

        rPr = etree.SubElement(r_e, f"{W}rPr")
        rf_r = etree.SubElement(rPr, f"{W}rFonts")
        rf_r.set(f"{W}asciiTheme", "minorHAnsi")
        rf_r.set(f"{W}hAnsiTheme", "minorHAnsi")
        rf_r.set(f"{W}cstheme",    "minorHAnsi")
        etree.SubElement(rPr, f"{W}b")
        etree.SubElement(rPr, f"{W}bCs")
        col_r = etree.SubElement(rPr, f"{W}color")
        col_r.set(f"{W}val", "auto")
        sz_r = etree.SubElement(rPr, f"{W}sz");    sz_r.set(f"{W}val", "22")
        szc_r = etree.SubElement(rPr, f"{W}szCs"); szc_r.set(f"{W}val", "22")
        lg_r = etree.SubElement(rPr, f"{W}lang");  lg_r.set(f"{W}val", "en-US")

        t_e = etree.SubElement(r_e, f"{W}t")
        t_e.text = col_text
        t_e.set(f"{{{XML_NS}}}space", "preserve")

    # Prepend before the current first row
    tbl_elem.insert(list(tbl_elem).index(rows[0]), new_row)

# ---------------------------------------------------------------------------
# Pass 16 — Normalize rFonts attributes to remove locale-specific overrides
# ---------------------------------------------------------------------------

def normalize_rfonts(doc: Document) -> None:
    """
    Remove specific w:rFonts attributes (e.g. eastAsia, cs) that can cause
    inconsistent font rendering when rows were typed with different locale
    settings in the source document.  Driven by CFG["normalize_rfonts_remove_attributes"].
    """
    attrs_to_remove: list[str] = _CFG.get("normalize_rfonts_remove_attributes", [])
    if not attrs_to_remove:
        return

    W = f"{{{_NS}}}"
    # Build fully-qualified attribute names
    fq_attrs = [f"{W}{a}" for a in attrs_to_remove]

    for rFonts in doc.element.body.iter(f"{W}rFonts"):
        for attr in fq_attrs:
            if attr in rFonts.attrib:
                del rFonts.attrib[attr]


# Pass 15
def mark_toc_dirty(doc: Document) -> None:
    """
    Mark every TOC field as dirty so Word auto-updates the Table of Contents
    when the document is first opened.
    """
    W = f"{{{_NS}}}"
    for p_elem in doc.element.body.iter(f"{W}p"):
        instr_texts = [(t.text or "").strip() for t in p_elem.iter(f"{W}instrText")]
        if not any(txt.upper().startswith("TOC") for txt in instr_texts):
            continue
        for fldChar in p_elem.iter(f"{W}fldChar"):
            if fldChar.get(f"{W}fldCharType") == "begin":
                fldChar.set(f"{W}dirty", "1")

# ---------------------------------------------------------------------------
# Pass 17 — Fill ISO info table cells from source doc's matching table
#            Copies full paragraph XML (preserving fonts, colours, images)
# ---------------------------------------------------------------------------

_FILL_PLACEHOLDER_PREFIX = "rId_CDRCOPY_FILL_"
_fill_placeholder_counter = 0


def _remap_images_in_elem(elem, src_doc) -> list[dict]:
    """
    Deep-copy *elem* is assumed to have already been made.  Rewrite every
    a:blip r:embed attribute to a unique placeholder rId and collect
    { placeholder_rid, bytes, content_type } for later zip-level insertion.
    """
    global _fill_placeholder_counter
    R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
    A_BLIP  = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}blip"
    # Use the drawingml namespace for blip
    BLIP_TAG = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"

    images: list[dict] = []
    orig_to_ph: dict[str, str] = {}

    for elem_node in elem.iter():
        rid = elem_node.get(R_EMBED)
        if rid is None:
            continue
        if rid not in orig_to_ph:
            _fill_placeholder_counter += 1
            ph = f"{_FILL_PLACEHOLDER_PREFIX}{_fill_placeholder_counter:06d}"
            orig_to_ph[rid] = ph
            try:
                img_part = src_doc.part.related_parts[rid]
                images.append({
                    "placeholder_rid": ph,
                    "bytes":           img_part.blob,
                    "content_type":    img_part.content_type,
                })
            except (KeyError, AttributeError):
                pass
        elem_node.set(R_EMBED, orig_to_ph[rid])

    return images


def fill_table_from_source(doc: Document, data_doc_path: str | None) -> list[dict]:
    """
    Fill cells in the preserved ISO info table by copying full paragraph XML
    (runs, formatting, images) from the matching rows in the source document.

    Returns a list of image_parts dicts (placeholder_rid / bytes / content_type)
    that must be passed to _copy_table_images for final zip-level wiring.

    Config key: "fill_table_from_source"
      target_heading : heading text under which the target table lives in output
      source_heading : heading text under which the source table lives in input
      row_label_map  : {target_label: source_label}
                       For 2-column target rows the value cell (col 1) is
                       replaced with the XML paragraphs from the source value cell.
                       For 1-column target rows the entire cell content is replaced
                       with the XML from the source cell (full cell clone).
    """
    from copy import deepcopy

    cfg: dict = _CFG.get("fill_table_from_source", {})
    if not cfg or not data_doc_path:
        return []

    target_heading: str = cfg.get("target_heading", "")
    source_heading: str = cfg.get("source_heading", "")
    row_label_map: dict = cfg.get("row_label_map", {})
    if not target_heading or not source_heading or not row_label_map:
        return []

    W = f"{{{_NS}}}"
    h_ids = _heading_style_ids(doc)
    all_image_parts: list[dict] = []

    def _cell_text(tc_elem) -> str:
        return "".join(t.text or "" for t in tc_elem.iter(f"{W}t")).strip()

    def _find_table_after_heading_output(body_children, heading_text: str):
        for i, elem in enumerate(body_children):
            if elem.tag == f"{W}p" and _para_style_id(elem) in h_ids:
                if _para_text(elem).strip().lower() == heading_text.lower():
                    for nxt in body_children[i + 1:]:
                        if nxt.tag == f"{W}p" and _para_style_id(nxt) in h_ids:
                            break
                        if nxt.tag == f"{W}tbl":
                            return nxt
        return None

    def _find_source_table_by_content(body_children, key_labels: list):
        key_labels_lower = {k.lower() for k in key_labels}
        best_tbl, best_score = None, 0
        for elem in body_children:
            if elem.tag != f"{W}tbl":
                continue
            score = 0
            for tr in elem.findall(f"{W}tr"):
                cells = tr.findall(f"{W}tc")
                if not cells:
                    continue
                cell0 = _cell_text(cells[0]).lower()
                for kl in key_labels_lower:
                    if cell0.startswith(kl):
                        score += 1
                        break
            if score > best_score:
                best_score = score
                best_tbl = elem
        return best_tbl if best_score > 0 else None

    # --- Load source doc ---
    from docx import Document as _Doc
    src_doc = _Doc(data_doc_path)
    src_children = list(src_doc.element.body)
    src_key_labels = list(row_label_map.values())
    src_tbl = _find_source_table_by_content(src_children, src_key_labels)
    if src_tbl is None:
        return []

    # Build map: lowercase source label → source tc element
    src_label_to_tc: dict = {}
    for tr in src_tbl.findall(f"{W}tr"):
        cells = tr.findall(f"{W}tc")
        if not cells:
            continue
        cell_text = _cell_text(cells[0])
        if len(cells) == 2:
            src_label_to_tc[cell_text.lower()] = cells[1]
        else:
            colon_pos = cell_text.find(":")
            label_key = (cell_text[:colon_pos + 1] if colon_pos != -1 else cell_text).lower()
            src_label_to_tc[label_key] = cells[0]

    row_label_map_lower = {k.lower(): v.lower() for k, v in row_label_map.items()}

    # --- Find target table in output ---
    out_children = list(doc.element.body)
    target_tbl = _find_table_after_heading_output(out_children, target_heading)
    if target_tbl is None:
        return []

    # --- For each mapped target row, replace cell content with source XML ---
    for tr in target_tbl.findall(f"{W}tr"):
        cells = tr.findall(f"{W}tc")
        if not cells:
            continue
        label_text = _cell_text(cells[0])
        label_lower = label_text.rstrip().lower()
        # For 1-column rows the template cell may have placeholder text after
        # the label (e.g. "Copy of marking plate:The artwork below...").
        # Match by the prefix up to (and including) the first ':' so we can
        # use clean short keys in row_label_map.
        if len(cells) == 1:
            colon_pos = label_lower.find(":")
            label_key = label_lower[:colon_pos + 1] if colon_pos != -1 else label_lower
        else:
            label_key = label_lower

        src_key = row_label_map_lower.get(label_key)
        if src_key is None:
            continue

        src_tc = src_label_to_tc.get(src_key)
        if src_tc is None:
            continue

        if len(cells) == 2:
            # 2-column row: replace the value cell (col 1) content
            # Copy all <w:p> elements from source value cell into target value cell
            val_tc = cells[1]
            # Remove existing paragraphs from target value cell (keep tcPr)
            for p in list(val_tc.findall(f"{W}p")):
                val_tc.remove(p)
            # Deep-copy all content children from source value cell
            # (paragraphs AND any nested tables that carry images)
            src_children = [c for c in src_tc if c.tag != f"{W}tcPr"]
            for src_child in src_children:
                new_child = deepcopy(src_child)
                imgs = _remap_images_in_elem(new_child, src_doc)
                all_image_parts.extend(imgs)
                val_tc.append(new_child)
            # Ensure there is at least one paragraph (Word requires it)
            if not val_tc.findall(f"{W}p"):
                val_tc.append(etree.SubElement(val_tc, f"{W}p"))
        else:
            # 1-column row: replace all content in the target cell with
            # a deep copy of the source cell's contents (paragraphs + nested tables)
            val_tc = cells[0]
            # Save tcPr if present
            tc_pr = val_tc.find(f"{W}tcPr")
            # Remove all children
            for child in list(val_tc):
                val_tc.remove(child)
            # Restore tcPr
            if tc_pr is not None:
                val_tc.append(deepcopy(tc_pr))
            # Copy all content children (not tcPr)
            src_children = [c for c in src_tc if c.tag != f"{W}tcPr"]
            for src_child in src_children:
                new_child = deepcopy(src_child)
                imgs = _remap_images_in_elem(new_child, src_doc)
                all_image_parts.extend(imgs)
                val_tc.append(new_child)
            # Ensure at least one paragraph
            if not val_tc.findall(f"{W}p"):
                val_tc.append(etree.SubElement(val_tc, f"{W}p"))

    return all_image_parts


def apply_all(doc: Document, product_name: str | None = None, data_doc_path: str | None = None, template_path: str | None = None) -> list[dict]:
    """Run every configured post-processing pass on *doc* in-place.
    Returns a list of image_parts dicts (may be empty) collected during
    fill_table_from_source that must be passed to _copy_table_images.
    """
    fill_header_footer_placeholders(doc, product_name)   # must run BEFORE bracket stripping
    remove_preamble_before_first_heading(doc)
    remove_styled_paragraphs(doc)
    remove_template_instructions(doc)
    strip_inline_angle_brackets(doc)
    remove_paragraphs_with_text(doc)
    remove_empty_table_rows(doc)
    set_note_text_size(doc)
    strip_superscript_list_markers(doc)
    remove_sections(doc)
    remove_headings_only(doc)
    prepend_compliance_table_header(doc, template_path)   # before set_all_text_black so header gets same colour pass
    set_all_text_black(doc)
    remove_blank_paragraph_before_headings(doc)
    remove_blank_paragraph_after_headings(doc)
    inject_definitions_fixed_rows(doc)
    sort_tables_alphabetically(doc)
    normalize_rfonts(doc)
    extra_image_parts = fill_table_from_source(doc, data_doc_path)
    mark_toc_dirty(doc)
    return extra_image_parts
