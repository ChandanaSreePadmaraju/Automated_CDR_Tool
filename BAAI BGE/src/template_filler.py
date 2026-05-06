"""
template_filler.py
------------------
Insert content extracted from the data doc into the template doc under each
matched heading, then save the result to *output_path*.

Insertions are processed in **reverse paragraph-index order** so that adding
content after heading N never shifts the indices of headings above it.
"""

import copy
import json
import os
import re
import zipfile
from io import BytesIO

from docx import Document
from lxml import etree

from src.content_extractor import extract_section, extract_pre_heading_content, _NUMID_OFFSET
from src import post_processor, detect_compliance_heading


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _strip_inline_sectpr(elem) -> None:
    """
    Remove every inline <w:sectPr> found inside a <w:pPr> in *elem*.

    Paragraphs extracted from the data document carry their own
    <w:pPr><w:sectPr> elements whose rId values reference the data doc's
    header/footer files.  Those rIds are meaningless (or actively wrong) in
    the output document context, so they must be stripped before insertion.
    """
    _ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    for pPr in elem.findall(f".//{{{_ns}}}pPr"):
        for s in list(pPr.findall(f"{{{_ns}}}sectPr")):
            pPr.remove(s)


def _build_remapped_table(tmpl_tbl, data_tbl, col_indices: list, ns: str):
    """
    Return a new <w:tbl> element that:
      - Keeps the template table's properties (tblPr) and header row exactly
      - Populates data rows from *data_tbl* rows[1:], taking only the
        columns listed in *col_indices*
      - Applies the template's per-column cell properties (widths, borders)
        so the output table looks like the template, not the data doc.
    """
    W = f"{{{ns}}}"

    data_rows = data_tbl.findall(f"{W}tr")

    # Deep-copy the whole template table (preserves tblPr, tblGrid, header)
    result = copy.deepcopy(tmpl_tbl)
    result_rows = result.findall(f"{W}tr")

    # Remove everything after the first (header) row
    for row in result_rows[1:]:
        result.remove(row)

    # Capture template cell properties (widths/borders) from the header row
    tmpl_tcPr_list: list = []
    if result_rows:
        for tc in result_rows[0].findall(f"{W}tc"):
            tcPr = tc.find(f"{W}tcPr")
            tmpl_tcPr_list.append(copy.deepcopy(tcPr) if tcPr is not None else None)

    # Build new data rows from data_tbl (skip its header row)
    for data_row in data_rows[1:]:
        data_cells = data_row.findall(f"{W}tc")
        if not data_cells:
            continue

        new_row = copy.deepcopy(data_row)
        # Remove all cells from the copied row
        for tc in new_row.findall(f"{W}tc"):
            new_row.remove(tc)

        for i, col_idx in enumerate(col_indices):
            if col_idx < len(data_cells):
                new_tc = copy.deepcopy(data_cells[col_idx])
                # Replace cell properties with template's to keep correct widths
                if i < len(tmpl_tcPr_list) and tmpl_tcPr_list[i] is not None:
                    existing = new_tc.find(f"{W}tcPr")
                    if existing is not None:
                        new_tc.remove(existing)
                    new_tc.insert(0, copy.deepcopy(tmpl_tcPr_list[i]))
            else:
                # Column doesn't exist in data doc — add an empty cell
                new_tc = etree.Element(f"{W}tc")
                if i < len(tmpl_tcPr_list) and tmpl_tcPr_list[i] is not None:
                    new_tc.append(copy.deepcopy(tmpl_tcPr_list[i]))
                new_tc.append(etree.Element(f"{W}p"))
            new_row.append(new_tc)

        result.append(new_row)

    return result


def _insert_xml_after(ref_elem, xml_bytes: bytes):
    """
    Deserialise *xml_bytes* into a fresh lxml element, strip any inline sectPr
    elements, then insert it immediately after *ref_elem*.
    Comment markup (commentRangeStart/End/Reference) is preserved so Word
    comments from the data doc appear in the output.
    Works for any body child (<w:p>, <w:tbl>, …).
    Returns the newly inserted element.
    """
    new_elem = copy.deepcopy(etree.fromstring(xml_bytes))
    _strip_inline_sectpr(new_elem)
    ref_elem.addnext(new_elem)
    return new_elem


def _copy_table_images(output_bytes: bytes, all_image_parts: list[dict]) -> bytes:
    """
    Zip-level pass: copy images extracted from the data doc into the output
    docx and rewrite the placeholder r:embed references to final rIds.

    The XML bytes stored in each item already have the data-doc rIds replaced
    with unique placeholder values (rId_CDRCOPY_NNNNNN) by content_extractor,
    so this replacement never touches template-original rId references.
    """
    if not all_image_parts:
        return output_bytes

    # Deduplicate by placeholder_rid (one entry per unique placeholder)
    seen: dict[str, dict] = {}
    for item in all_image_parts:
        seen.setdefault(item["placeholder_rid"], item)
    unique = list(seen.values())

    ext_map = {
        "image/png": ".png",  "image/jpeg": ".jpg",  "image/jpg": ".jpg",
        "image/gif": ".gif",  "image/bmp": ".bmp",   "image/tiff": ".tif",
        "image/emf": ".emf",  "image/x-emf": ".emf", "image/wmf": ".wmf",
    }
    img_rel = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    )

    with zipfile.ZipFile(BytesIO(output_bytes), "r") as zin:
        existing = {n: zin.read(n) for n in zin.namelist()}

    rels_text = existing.get("word/_rels/document.xml.rels", b"").decode("utf-8")
    doc_text  = existing.get("word/document.xml",            b"").decode("utf-8")

    # Determine the highest existing rId number to avoid collisions
    existing_nums = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_text)]
    next_id = max(existing_nums, default=0) + 1

    rid_map: dict[str, str] = {}
    new_media: dict[str, bytes] = {}
    new_rels: list[str] = []

    for item in unique:
        placeholder = item["placeholder_rid"]
        ct  = item.get("content_type", "image/png")
        ext = ext_map.get(ct.lower(), ".bin")
        new_rid    = f"rId{next_id}"
        next_id   += 1
        media_name = f"img_copied_{new_rid}{ext}"
        new_media[f"word/media/{media_name}"] = item["bytes"]
        new_rels.append(
            f'<Relationship Id="{new_rid}" Type="{img_rel}" '
            f'Target="media/{media_name}"/>'
        )
        rid_map[placeholder] = new_rid

    if not rid_map:
        return output_bytes

    # Replace ONLY the safe placeholder rIds — never touches template rIds
    for placeholder, new_rid in rid_map.items():
        doc_text = doc_text.replace(f'r:embed="{placeholder}"', f'r:embed="{new_rid}"')

    # Inject new relationships before </Relationships>
    insert_at = rels_text.rfind("</")
    rels_text = (
        rels_text[:insert_at]
        + "\n  ".join(new_rels) + "\n"
        + rels_text[insert_at:]
    )

    out_buf = BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        skip = {"word/_rels/document.xml.rels", "word/document.xml"}
        for name, data in existing.items():
            if name not in skip:
                zout.writestr(name, data)
        zout.writestr("word/_rels/document.xml.rels", rels_text.encode("utf-8"))
        zout.writestr("word/document.xml",            doc_text.encode("utf-8"))
        for name, data in new_media.items():
            zout.writestr(name, data)

    return out_buf.getvalue()


_COMMENTS_REL  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
_COMMENTS_CT   = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
_COMMENTSEXT_REL = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
_COMMENTSEXT_CT  = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
_COMMENTSIDS_REL = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
_COMMENTSIDS_CT  = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"
_COMMENTSEXTENSIBLE_REL = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"
_COMMENTSEXTENSIBLE_CT  = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"


def _merge_comments(output_bytes: bytes, data_doc_path: str) -> bytes:
    """
    Zip-level pass: copy word/comments.xml (and related extended comment files)
    from the data doc into the output docx.

    If the output already has a comments part, the data-doc comments are merged
    in by appending their <w:comment> elements so both sets are preserved.
    Related parts (commentsExtended, commentsIds, commentsExtensible) are
    copied wholesale if not already present.
    """
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    # Files to carry over if present in the data doc (rel type, content type, zip name)
    _EXTRA_PARTS = [
        ("word/commentsExtended.xml",    _COMMENTSEXT_REL,         _COMMENTSEXT_CT),
        ("word/commentsIds.xml",         _COMMENTSIDS_REL,         _COMMENTSIDS_CT),
        ("word/commentsExtensible.xml",  _COMMENTSEXTENSIBLE_REL,  _COMMENTSEXTENSIBLE_CT),
    ]

    try:
        with zipfile.ZipFile(data_doc_path, "r") as zdata:
            data_names = zdata.namelist()
            if "word/comments.xml" not in data_names:
                return output_bytes
            src_comments_bytes = zdata.read("word/comments.xml")
            extra_bytes = {
                name: zdata.read(name)
                for name, _, _ in _EXTRA_PARTS
                if name in data_names
            }
    except Exception:
        return output_bytes

    with zipfile.ZipFile(BytesIO(output_bytes), "r") as zin:
        existing = {n: zin.read(n) for n in zin.namelist()}

    # ── Merge main comments.xml ────────────────────────────────────────────
    if "word/comments.xml" in existing:
        # Merge: append data-doc <w:comment> elements into the output's root
        try:
            out_root  = etree.fromstring(existing["word/comments.xml"])
            src_root  = etree.fromstring(src_comments_bytes)
            for child in src_root:
                if child.tag == f"{{{W}}}comment":
                    out_root.append(copy.deepcopy(child))
            existing["word/comments.xml"] = etree.tostring(
                out_root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
        except Exception:
            pass  # leave existing unchanged if parsing fails
    else:
        # Output has no comments part — add it and wire up relationship + content type
        existing["word/comments.xml"] = src_comments_bytes
        rels_text = existing.get("word/_rels/document.xml.rels", b"").decode("utf-8")
        if _COMMENTS_REL not in rels_text:
            existing_nums = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_text)]
            new_rid = f"rId{max(existing_nums, default=0) + 1}"
            insert_at = rels_text.rfind("</")
            rels_text = (
                rels_text[:insert_at]
                + f'  <Relationship Id="{new_rid}" Type="{_COMMENTS_REL}" Target="comments.xml"/>\n'
                + rels_text[insert_at:]
            )
            existing["word/_rels/document.xml.rels"] = rels_text.encode("utf-8")
        ct_text = existing.get("[Content_Types].xml", b"").decode("utf-8")
        if "comments.xml" not in ct_text:
            insert_at = ct_text.rfind("</")
            ct_text = (
                ct_text[:insert_at]
                + f'  <Override PartName="/word/comments.xml" ContentType="{_COMMENTS_CT}"/>\n'
                + ct_text[insert_at:]
            )
            existing["[Content_Types].xml"] = ct_text.encode("utf-8")

    # ── Drop extended comment parts entirely ──────────────────────────────
    # python-docx strips w14:paraId from every paragraph on save, so any
    # commentsIds / commentsExtended / commentsExtensible entries (which map
    # comment IDs to those paraIds) become dangling references.  Word's modern
    # comment renderer then fails to resolve the author identity and shows
    # blank names.  Removing these files forces Word to use its classic
    # comment display, which reads author names directly from comments.xml and
    # always shows them correctly.
    rels_text    = existing.get("word/_rels/document.xml.rels", b"").decode("utf-8")
    ct_text      = existing.get("[Content_Types].xml",           b"").decode("utf-8")
    changed_rels = False
    changed_ct   = False

    for zip_name, rel_type, ct_type in _EXTRA_PARTS:
        base = zip_name.split("/")[-1]
        if zip_name in existing:
            del existing[zip_name]
        # Remove relationship entry
        if rel_type in rels_text:
            rels_text = re.sub(
                r'\s*<Relationship[^>]*Type="' + re.escape(rel_type) + r'"[^>]*/>\s*',
                "\n",
                rels_text,
            )
            changed_rels = True
        # Remove content-type entry
        if base in ct_text:
            ct_text = re.sub(
                r'\s*<Override[^>]*PartName="/word/' + re.escape(base) + r'"[^>]*/>\s*',
                "\n",
                ct_text,
            )
            changed_ct = True

    if changed_rels:
        existing["word/_rels/document.xml.rels"] = rels_text.encode("utf-8")
    if changed_ct:
        existing["[Content_Types].xml"] = ct_text.encode("utf-8")

    # ── Remove people.xml ─────────────────────────────────────────────────
    # people.xml carries Azure AD (providerId="AD") identity entries.  When
    # present, Word tries to resolve commenters' identities against the
    # originating AD tenant.  If the reader is in a different tenant the
    # resolution fails and Word displays blank author names even though the
    # correct name is already stored in the w:author attribute of comments.xml.
    # Removing people.xml forces Word to fall back to the classic comment
    # display which reads w:author directly — always showing the right name.
    _PEOPLE_REL = "http://schemas.microsoft.com/office/2011/relationships/people"
    if "word/people.xml" in existing:
        del existing["word/people.xml"]
    rels_text = existing.get("word/_rels/document.xml.rels", b"").decode("utf-8")
    ct_text   = existing.get("[Content_Types].xml",           b"").decode("utf-8")
    if _PEOPLE_REL in rels_text:
        rels_text = re.sub(
            r'\s*<Relationship[^>]*Type="' + re.escape(_PEOPLE_REL) + r'"[^>]*/>' ,
            "",
            rels_text,
        )
        existing["word/_rels/document.xml.rels"] = rels_text.encode("utf-8")
    if "people.xml" in ct_text:
        ct_text = re.sub(
            r'\s*<Override[^>]*PartName="/word/people\.xml"[^>]*/>' ,
            "",
            ct_text,
        )
        existing["[Content_Types].xml"] = ct_text.encode("utf-8")

    out_buf = BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in existing.items():
            zout.writestr(name, data)

    return out_buf.getvalue()


def _merge_numbering(output_bytes: bytes, data_doc_path: str) -> bytes:
    """
    Zip-level pass: copy the abstractNum/num definitions that correspond to
    the sentinel numId values (>= _NUMID_OFFSET) inserted during extraction
    from the data doc's numbering.xml into the output doc's numbering.xml,
    then rewrite the sentinel values to the new real IDs.

    If the output doc has no sentinel numIds this is a no-op.
    """
    with zipfile.ZipFile(BytesIO(output_bytes), "r") as zin:
        existing = {n: zin.read(n) for n in zin.namelist()}

    doc_text = existing.get("word/document.xml", b"").decode("utf-8")

    # Collect all sentinel numId values present in the output document
    sentinel_ids = {
        int(v) for v in re.findall(r'<w:numId w:val="(\d+)"', doc_text)
        if int(v) >= _NUMID_OFFSET
    }
    if not sentinel_ids:
        return output_bytes

    # Map sentinel → original data-doc numId
    orig_numids = {s: s - _NUMID_OFFSET for s in sentinel_ids}
    needed_orig = set(orig_numids.values())

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    WNS = f"{{{W}}}"

    # Read data doc's numbering.xml
    try:
        with zipfile.ZipFile(data_doc_path, "r") as zdata:
            data_num_bytes = zdata.read("word/numbering.xml")
    except (KeyError, Exception):
        return output_bytes

    data_root = etree.fromstring(data_num_bytes)

    # Build: data numId → abstractNumId
    num_to_abstract: dict[int, int] = {}
    for num_elem in data_root.findall(f"{WNS}num"):
        nid = int(num_elem.get(f"{WNS}numId", 0))
        if nid in needed_orig:
            ref = num_elem.find(f"{WNS}abstractNumId")
            if ref is not None:
                num_to_abstract[nid] = int(ref.get(f"{WNS}val", 0))

    needed_abstract = set(num_to_abstract.values())

    # Read / create output numbering.xml
    out_num_bytes = existing.get("word/numbering.xml")
    if out_num_bytes:
        out_root = etree.fromstring(out_num_bytes)
    else:
        out_root = etree.fromstring(
            b'<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
        )

    # Max existing IDs in output
    max_abstract = max(
        (int(e.get(f"{WNS}abstractNumId", 0)) for e in out_root.findall(f"{WNS}abstractNum")),
        default=0,
    )
    max_num = max(
        (int(e.get(f"{WNS}numId", 0)) for e in out_root.findall(f"{WNS}num")),
        default=0,
    )

    # Copy abstractNum entries (must come before <w:num> in the XML)
    abstract_id_map: dict[int, int] = {}  # old → new
    for ab_elem in data_root.findall(f"{WNS}abstractNum"):
        old_id = int(ab_elem.get(f"{WNS}abstractNumId", 0))
        if old_id not in needed_abstract:
            continue
        max_abstract += 1
        abstract_id_map[old_id] = max_abstract
        new_ab = copy.deepcopy(ab_elem)
        new_ab.set(f"{WNS}abstractNumId", str(max_abstract))
        # Remove numStyleLink — it references a style name that may not exist
        for lnk in new_ab.findall(f"{WNS}numStyleLink"):
            new_ab.remove(lnk)
        # Insert before the first <w:num> element so ordering is correct
        first_num = out_root.find(f"{WNS}num")
        if first_num is not None:
            first_num.addprevious(new_ab)
        else:
            out_root.append(new_ab)

    # Add <w:num> entries
    numid_map: dict[int, int] = {}  # orig data numId → new output numId
    for orig_nid, orig_abstract in num_to_abstract.items():
        new_abstract = abstract_id_map.get(orig_abstract)
        if new_abstract is None:
            continue
        max_num += 1
        numid_map[orig_nid] = max_num
        num_elem = etree.SubElement(out_root, f"{WNS}num")
        num_elem.set(f"{WNS}numId", str(max_num))
        ab_ref = etree.SubElement(num_elem, f"{WNS}abstractNumId")
        ab_ref.set(f"{WNS}val", str(new_abstract))

    # Rewrite sentinel numIds in document.xml
    for orig_nid, new_nid in numid_map.items():
        sentinel = orig_nid + _NUMID_OFFSET
        doc_text = doc_text.replace(
            f'<w:numId w:val="{sentinel}"',
            f'<w:numId w:val="{new_nid}"',
        )

    # Persist numbering.xml — ensure relationship exists
    existing["word/numbering.xml"] = etree.tostring(
        out_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    existing["word/document.xml"] = doc_text.encode("utf-8")

    # Add numbering relationship if not present
    rels_text = existing.get("word/_rels/document.xml.rels", b"").decode("utf-8")
    _NUM_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    if _NUM_REL not in rels_text:
        existing_nums = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_text)]
        new_rid = f"rId{max(existing_nums, default=0) + 1}"
        insert_at = rels_text.rfind("</")
        rels_text = (
            rels_text[:insert_at]
            + f'  <Relationship Id="{new_rid}" Type="{_NUM_REL}" Target="numbering.xml"/>\n'
            + rels_text[insert_at:]
        )
        existing["word/_rels/document.xml.rels"] = rels_text.encode("utf-8")

    out_buf = BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in existing.items():
            zout.writestr(name, data)

    return out_buf.getvalue()


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def fill_template(
    template_path: str,
    data_doc_path: str,
    matches: list[dict],
    output_path: str,
    product_name: str | None = None,
) -> None:
    """
    For each successful match extract the corresponding section from
    *data_doc_path* and insert its text + images after the template heading.

    Parameters
    ----------
    template_path : str   – path to the blank / partial template .docx
    data_doc_path : str   – path to the source data .docx
    matches       : list  – output of heading_matcher.match_headings()
    output_path   : str   – where to write the filled document
    """
    template = Document(template_path)

    # Collects image parts from all inserted tables for zip-level copying later
    all_image_parts: list[dict] = []

    _prompts_file = os.path.join(os.path.dirname(__file__), "..", "prompts.json")
    with open(_prompts_file, "r", encoding="utf-8") as _pf:
        _prompts_cfg = json.load(_pf)
    _keep_n: dict[str, int] = {
        k.lower(): v
        for k, v in _prompts_cfg.get("sections_keep_first_n_tables", {}).items()
    }
    # Auto-detect compliance heading and apply sections_keep_first_n_auto
    _auto_keep: int = _prompts_cfg.get("sections_keep_first_n_auto", 0)
    if _auto_keep and os.path.isfile(template_path):
        _auto_h = detect_compliance_heading(template_path)
        if _auto_h:
            _keep_n.setdefault(_auto_h.lower(), _auto_keep)
    _remap_cols: dict[str, list] = {
        k.lower(): v
        for k, v in _prompts_cfg.get("sections_remap_table_columns", {}).items()
    }
    # filter_headings: for these sections, skip any extracted item that is a
    # heading paragraph (avoids sub-headings like "Test Administration" appearing
    # in the Compliance Checklist section).
    _filter_heading_sections: set[str] = {
        h.lower() for h in _prompts_cfg.get("sections_filter_heading_content", [])
    }
    # keep_template_sections: for these headings, preserve the template's own
    # content entirely — do NOT extract from the data doc or clear the template.
    # Post-processing passes still run and fill placeholders / remove guidance.
    _keep_tmpl_sections: set[str] = {
        h.lower() for h in _prompts_cfg.get("sections_keep_template_content", [])
    }

    valid_matches = [
        m for m in matches
        if m["matched_heading"] is not None
        and m["template_heading"]["text"].lower() not in _keep_tmpl_sections
    ]
    valid_matches.sort(
        key=lambda m: m["template_heading"]["paragraph_index"],
        reverse=True,
    )

    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    # Build a set of heading style IDs (e.g. "Heading1", "Heading2", …)
    heading_style_ids: set[str] = set()
    for s in template.styles:
        if s.name and s.name.startswith("Heading"):
            heading_style_ids.add(s.style_id)

    def _is_heading_elem(elem) -> bool:
        """Return True if *elem* is a paragraph with a Heading style."""
        if elem.tag != f"{{{ns}}}p":
            return False
        pPr = elem.find(f"{{{ns}}}pPr")
        if pPr is None:
            return False
        pStyle = pPr.find(f"{{{ns}}}pStyle")
        if pStyle is None:
            return False
        return pStyle.get(f"{{{ns}}}val", "") in heading_style_ids

    def _is_layout_para(elem) -> bool:
        """True if elem is a paragraph used only for page layout (page break or sectPr)."""
        if elem.tag != f"{{{ns}}}p":
            return False
        for br in elem.iter(f"{{{ns}}}br"):
            if br.get(f"{{{ns}}}type") == "page":
                return True
        pPr = elem.find(f"{{{ns}}}pPr")
        if pPr is not None and pPr.find(f"{{{ns}}}sectPr") is not None:
            return True
        return False

    def _clear_section(heading_elem) -> None:
        """
        Remove every body element immediately after *heading_elem* up to
        (but not including) the next heading element at any level.
        Preserves trailing page-break / sectPr paragraphs so that section
        separation is maintained in the filled document.
        """
        body = heading_elem.getparent()
        body_children = list(body)
        try:
            start = body_children.index(heading_elem)
        except ValueError:
            return
        to_remove = []
        for elem in body_children[start + 1:]:
            if _is_heading_elem(elem):
                break
            to_remove.append(elem)
        # Preserve ALL layout paragraphs (page-break paras and inline sectPr
        # paras).  The inline sectPr ones carry the template's header/footer
        # rId references; deleting them would break the page header/footer.
        to_remove = [e for e in to_remove if not _is_layout_para(e)]
        for elem in to_remove:
            body.remove(elem)

    for match in valid_matches:
        tmpl_h = match["template_heading"]
        data_h = match["matched_heading"]

        content = extract_section(
            data_doc_path,
            data_h["paragraph_index"],
            data_h["level"],
        )

        if not content["items"]:
            print(f"  [SKIP] No content found under: '{data_h['text']}'")
            continue

        # Anchor: start from the heading's raw lxml element
        anchor = template.paragraphs[tmpl_h["paragraph_index"]]._p
        tmpl_h_lower = tmpl_h["text"].lower()

        if tmpl_h_lower in _remap_cols:
            # Column-remap: keep the template's table structure (header row +
            # column widths) and fill it with mapped rows from the data doc.
            col_indices = _remap_cols[tmpl_h_lower]

            # Save a deep copy of the template's first table BEFORE clearing
            saved_tmpl_tbl = None
            body_pre = list(anchor.getparent())
            anchor_idx = body_pre.index(anchor)
            for elem in body_pre[anchor_idx + 1:]:
                if _is_heading_elem(elem):
                    break
                if elem.tag == f"{{{ns}}}tbl":
                    saved_tmpl_tbl = copy.deepcopy(elem)
                    break

            # Clear template placeholder content
            _clear_section(anchor)

            # Find first table in extracted data content
            data_tbl_xml = None
            for item in content["items"]:
                if item["type"] == "table":
                    data_tbl_xml = item["xml"]
                    all_image_parts.extend(item.get("image_parts", []))
                    break

            if saved_tmpl_tbl is not None and data_tbl_xml is not None:
                data_tbl_elem = etree.fromstring(data_tbl_xml)
                _strip_inline_sectpr(data_tbl_elem)
                merged = _build_remapped_table(saved_tmpl_tbl, data_tbl_elem, col_indices, ns)
                anchor = _insert_xml_after(anchor, etree.tostring(merged, encoding="unicode").encode())

            continue  # section fully handled — skip normal insert

        elif tmpl_h_lower in _keep_n:
            # Partial clear: keep first N tables after the heading, remove the rest.
            n = _keep_n[tmpl_h_lower]
            body = anchor.getparent()
            body_children = list(body)
            start = body_children.index(anchor)
            tables_seen = 0
            after_kept = None   # last kept element becomes the new anchor
            to_remove = []
            for elem in body_children[start + 1:]:
                if _is_heading_elem(elem):
                    break
                # Preserve layout paragraphs (page breaks / inline sectPr) —
                # never delete them.  Do NOT update after_kept for them so
                # the anchor always points to the last *kept table*, not a
                # layout paragraph inserted after it.
                if _is_layout_para(elem):
                    continue
                if elem.tag == f"{{{ns}}}tbl":
                    tables_seen += 1
                    if tables_seen <= n:
                        after_kept = elem
                        continue
                to_remove.append(elem)
            for elem in to_remove:
                body.remove(elem)
            if after_kept is not None:
                anchor = after_kept
        else:
            # Full clear of all template placeholder content under this heading
            _clear_section(anchor)

        # Filter flags for this heading
        filter_hdgs = tmpl_h_lower in _filter_heading_sections

        for item in content["items"]:
            if filter_hdgs and item["type"] == "paragraph":
                p_elem = etree.fromstring(item["xml"])
                pPr = p_elem.find(f"{{{ns}}}pPr")
                if pPr is not None:
                    pStyle = pPr.find(f"{{{ns}}}pStyle")
                    if pStyle is not None and pStyle.get(f"{{{ns}}}val", "") in heading_style_ids:
                        continue
            anchor = _insert_xml_after(anchor, item["xml"])
            all_image_parts.extend(item.get("image_parts", []))

    # ── Pre-heading content (text before the first heading in the data doc) ─
    # E.g. "Philips Compliance Data Record  ISO …" lines that sit at the very
    # top of the source file, outside any section.
    pre = extract_pre_heading_content(data_doc_path)
    if pre["items"]:
        # Find the first heading element in the template body and insert
        # all pre-heading items immediately before it.
        tmpl_body = template.element.body
        first_heading_elem = None
        for elem in list(tmpl_body):
            if _is_heading_elem(elem):
                first_heading_elem = elem
                break
        if first_heading_elem is not None:
            # Insert in reverse order so the final sequence is preserved
            for item in reversed(pre["items"]):
                new_elem = copy.deepcopy(etree.fromstring(item["xml"]))
                _strip_inline_sectpr(new_elem)
                first_heading_elem.addprevious(new_elem)
                all_image_parts.extend(item.get("image_parts", []))
        else:
            # No headings at all — append at end of body
            ref = list(tmpl_body)[-1]
            for item in pre["items"]:
                ref = _insert_xml_after(ref, item["xml"])
                all_image_parts.extend(item.get("image_parts", []))

    # ── Post-processing ────────────────────────────────────────────────────
    extra_image_parts = post_processor.apply_all(template, product_name, data_doc_path=data_doc_path, template_path=template_path)
    all_image_parts.extend(extra_image_parts)

    # ── Save: first to bytes so we can do zip-level image patching ─────────
    buf = BytesIO()
    template.save(buf)
    final_bytes = _copy_table_images(buf.getvalue(), all_image_parts)
    final_bytes = _merge_comments(final_bytes, data_doc_path)
    final_bytes = _merge_numbering(final_bytes, data_doc_path)

    with open(output_path, "wb") as fh:
        fh.write(final_bytes)

    print(f"\nSaved -> {output_path}")
