"""
template_filler.py
------------------
Insert content extracted from the data doc into the template doc under each
matched heading, then save the result to *output_path*.

Insertions are processed in **reverse paragraph-index order** so that adding
content after heading N never shifts the indices of headings above it.
"""

import copy
import io
import re
import zipfile
from io import BytesIO

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from src.content_extractor import extract_section
from src import post_processor


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

_COMMENT_TAGS = {
    qn("w:commentRangeStart"),
    qn("w:commentRangeEnd"),
    qn("w:commentReference"),
}


def _strip_comments(elem) -> None:
    """Remove all comment-related elements from *elem* in-place."""
    for tag in _COMMENT_TAGS:
        for node in elem.findall(f".//{tag}"):
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)


def _insert_xml_after(ref_elem, xml_bytes: bytes):
    """
    Deserialise *xml_bytes* into a fresh lxml element, strip comment markers,
    then insert it immediately after *ref_elem*.  Works for any body child
    (<w:p>, <w:tbl>, …).  Returns the newly inserted element.
    """
    new_elem = copy.deepcopy(etree.fromstring(xml_bytes))
    _strip_comments(new_elem)
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

    # Only process matched headings; reverse by template paragraph index so
    # content insertions don't shift paragraph indices of later headings.
    import json, os as _os
    _prompts_file = _os.path.join(_os.path.dirname(__file__), "..", "prompts.json")
    with open(_prompts_file, "r", encoding="utf-8") as _pf:
        _prompts_cfg = json.load(_pf)
    _skip_headings = {h.lower() for h in _prompts_cfg.get("skip_template_headings", [])}
    # keep_first_n: for these headings, keep the first N template tables and
    # only clear what follows them (rather than clearing the whole section).
    _keep_n: dict[str, int] = {
        k.lower(): v
        for k, v in _prompts_cfg.get("sections_keep_first_n_tables", {}).items()
    }
    # filter_headings: for these sections, skip any extracted item that is a
    # heading paragraph (avoids sub-headings like "Test Administration" appearing
    # in the Compliance Checklist section).
    _filter_heading_sections: set[str] = {
        h.lower() for h in _prompts_cfg.get("sections_filter_heading_content", [])
    }

    valid_matches = [
        m for m in matches
        if m["matched_heading"] is not None
        and m["template_heading"]["text"].lower() not in _skip_headings
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
        # Keep trailing layout paragraphs (page breaks / sectPr) in place so
        # the section boundary spacing survives after content replacement.
        while to_remove and _is_layout_para(to_remove[-1]):
            to_remove.pop()
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

        if tmpl_h_lower in _keep_n:
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
            # Optionally skip heading paragraphs from sub-sections
            if filter_hdgs and item["type"] == "paragraph":
                from lxml import etree as _etree
                p_elem = _etree.fromstring(item["xml"])
                pPr = p_elem.find(f"{{{ns}}}pPr")
                if pPr is not None:
                    pStyle = pPr.find(f"{{{ns}}}pStyle")
                    if pStyle is not None and pStyle.get(f"{{{ns}}}val", "") in heading_style_ids:
                        continue
            anchor = _insert_xml_after(anchor, item["xml"])
            all_image_parts.extend(item.get("image_parts", []))

    # ── Post-processing ────────────────────────────────────────────────────
    post_processor.apply_all(template, product_name)

    # ── Save: first to bytes so we can do zip-level image patching ─────────
    buf = io.BytesIO()
    template.save(buf)
    final_bytes = _copy_table_images(buf.getvalue(), all_image_parts)

    with open(output_path, "wb") as fh:
        fh.write(final_bytes)

    print(f"\nSaved -> {output_path}")
