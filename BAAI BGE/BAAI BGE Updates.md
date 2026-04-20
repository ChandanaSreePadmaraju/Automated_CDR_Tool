# BAAI BGE Updates

---

## April 20, 2026

### Cleanup

- Removed dev/debug scripts: `diag.py`, `quick_check.py`, `verify_output.py`
- Updated `.gitignore`: entire `output/` excluded (no pinned file), added `~$*` for Word lock files
- Output is now auto-versioned at runtime (`Filled_CDR.docx`, `_v2`, `_v3`, …)

---

## April 14, 2026

**Selected Model:** `bge-base-en-v1.5` → fast + accurate enough

---

### Clarification on Image Handling with BGE

- `bge-base-en-v1.5` is a **text-only** embedding model — it **cannot process or paste images**.
- It handles **only the heading matching step** (text-to-text semantic similarity).
- Image extraction and pasting must be handled **separately** via document parsing libraries.

---

### Workflow

Template heading  ──► BGE (match) ──► Find section in data doc
                                              │
                                    python-docx / PyMuPDF
                                              │
                               Extract text + images from that section
                                              │
                               Paste into template under that heading

---

### Step-by-Step Breakdown

| Step | Task                                                        | Tool                                                        |
|:----:|:------------------------------------------------------------|:------------------------------------------------------------|
|  1   | Match headings between template and data doc                | BGE (`bge-base-en-v1.5`) — semantic text similarity         |
|  2   | Extract text + images from matched section in data doc      | `python-docx` / `PyMuPDF` — document parsing                |
|  3   | Paste extracted text + images under matched heading in template | `python-docx` — document writing                        |

> **Key Insight:** BGE finds *where* to pull from (heading match). `python-docx` / `PyMuPDF` does the actual extraction and pasting of both text and images.

---

## April 15, 2026

### Final Implementation — Full Pipeline

---

### Project Structure

```
BAAI BGE/
├── main.py                    # CLI entry point — runs the full pipeline
├── prompts.json               # All config (query prefix, heading mappings, post-processing flags)
├── requirements.txt           # Python dependencies
├── src/
│   ├── heading_extractor.py   # Step 1 — extract headings from any .docx
│   ├── heading_matcher.py     # Step 2 — BGE semantic heading matching
│   ├── content_extractor.py   # Step 3 — extract section content as raw XML
│   ├── template_filler.py     # Step 4 — insert content into template, save output
│   └── post_processor.py      # Step 5 — 11 post-processing cleanup passes
└── output/                    # Generated output (git-ignored)
```

---

### `main.py` — CLI Entry Point

- **argparse** CLI with flags: `--template`, `--input`, `--output`, `--threshold`, `--product-name`
- Default threshold: `0.55` cosine similarity
- Default output: `output/Filled_CDR.docx`
- Orchestrates all 4 pipeline steps in sequence

**Run command:**
```powershell
python main.py `
  --template "Template_D001024021 CDR ISO 17664-2 (2021) ProductName RX.Y Rev C.docx" `
  --input "Input_D001352871 Test Record ISO 17664-2 (2021) Azurion HW R3.docx" `
  --product-name "Azurion HW R3"
```

---

### `src/heading_extractor.py` — Step 1

- Reads every paragraph in a `.docx` using `python-docx`
- Returns only paragraphs whose style starts with `"Heading"`
- Skips blank headings
- Returns: `{ "text", "level", "paragraph_index" }` per heading — works on both template and data doc

---

### `src/heading_matcher.py` — Step 2

- Loads `BAAI/bge-base-en-v1.5` via `sentence-transformers`
- Prepends BGE retrieval prefix (`"Represent this sentence for searching relevant passages: "`) to template headings before encoding
- Computes cosine similarity matrix between template headings and data doc headings
- **Explicit override support** — `heading_mappings` in `prompts.json` forces specific matches (score = 1.0) bypassing semantic matching
  - Example: `"Compliance Checklist"` → `"Test Record"` (semantic score was below threshold)
  - Example: `"Record history"` → `"Document Revision History"`
- **`skip_template_headings`** — list of template headings to skip entirely (content kept as-is from template)
- Returns best match per template heading, or `None` if below threshold

---

### `src/content_extractor.py` — Step 3

- Extracts all body elements from a section (from named heading to next same-or-higher-level heading)
- **Raw XML extraction** — every `<w:p>` and `<w:tbl>` is deep-copied and serialised to bytes
  - Preserves all formatting: bold, italic, colour, font, nested tables, merged cells, inline images
  - No hardcoded structure assumptions (no "paragraphs then tables then paragraphs")
- **Placeholder rId system** — at extraction time, every `r:embed="rIdXX"` in image blips is rewritten to a globally-unique `rId_CDRCOPY_NNNNNN` value
  - Prevents any data-doc rId from accidentally overwriting a template rId (e.g. cover page logo `rId12`)
  - Uses a module-level `itertools.count` so IDs are unique across all `extract_section()` calls per run
- Returns: list of `{ "type", "xml", "image_parts" }` items — one per body element

---

### `src/template_filler.py` — Step 4

- **Reverse-order insertion** — sections processed from last heading to first so earlier insertions don't shift paragraph indices
- **Single insertion function** `_insert_xml_after(ref_elem, xml_bytes)` — handles any element type (`<w:p>` or `<w:tbl>`) via `lxml`
- Strips Word comment markers (`<w:commentRangeStart>`, `<w:commentRangeEnd>`, `<w:commentReference>`) from inserted XML to avoid rId conflicts
- **`_copy_table_images()`** — zip-level image injection:
  - Deduplicates images by `placeholder_rid`
  - Adds image bytes to `word/media/img_copied_rIdXX.ext`
  - Adds `<Relationship>` entries to `word/_rels/document.xml.rels`
  - Replaces **only** `rId_CDRCOPY_...` strings in `document.xml` — never touches template-original rIds
- Calls all 11 `post_processor` passes before saving
- **`sections_keep_first_n_tables`** — for specified sections, keeps only the first N tables (e.g. `"Compliance Checklist": 1`)
- **`sections_filter_heading_content`** — strips sub-heading content from specified sections before insertion

---

### `src/post_processor.py` — Step 5 (11 Passes)

All behaviour driven by `"post_processing"` block in `prompts.json` — nothing hardcoded.

| Pass | Function | What it does |
|:----:|:---------|:-------------|
| 1 | `remove_styled_paragraphs` | Removes all paragraphs with styles listed in `styles_to_remove` (e.g. `"Guidance"`) |
| 2 | `remove_template_instructions` | Removes multi-paragraph `< ... >` instruction blocks left from the template |
| 3 | `remove_empty_table_rows` | Removes empty non-header table rows |
| 4 | `set_all_text_black` | Forces all text to black; removes GuidanceChar rStyle overrides; inserts explicit `<w:color val="000000"/>` |
| 5 | `sort_tables_alphabetically` | Sorts data rows of specified tables alphabetically by first column (e.g. References, Definitions) |
| 6 | `remove_sections` | Removes entire sections (heading + all content) by name (e.g. `"Pre-created CDR history"`) |
| 7 | `fill_header_footer_placeholders` | Replaces product name placeholders (e.g. `<ProductName RX.Y>`) in body, headers, and footers |
| 8 | `remove_preamble_before_first_heading` | Removes instruction-page content before first heading; preserves cover and TOC page-break paragraphs |
| 9 | `strip_inline_angle_brackets` | Removes standalone `<` and `>` `<w:r>` runs left as bracket artifacts in body paragraphs |
| 10 | `remove_paragraphs_with_text` | Removes body paragraphs containing specific strings (e.g. template metadata notes) |
| 11 | `inject_definitions_fixed_rows` | Ensures fixed abbreviation rows (e.g. CDR, ISO, IGT-S) always exist in the Definitions table |

**Pass 7 details — multi-run placeholder collapse:**
- Uses `lxml` directly (not `python-docx` `.runs` API) to avoid `AttributeError` on raw-inserted XML elements
- Word splits placeholders like `<ProductName RX.Y>` across multiple `<w:r>` runs: `'<'`, `'ProductName '`, `'RX.Y'`, `'>'`
- Algorithm: concatenates all run texts in a paragraph, finds the placeholder span, collapses the matching runs into a single run with the replacement text, removes the remaining runs
- Scans: `doc.element.body` (covers body paragraphs, cover page table, all body tables) + all header/footer parts

---

### `prompts.json` — Configuration File

```json
{
    "model_name": "BAAI/bge-base-en-v1.5",
    "query_prefix": "Represent this sentence for searching relevant passages: ",

    "heading_mappings": {
        "Compliance Checklist": "Test Record",
        "Record history": "Document Revision History"
    },

    "skip_template_headings": ["Purpose", "Scope"],

    "sections_keep_first_n_tables": {
        "Compliance Checklist": 1
    },

    "sections_filter_heading_content": ["Compliance Checklist"],

    "post_processing": {
        "styles_to_remove": ["Guidance"],
        "remove_template_instructions": true,
        "strip_inline_angle_brackets": true,
        "remove_paragraphs_with_text": [
            "Note: for template information, see custom properties of this document"
        ],
        "remove_preamble_before_first_heading": true,
        "remove_empty_table_rows": true,
        "set_all_text_black": true,
        "sort_table_alphabetically_under_headings": [
            "References",
            "Definitions & abbreviations"
        ],
        "definitions_fixed_rows": [
            ["CDR",   "Compliance Data Record"],
            ["ISO",   "International Standardization Organization"],
            ["IGT-S", "Image Guided Therapy - Systems"]
        ],
        "definitions_heading": "Definitions & abbreviations",
        "remove_sections": ["Pre-created CDR history"],
        "header_footer_placeholders": ["<ProductName RX.Y>"]
    }
}
```

**Configurable without code changes:**
- Add more heading overrides in `heading_mappings`
- Add headings to skip in `skip_template_headings` (template content kept as-is)
- Control how many tables to keep per section via `sections_keep_first_n_tables`
- Add more placeholder strings in `header_footer_placeholders`
- Add/remove section names in `remove_sections`
- Add/remove table headings in `sort_table_alphabetically_under_headings`
- Add/remove fixed rows in `definitions_fixed_rows`
- Add strings to purge in `remove_paragraphs_with_text`

---

### Bugs Fixed

| Bug | Root Cause | Fix |
|---|---|---|
| `"Compliance Checklist"` section not filled | Semantic score 0.509 < 0.55 threshold | `heading_mappings` override in `prompts.json` + explicit check in `heading_matcher.py` |
| Hardcoded document structure assumed | `<w:p>` extracted as plain text; images emitted as separate items | All elements serialised as raw XML bytes; single `_insert_xml_after` handles any element |
| Cover page `<ProductName RX.Y>` not replaced | Only headers/footers scanned; placeholder split across 4 `<w:r>` runs by Word | lxml-based multi-run collapse algorithm; added `doc.element.body` scan |
| Cover page logo replaced with data doc image | Global `str.replace('r:embed="rId12"')` overwrote template's `rId12` | Placeholder rIds (`rId_CDRCOPY_...`) injected at extraction time; only those replaced at zip level |

---

### Final Output

- Output is auto-versioned: `output/Filled_CDR.docx`, `output/Filled_CDR_v2.docx`, `output/Filled_CDR_v3.docx` …
- The entire `output/` folder is git-ignored — generated files are never committed
- Run the pipeline to regenerate the output at any time
