"""
main.py  –  Automated CDR Tool (BAAI BGE)
-----------------------------------------
Full pipeline:
  1. Extract headings from each template + the data doc
  2. Match headings semantically via BAAI/bge-base-en-v1.5
  3. Extract matched sections (text + images) from the data doc
  4. Fill each template and save outputs named op1_…, op2_…, op3_…

Output file naming
------------------
The base name is derived from the input data filename:
  • If *product_name* (from prompts.json or --product-name) appears in the
    filename, it and everything after it (including ".docx") are stripped.
    → output: op{N}_<stem>  (no extension in the base; .docx is appended)
    e.g. "Input1_D001352871 Test Record ISO 17664-2 (2021) Azurion HW R3.docx"
         → op1_Input1_D001352871 Test Record ISO 17664-2 (2021).docx
  • Otherwise the full basename (with ".docx") is kept as the base.
    → output: op{N}_<full_filename>  (already ends in .docx)
    e.g. "Input4_ D001024006 CDR EN 62479 (2010) ProductName RX.Y Rev C.docx"
         → op1_Input4_ D001024006 CDR EN 62479 (2010) ProductName RX.Y Rev C.docx

Usage
-----
  # Single template (backward-compatible)
  python main.py --template <template.docx> --input <data.docx>

  # Multiple templates → op1_, op2_, op3_ outputs
  python main.py --templates t1.docx t2.docx t3.docx --input <data.docx>
"""

import argparse
import json
import os

from src.heading_extractor import extract_headings
from src.heading_matcher   import load_model, match_headings
from src.template_filler   import fill_template

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ---------------------------------------------------------------------------
# Auto-detect product name from a .docx's Word core properties
# ---------------------------------------------------------------------------

def auto_detect_product_name(doc_path: str) -> str | None:
    """Read product name from the document's Word core properties,
    then fall back to parsing from the filename.

    Strategy:
      1. Core properties: title → subject → description
      2. Filename parsing: text after '(YYYY) ' at end of stem
         e.g. 'Input1_D001352871 Test Record ISO 17664-2 (2021) Azurion HW R3.docx'
              → 'Azurion HW R3'
    """
    import re
    try:
        from docx import Document
        doc = Document(doc_path)
        cp  = doc.core_properties
        for attr in ("title", "subject", "description"):
            val = (getattr(cp, attr, None) or "").strip()
            if val:
                return val
    except Exception:
        pass
    # Filename fallback: extract product name after '(YEAR) ' pattern
    stem = os.path.splitext(os.path.basename(doc_path))[0]
    m = re.search(r'\((\d{4})\)\s+(.+)$', stem)
    if m:
        return m.group(2).strip()
    return None


# ---------------------------------------------------------------------------
# Output filename helpers
# ---------------------------------------------------------------------------

def derive_output_base(input_path: str, product_name: str | None) -> str:
    """
    Derive the output file base name from the input data filename.

    Rules
    -----
    - If *product_name* is found in the basename:
        Strip everything from *product_name* onwards (incl. trailing space
        and '.docx') → returns the stem **without** an extension.
        Caller must append '.docx' to get the actual file path.
    - Otherwise:
        Returns the full basename including '.docx'.
        Caller uses it as-is (it already ends in '.docx').
    """
    basename = os.path.basename(input_path)
    if product_name and product_name in basename:
        idx = basename.find(product_name)
        return basename[:idx].rstrip()
    return basename


def build_output_path(output_dir: str, op_index: int, base: str) -> str:
    """
    Build the full output path for op{op_index}.

    *base* is the value returned by :func:`derive_output_base`:
    - If it already ends in '.docx', prefix op{N}_ and use as-is.
    - Otherwise append '.docx' after the prefix.
    """
    prefix = f"op{op_index}_"
    if base.lower().endswith(".docx"):
        filename = f"{prefix}{base}"
    else:
        filename = f"{prefix}{base}.docx"
    return os.path.join(output_dir, filename)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Automated CDR Tool — fill template(s) from a data .docx"
    )
    tmpl_group = parser.add_mutually_exclusive_group(required=True)
    tmpl_group.add_argument(
        "--template",
        help="Path to a single template .docx file",
    )
    tmpl_group.add_argument(
        "--templates", nargs="+",
        metavar="TEMPLATE",
        help="Paths to multiple template .docx files (generates op1_, op2_, … outputs)",
    )
    parser.add_argument(
        "--input", required=True,
        help="Path to the input data .docx file",
    )
    parser.add_argument(
        "--output", default=None,
        help="Explicit output path — only used when a single template is provided",
    )
    parser.add_argument(
        "--threshold", type=float, default=None,
        help="Minimum match score 0.0–1.0 (default: read from prompts.json 'threshold')",
    )
    parser.add_argument(
        "--product-name", default=None, dest="product_name",
        help="Product name override (default: read from prompts.json 'product_name')",
    )
    return parser.parse_args()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    args = parse_args()

    data_path = os.path.abspath(args.input)

    # ── Resolve template list ──────────────────────────────────────────────
    if args.templates:
        template_paths = [os.path.abspath(t) for t in args.templates]
    else:
        template_paths = [os.path.abspath(args.template)]

    # ── Load config ────────────────────────────────────────────────────────
    _prompts_path = os.path.join(BASE_DIR, "prompts.json")
    with open(_prompts_path, "r", encoding="utf-8") as _pf:
        _prompts = json.load(_pf)
    product_name = args.product_name or auto_detect_product_name(data_path)
    if product_name:
        print(f"  Product name      : {product_name} (auto-detected)")
    else:
        print("  Product name      : not detected — header/footer placeholders will not be replaced")
    threshold    = args.threshold if args.threshold is not None else float(_prompts.get("threshold", 0.6))

    # ── Output directory ───────────────────────────────────────────────────
    output_dir = os.path.join(BASE_DIR, "output")
    os.makedirs(output_dir, exist_ok=True)

    # ── Derive output base from input filename ─────────────────────────────
    base = derive_output_base(data_path, product_name)

    # ── Validate input ─────────────────────────────────────────────────────
    if not os.path.isfile(data_path):
        print(f"[ERROR] Input data doc not found: {data_path}")
        return

    # ── Step 1: Extract data-doc headings (once) ───────────────────────────
    print("=" * 60)
    print("STEP 1 — Extracting data-doc headings")
    print("=" * 60)

    data_headings = extract_headings(data_path)
    print(f"  Data doc headings : {len(data_headings)}")

    if not data_headings:
        print("\n[ERROR] No headings found in the data document. Aborting.")
        return

    # ── Step 2: Load embedding model (once) ───────────────────────────────
    print("\n" + "=" * 60)
    print(f"STEP 2 — Loading {_prompts.get('model_name', 'BAAI/bge-base-en-v1.5')}")
    print("=" * 60)

    model = load_model()

    # ── Process each template ──────────────────────────────────────────────
    for idx, template_path in enumerate(template_paths, start=1):

        # Determine output path
        if args.output and len(template_paths) == 1:
            output_path = os.path.abspath(args.output)
        else:
            output_path = build_output_path(output_dir, idx, base)

        # Auto-increment op index: if op1_ exists, try op2_, op3_, …
        if os.path.isfile(output_path):
            next_op = idx + 1
            while os.path.isfile(build_output_path(output_dir, next_op, base)):
                next_op += 1
            output_path = build_output_path(output_dir, next_op, base)

        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        print("\n" + "=" * 60)
        print(f"TEMPLATE {idx}/{len(template_paths)}: {os.path.basename(template_path)}")
        print(f"OUTPUT             : {os.path.basename(output_path)}")
        print("=" * 60)

        if not os.path.isfile(template_path):
            print(f"[ERROR] Template not found: {template_path}")
            continue

        # Step A: Extract template headings
        print("\nExtracting template headings …")
        template_headings = extract_headings(template_path)
        print(f"  Template headings : {len(template_headings)}")

        if not template_headings:
            print(f"[WARN] No headings found in template, skipping.")
            continue

        # Step B: Match headings
        print("\nMatching headings …")
        matches = match_headings(
            template_headings,
            data_headings,
            model,
            threshold=threshold,
        )

        matched_count = sum(1 for m in matches if m["matched_heading"] is not None)
        print(f"  Matched: {matched_count} / {len(matches)}\n")

        col_t, col_d = 45, 40
        print(f"  {'Template heading':<{col_t}}  {'Score':>6}  Data-doc heading")
        print(f"  {'-'*col_t}  {'------'}  {'-'*col_d}")
        for m in matches:
            tmpl  = m["template_heading"]["text"][:col_t - 1]
            score = f"{m['score']:.2f}"
            data  = (
                m["matched_heading"]["text"][:col_d - 1]
                if m["matched_heading"]
                else "— NO MATCH —"
            )
            print(f"  {tmpl:<{col_t}}  {score:>6}  {data}")

        # Step C: Fill template → save output
        print(f"\nFilling template …")
        fill_template(template_path, data_path, matches, output_path, product_name=product_name)
        print(f"  Saved -> {output_path}")

    print("\nDone.")


if __name__ == "__main__":
    main()
