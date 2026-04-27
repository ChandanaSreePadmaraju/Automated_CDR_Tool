"""
main.py  –  Automated CDR Tool (BAAI BGE)
-----------------------------------------
Full pipeline:
  1. Extract headings from template + data doc
  2. Match headings semantically via BAAI/bge-base-en-v1.5
  3. Extract matched sections (text + images) from the data doc
  4. Fill the template and save the output to output/Filled_CDR.docx

Usage
-----
  python main.py --template <template.docx> --input <data.docx>
  python main.py --template <template.docx> --input <data.docx> --output <out.docx> --threshold 0.55
"""

import argparse
import json
import os
import re

from src.heading_extractor import extract_headings
from src.heading_matcher   import load_model, match_headings
from src.template_filler   import fill_template

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Automated CDR Tool — fill a template .docx from a data .docx"
    )
    parser.add_argument(
        "--template", required=True,
        help="Path to the template .docx file",
    )
    parser.add_argument(
        "--input", required=True,
        help="Path to the input data .docx file",
    )
    parser.add_argument(
        "--output", default=None,
        help="Path for the filled output .docx (default: output/Filled_CDR.docx)",
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


def main() -> None:
    args = parse_args()

    template_path = os.path.abspath(args.template)
    data_path     = os.path.abspath(args.input)

    # Load product_name from prompts.json; CLI arg overrides if provided
    _prompts_path = os.path.join(BASE_DIR, "prompts.json")
    with open(_prompts_path, "r", encoding="utf-8") as _pf:
        _prompts = json.load(_pf)
    product_name = args.product_name or _prompts.get("product_name")
    threshold    = args.threshold if args.threshold is not None else float(_prompts.get("threshold", 0.6))

    if args.output:
        output_path = os.path.abspath(args.output)
    else:
        output_dir  = os.path.join(BASE_DIR, "output")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "Filled_CDR.docx")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # Auto-version: if the file already exists, append _v2, _v3, … until free
    if os.path.isfile(output_path):
        base, ext = os.path.splitext(output_path)
        base = re.sub(r'_v\d+$', '', base)
        version = 2
        while os.path.isfile(f"{base}_v{version}{ext}"):
            version += 1
        output_path = f"{base}_v{version}{ext}"

    if not os.path.isfile(template_path):
        print(f"[ERROR] Template not found: {template_path}")
        return
    if not os.path.isfile(data_path):
        print(f"[ERROR] Input data doc not found: {data_path}")
        return

    # ── Step 1: Heading extraction ─────────────────────────────────────────
    print("=" * 60)
    print("STEP 1 — Extracting headings")
    print("=" * 60)

    template_headings = extract_headings(template_path)
    data_headings     = extract_headings(data_path)

    print(f"  Template headings : {len(template_headings)}")
    print(f"  Data doc headings : {len(data_headings)}")

    if not template_headings:
        print("\n[ERROR] No headings found in the template. Aborting.")
        return
    if not data_headings:
        print("\n[ERROR] No headings found in the data document. Aborting.")
        return

    # ── Step 2: Semantic heading matching ──────────────────────────────────
    print("\n" + "=" * 60)
    print(f"STEP 2 — Loading {_prompts.get('model_name', 'BAAI/bge-base-en-v1.5')} & matching headings")
    print("=" * 60)

    model   = load_model()
    matches = match_headings(
        template_headings,
        data_headings,
        model,
        threshold=threshold,
    )

    matched_count = sum(1 for m in matches if m["matched_heading"] is not None)
    print(f"\n  Matched: {matched_count} / {len(matches)}\n")

    # Print match report table
    col_t = 45
    col_d = 40
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

    # ── Step 3: Fill template ──────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("STEP 3 — Filling template")
    print("=" * 60)

    fill_template(template_path, data_path, matches, output_path, product_name=product_name)

    print("\nDone.")


if __name__ == "__main__":
    main()
