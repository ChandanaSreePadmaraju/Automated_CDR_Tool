"""
heading_matcher.py
------------------
Use BAAI/bge-base-en-v1.5 to semantically match every template heading
to the best-fitting heading in the data document.
"""

import json
import os
import numpy as np
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity

# Load all prompts from the shared prompts.json file
_PROMPTS_FILE = os.path.join(os.path.dirname(__file__), "..", "prompts.json")
with open(_PROMPTS_FILE, "r", encoding="utf-8") as _f:
    _PROMPTS = json.load(_f)

# BGE retrieval prefix — prepended to template headings before encoding
_QUERY_PREFIX: str = _PROMPTS["query_prefix"]

# Model name is read from prompts.json so it can be changed without touching code
_MODEL_NAME: str = _PROMPTS.get("model_name", "BAAI/bge-base-en-v1.5")


def load_model() -> SentenceTransformer:
    """Download (first run only) and return the BGE model."""
    print(f"Loading model: {_MODEL_NAME}")
    return SentenceTransformer(_MODEL_NAME)


def match_headings(
    template_headings: list[dict],
    data_headings: list[dict],
    model: SentenceTransformer,
    threshold: float,
) -> list[dict]:
    """
    For every template heading find the closest heading in the data doc.

    Returns a list (same length as *template_headings*) where each entry is:
    {
        "template_heading": dict   – original template heading entry,
        "matched_heading":  dict   – best data-doc heading, or None if below threshold,
        "score":            float  – cosine similarity (0–1),
    }

    Parameters
    ----------
    template_headings : list[dict]  – output of heading_extractor.extract_headings() on template
    data_headings     : list[dict]  – output of heading_extractor.extract_headings() on data doc
    model             : SentenceTransformer – loaded BGE model
    threshold         : float – minimum score to accept a match (supplied by caller; see DEFAULT_THRESHOLD in main.py)
    """
    if not template_headings or not data_headings:
        return [
            {"template_heading": h, "matched_heading": None, "score": 0.0}
            for h in template_headings
        ]

    # Query side uses BAAI's retrieval prefix; corpus side is plain text
    query_texts  = [_QUERY_PREFIX + h["text"] for h in template_headings]
    corpus_texts = [h["text"] for h in data_headings]

    query_embs  = model.encode(query_texts,  normalize_embeddings=True, show_progress_bar=False)
    corpus_embs = model.encode(corpus_texts, normalize_embeddings=True, show_progress_bar=False)

    # Shape: (len(template_headings), len(data_headings))
    sim_matrix: np.ndarray = cosine_similarity(query_embs, corpus_embs)

    # Build a lookup for explicit heading mappings from prompts.json (case-insensitive)
    explicit: dict[str, str] = {
        k.lower(): v.lower()
        for k, v in _PROMPTS.get("heading_mappings", {}).items()
    }
    data_by_lower: dict[str, dict] = {h["text"].lower(): h for h in data_headings}

    matches: list[dict] = []
    for i, tmpl_h in enumerate(template_headings):
        # Check for an explicit override first
        forced_target = explicit.get(tmpl_h["text"].lower())
        if forced_target and forced_target in data_by_lower:
            matches.append({
                "template_heading": tmpl_h,
                "matched_heading":  data_by_lower[forced_target],
                "score":            1.0,
            })
            continue

        best_idx   = int(np.argmax(sim_matrix[i]))
        best_score = float(sim_matrix[i][best_idx])

        matches.append({
            "template_heading": tmpl_h,
            "matched_heading":  data_headings[best_idx] if best_score >= threshold else None,
            "score":            round(best_score, 4),
        })

    return matches
