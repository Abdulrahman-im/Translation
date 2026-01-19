"""Alignment service for building dictionary from parallel PowerPoint files."""

import requests
from typing import Dict, List, Tuple, Optional
from .pptx_parser import extract_text_from_pptx
from .translator import API_URL, API_KEY
from .dictionary import add_entries_bulk


def extract_texts_by_slide(file_path: str) -> Dict[int, List[str]]:
    """
    Extract texts grouped by slide number.

    Returns:
        Dictionary mapping slide number to list of texts on that slide
    """
    texts = extract_text_from_pptx(file_path)
    slides = {}

    for slide_num, text in texts:
        if slide_num not in slides:
            slides[slide_num] = []
        slides[slide_num].append(text)

    return slides


def align_by_slide(english_file: str, arabic_file: str) -> List[Dict]:
    """
    Align texts from two parallel PPTX files by slide number.

    Args:
        english_file: Path to English PPTX
        arabic_file: Path to Arabic PPTX

    Returns:
        List of candidate pairs: [{"english": ..., "arabic": ..., "slide": ..., "confidence": ...}]
    """
    english_slides = extract_texts_by_slide(english_file)
    arabic_slides = extract_texts_by_slide(arabic_file)

    candidates = []

    # Find common slide numbers
    common_slides = set(english_slides.keys()) & set(arabic_slides.keys())

    for slide_num in sorted(common_slides):
        en_texts = english_slides[slide_num]
        ar_texts = arabic_slides[slide_num]

        # Simple alignment: match by position if same count
        if len(en_texts) == len(ar_texts):
            for en, ar in zip(en_texts, ar_texts):
                candidates.append({
                    "english": en,
                    "arabic": ar,
                    "slide": slide_num,
                    "alignment_method": "position",
                    "validated": False
                })
        else:
            # Different counts - still try to align, mark as uncertain
            # Pair up what we can, shorter list determines count
            min_len = min(len(en_texts), len(ar_texts))
            for i in range(min_len):
                candidates.append({
                    "english": en_texts[i],
                    "arabic": ar_texts[i],
                    "slide": slide_num,
                    "alignment_method": "position_uncertain",
                    "validated": False
                })

            # Note unpaired texts
            if len(en_texts) > min_len:
                for i in range(min_len, len(en_texts)):
                    candidates.append({
                        "english": en_texts[i],
                        "arabic": "[UNPAIRED]",
                        "slide": slide_num,
                        "alignment_method": "unpaired_english",
                        "validated": False
                    })

            if len(ar_texts) > min_len:
                for i in range(min_len, len(ar_texts)):
                    candidates.append({
                        "english": "[UNPAIRED]",
                        "arabic": ar_texts[i],
                        "slide": slide_num,
                        "alignment_method": "unpaired_arabic",
                        "validated": False
                    })

    return candidates


def validate_pair_with_llm(english: str, arabic: str) -> Tuple[bool, str]:
    """
    Use LLM to validate if an English-Arabic pair are true translations.

    Returns:
        Tuple of (is_valid, reason)
    """
    if not API_URL or not API_KEY or "......" in API_URL or "......" in API_KEY:
        return False, "API not configured"

    if english == "[UNPAIRED]" or arabic == "[UNPAIRED]":
        return False, "Unpaired text"

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {API_KEY}"
    }

    payload = {
        "model": "llama-3.3-70b-versatile",
        "messages": [
            {
                "role": "system",
                "content": """You are a translation validation expert. Given an English text and an Arabic text,
determine if they are valid translations of each other (same meaning, same scope).

Respond in this exact format:
VALID: yes/no
REASON: brief explanation

Be strict - reject pairs that:
- Have different meanings
- Are only partially matching
- Seem to be from different contexts
- One is significantly longer/shorter suggesting missing content"""
            },
            {
                "role": "user",
                "content": f"English: \"{english}\"\nArabic: \"{arabic}\""
            }
        ],
        "temperature": 0.1
    }

    try:
        response = requests.post(API_URL, headers=headers, json=payload, timeout=30)
        response.raise_for_status()

        data = response.json()
        result = data["choices"][0]["message"]["content"].strip()

        # Parse response
        is_valid = "VALID: yes" in result.lower() or "valid:yes" in result.lower().replace(" ", "")
        reason = result.split("REASON:")[-1].strip() if "REASON:" in result else result

        return is_valid, reason

    except Exception as e:
        return False, f"Validation error: {str(e)}"


def validate_candidates(candidates: List[Dict], validate_all: bool = True) -> List[Dict]:
    """
    Validate candidate pairs using LLM.

    Args:
        candidates: List of candidate pairs
        validate_all: If True, validate all pairs. If False, only validate uncertain ones.

    Returns:
        Updated candidates with validation results
    """
    for candidate in candidates:
        # Skip already validated or unpaired
        if candidate.get("validated") and not validate_all:
            continue

        if candidate["english"] == "[UNPAIRED]" or candidate["arabic"] == "[UNPAIRED]":
            candidate["validated"] = False
            candidate["validation_reason"] = "Unpaired text"
            continue

        is_valid, reason = validate_pair_with_llm(candidate["english"], candidate["arabic"])
        candidate["validated"] = is_valid
        candidate["validation_reason"] = reason

    return candidates


def build_dictionary_from_parallel_pptx(
    english_file: str,
    arabic_file: str,
    validate: bool = True
) -> Dict:
    """
    Build dictionary from parallel PowerPoint files.

    Args:
        english_file: Path to English PPTX
        arabic_file: Path to Arabic PPTX
        validate: Whether to validate pairs with LLM

    Returns:
        Dictionary with results: candidates, validated count, added count
    """
    # Step 1: Align texts by slide
    candidates = align_by_slide(english_file, arabic_file)

    # Step 2: Validate with LLM if requested
    if validate:
        candidates = validate_candidates(candidates)

    # Step 3: Add validated pairs to dictionary
    valid_entries = [
        {"english": c["english"], "arabic": c["arabic"], "validated": True}
        for c in candidates
        if c.get("validated", False)
    ]

    added_count = 0
    if valid_entries:
        added_count = add_entries_bulk(valid_entries)

    return {
        "total_candidates": len(candidates),
        "validated_pairs": len(valid_entries),
        "added_to_dictionary": added_count,
        "candidates": candidates
    }
