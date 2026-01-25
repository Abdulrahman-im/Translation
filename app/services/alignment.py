"""Alignment service for building dictionary from parallel PowerPoint files."""

import requests
from typing import Dict, List, Tuple, Optional
from .pptx_parser import extract_text_from_pptx, get_slide_count
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


def call_llm(system_prompt: str, user_prompt: str) -> Optional[str]:
    """Helper function to call the LLM API."""
    if not API_URL or not API_KEY or "......" in API_URL or "......" in API_KEY:
        return None

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {API_KEY}"
    }

    payload = {
        "model": "llama-3.3-70b-versatile",
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.1
    }

    try:
        response = requests.post(API_URL, headers=headers, json=payload, timeout=60)
        response.raise_for_status()
        data = response.json()
        return data["choices"][0]["message"]["content"].strip()
    except Exception as e:
        print(f"LLM API error: {e}")
        return None


def validate_slide_correspondence(
    en_slide_num: int,
    en_texts: List[str],
    ar_slide_num: int,
    ar_texts: List[str]
) -> Tuple[bool, float, str]:
    """
    Use LLM to check if two slides are likely corresponding translations.

    Returns:
        Tuple of (is_match, confidence_score, reason)
    """
    en_content = "\n".join([f"- {t}" for t in en_texts[:10]])  # Limit to first 10 items
    ar_content = "\n".join([f"- {t}" for t in ar_texts[:10]])

    system_prompt = """You are an expert at comparing document slides to determine if they are translations of each other.
Analyze the structure and content of both slides and determine if they are likely parallel translations.

Consider:
1. Similar number of text elements
2. Similar structure/layout patterns
3. Content that appears to be translations (even if you can't fully verify the Arabic)
4. Similar formatting patterns (titles, bullet points, etc.)

Respond in this exact format:
MATCH: yes/no
CONFIDENCE: high/medium/low
REASON: brief explanation"""

    user_prompt = f"""English Slide {en_slide_num} content:
{en_content}

Arabic Slide {ar_slide_num} content:
{ar_content}

Are these slides likely translations of each other?"""

    result = call_llm(system_prompt, user_prompt)

    if not result:
        # Fallback to basic heuristic: similar text count
        ratio = min(len(en_texts), len(ar_texts)) / max(len(en_texts), len(ar_texts)) if max(len(en_texts), len(ar_texts)) > 0 else 0
        if ratio >= 0.7:
            return True, 0.5, "Similar text count (fallback heuristic)"
        return False, 0.2, "Different text counts (fallback heuristic)"

    is_match = "MATCH: yes" in result.lower() or "match:yes" in result.lower().replace(" ", "")

    confidence = 0.5
    if "CONFIDENCE: high" in result.lower():
        confidence = 0.9
    elif "CONFIDENCE: medium" in result.lower():
        confidence = 0.6
    elif "CONFIDENCE: low" in result.lower():
        confidence = 0.3

    reason = result.split("REASON:")[-1].strip() if "REASON:" in result else result

    return is_match, confidence, reason


def match_sentences_within_slides(
    en_texts: List[str],
    ar_texts: List[str]
) -> List[Dict]:
    """
    Use LLM to match sentences between corresponding slides.

    Returns:
        List of matched pairs with confidence scores
    """
    if not en_texts or not ar_texts:
        return []

    # Build numbered lists
    en_numbered = "\n".join([f"{i+1}. {t}" for i, t in enumerate(en_texts)])
    ar_numbered = "\n".join([f"{i+1}. {t}" for i, t in enumerate(ar_texts)])

    system_prompt = """You are an expert at matching English sentences with their Arabic translations.
Given numbered lists of English and Arabic texts from corresponding slides, identify which English sentences match which Arabic sentences.

IMPORTANT:
- Not all sentences may have matches (slides may have extra content in one language)
- Match based on meaning and context, not just position
- Only match pairs you are confident about

Respond with a list of matches in this exact format (one per line):
EN:1 -> AR:2 (confidence: high/medium/low)
EN:3 -> AR:1 (confidence: high/medium/low)

If a sentence has no match, don't include it.
If you cannot determine any matches, respond with: NO_MATCHES"""

    user_prompt = f"""English texts:
{en_numbered}

Arabic texts:
{ar_numbered}

Match the sentences:"""

    result = call_llm(system_prompt, user_prompt)

    matches = []

    if not result or "NO_MATCHES" in result:
        # Fallback: if same length, try position-based matching
        if len(en_texts) == len(ar_texts):
            for i, (en, ar) in enumerate(zip(en_texts, ar_texts)):
                matches.append({
                    "english": en,
                    "arabic": ar,
                    "en_index": i,
                    "ar_index": i,
                    "confidence": 0.4,
                    "match_method": "position_fallback"
                })
        return matches

    # Parse LLM response
    for line in result.strip().split("\n"):
        line = line.strip()
        if "->" not in line:
            continue

        try:
            # Parse format: EN:1 -> AR:2 (confidence: high)
            parts = line.split("->")
            en_part = parts[0].strip()
            ar_part = parts[1].strip()

            # Extract indices
            en_idx = int(en_part.replace("EN:", "").strip()) - 1
            ar_idx_str = ar_part.split("(")[0].replace("AR:", "").strip()
            ar_idx = int(ar_idx_str) - 1

            # Extract confidence
            confidence = 0.5
            if "high" in line.lower():
                confidence = 0.9
            elif "medium" in line.lower():
                confidence = 0.6
            elif "low" in line.lower():
                confidence = 0.3

            if 0 <= en_idx < len(en_texts) and 0 <= ar_idx < len(ar_texts):
                matches.append({
                    "english": en_texts[en_idx],
                    "arabic": ar_texts[ar_idx],
                    "en_index": en_idx,
                    "ar_index": ar_idx,
                    "confidence": confidence,
                    "match_method": "llm_semantic"
                })
        except (ValueError, IndexError):
            continue

    return matches


def find_slide_mappings(
    english_slides: Dict[int, List[str]],
    arabic_slides: Dict[int, List[str]]
) -> List[Dict]:
    """
    Find which English slides correspond to which Arabic slides.
    Slides may not be 1-to-1 due to additions/reorderings.

    Returns:
        List of slide mappings with confidence scores
    """
    mappings = []
    used_ar_slides = set()

    en_slide_nums = sorted(english_slides.keys())
    ar_slide_nums = sorted(arabic_slides.keys())

    for en_num in en_slide_nums:
        en_texts = english_slides[en_num]
        best_match = None
        best_confidence = 0

        # First, check the corresponding position (most likely match)
        candidates = []
        if en_num in ar_slide_nums and en_num not in used_ar_slides:
            candidates.append(en_num)

        # Also check nearby slides (+/- 2)
        for offset in [-1, 1, -2, 2]:
            nearby = en_num + offset
            if nearby in ar_slide_nums and nearby not in used_ar_slides and nearby not in candidates:
                candidates.append(nearby)

        for ar_num in candidates:
            if ar_num in used_ar_slides:
                continue

            ar_texts = arabic_slides[ar_num]
            is_match, confidence, reason = validate_slide_correspondence(
                en_num, en_texts, ar_num, ar_texts
            )

            if is_match and confidence > best_confidence:
                best_match = ar_num
                best_confidence = confidence

        if best_match is not None and best_confidence >= 0.4:
            mappings.append({
                "en_slide": en_num,
                "ar_slide": best_match,
                "confidence": best_confidence,
                "en_texts": en_texts,
                "ar_texts": arabic_slides[best_match]
            })
            used_ar_slides.add(best_match)

    return mappings


def align_with_heuristics(english_file: str, arabic_file: str) -> List[Dict]:
    """
    Align texts from two parallel PPTX files using smart heuristics.

    Heuristics:
    1. Slide-level matching: Check if slides correspond (not just by position)
    2. Sentence-level matching: Within matched slides, match sentences semantically

    Returns:
        List of candidate pairs with confidence scores
    """
    english_slides = extract_texts_by_slide(english_file)
    arabic_slides = extract_texts_by_slide(arabic_file)

    candidates = []

    # Step 1: Find slide mappings
    slide_mappings = find_slide_mappings(english_slides, arabic_slides)

    # Step 2: Within each mapped slide pair, match sentences
    for mapping in slide_mappings:
        en_slide = mapping["en_slide"]
        ar_slide = mapping["ar_slide"]
        slide_confidence = mapping["confidence"]
        en_texts = mapping["en_texts"]
        ar_texts = mapping["ar_texts"]

        # Match sentences within the slide
        sentence_matches = match_sentences_within_slides(en_texts, ar_texts)

        for match in sentence_matches:
            # Combined confidence: slide match * sentence match
            combined_confidence = slide_confidence * match["confidence"]

            candidates.append({
                "english": match["english"],
                "arabic": match["arabic"],
                "en_slide": en_slide,
                "ar_slide": ar_slide,
                "confidence": combined_confidence,
                "alignment_method": match["match_method"],
                "slide_match_confidence": slide_confidence,
                "sentence_match_confidence": match["confidence"],
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

    system_prompt = """You are a translation validation expert. Given an English text and an Arabic text,
determine if they are valid translations of each other (same meaning, same scope).

Respond in this exact format:
VALID: yes/no
REASON: brief explanation

Be strict - reject pairs that:
- Have different meanings
- Are only partially matching
- Seem to be from different contexts
- One is significantly longer/shorter suggesting missing content"""

    user_prompt = f"English: \"{english}\"\nArabic: \"{arabic}\""

    result = call_llm(system_prompt, user_prompt)

    if not result:
        return False, "Validation unavailable"

    is_valid = "VALID: yes" in result.lower() or "valid:yes" in result.lower().replace(" ", "")
    reason = result.split("REASON:")[-1].strip() if "REASON:" in result else result

    return is_valid, reason


def validate_candidates(candidates: List[Dict], validate_all: bool = True) -> List[Dict]:
    """
    Validate candidate pairs using LLM.
    """
    for candidate in candidates:
        if candidate.get("validated") and not validate_all:
            continue

        if candidate.get("english") == "[UNPAIRED]" or candidate.get("arabic") == "[UNPAIRED]":
            candidate["validated"] = False
            candidate["validation_reason"] = "Unpaired text"
            continue

        # High-confidence matches (>0.7) can skip detailed validation
        if candidate.get("confidence", 0) >= 0.7:
            candidate["validated"] = True
            candidate["validation_reason"] = "High confidence match"
            continue

        is_valid, reason = validate_pair_with_llm(candidate["english"], candidate["arabic"])
        candidate["validated"] = is_valid
        candidate["validation_reason"] = reason

    return candidates


def build_dictionary_from_parallel_pptx(
    english_file: str,
    arabic_file: str,
    validate: bool = True,
    use_heuristics: bool = True
) -> Dict:
    """
    Build dictionary from parallel PowerPoint files.

    Args:
        english_file: Path to English PPTX
        arabic_file: Path to Arabic PPTX
        validate: Whether to validate pairs with LLM
        use_heuristics: Whether to use smart heuristics for alignment

    Returns:
        Dictionary with results: candidates, validated count, added count
    """
    # Get slide counts for info
    en_slide_count = get_slide_count(english_file)
    ar_slide_count = get_slide_count(arabic_file)

    # Step 1: Align texts using heuristics or simple position-based
    if use_heuristics:
        candidates = align_with_heuristics(english_file, arabic_file)
    else:
        # Fallback to simple position-based alignment
        candidates = align_by_position(english_file, arabic_file)

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
        "en_slide_count": en_slide_count,
        "ar_slide_count": ar_slide_count,
        "total_candidates": len(candidates),
        "validated_pairs": len(valid_entries),
        "added_to_dictionary": added_count,
        "candidates": candidates
    }


def align_by_position(english_file: str, arabic_file: str) -> List[Dict]:
    """
    Simple position-based alignment (original method).
    """
    english_slides = extract_texts_by_slide(english_file)
    arabic_slides = extract_texts_by_slide(arabic_file)

    candidates = []
    common_slides = set(english_slides.keys()) & set(arabic_slides.keys())

    for slide_num in sorted(common_slides):
        en_texts = english_slides[slide_num]
        ar_texts = arabic_slides[slide_num]

        if len(en_texts) == len(ar_texts):
            for en, ar in zip(en_texts, ar_texts):
                candidates.append({
                    "english": en,
                    "arabic": ar,
                    "en_slide": slide_num,
                    "ar_slide": slide_num,
                    "confidence": 0.5,
                    "alignment_method": "position",
                    "validated": False
                })
        else:
            min_len = min(len(en_texts), len(ar_texts))
            for i in range(min_len):
                candidates.append({
                    "english": en_texts[i],
                    "arabic": ar_texts[i],
                    "en_slide": slide_num,
                    "ar_slide": slide_num,
                    "confidence": 0.3,
                    "alignment_method": "position_uncertain",
                    "validated": False
                })

    return candidates
