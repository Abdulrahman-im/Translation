"""
Alignment service for building dictionary from parallel PowerPoint files.

Handles imperfectly aligned files where:
- Slides may be several positions apart (not 1:1)
- Sentences within slides may be in different order
- Some content may be missing from one version
"""

import requests
import re
from typing import Dict, List, Tuple, Optional
from .pptx_parser import extract_text_from_pptx, get_slide_count
from .translator import API_URL, API_KEY
from .dictionary import add_entries_bulk


def extract_texts_by_slide(file_path: str) -> Dict[int, List[str]]:
    """Extract texts grouped by slide number."""
    texts = extract_text_from_pptx(file_path)
    slides = {}
    for slide_num, text in texts:
        if slide_num not in slides:
            slides[slide_num] = []
        # Clean and filter text
        text = text.strip()
        if text and len(text) > 1:  # Skip single characters
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


# ============================================================================
# SLIDE FINGERPRINTING - For finding matching slides even when far apart
# ============================================================================

def get_slide_fingerprint(texts: List[str]) -> Dict:
    """
    Create a fingerprint of a slide's content for matching.

    Features:
    - Text count
    - Total character count
    - Has numbers
    - Has bullet-like patterns
    - Word count ranges
    """
    if not texts:
        return {"empty": True}

    total_chars = sum(len(t) for t in texts)
    total_words = sum(len(t.split()) for t in texts)
    has_numbers = any(re.search(r'\d+', t) for t in texts)
    has_bullets = any(re.match(r'^[\-\•\*\d]+\.?\s', t) for t in texts)

    # Get word count distribution (short/medium/long texts)
    short_texts = sum(1 for t in texts if len(t.split()) <= 5)
    medium_texts = sum(1 for t in texts if 5 < len(t.split()) <= 20)
    long_texts = sum(1 for t in texts if len(t.split()) > 20)

    return {
        "empty": False,
        "text_count": len(texts),
        "total_chars": total_chars,
        "total_words": total_words,
        "has_numbers": has_numbers,
        "has_bullets": has_bullets,
        "short_texts": short_texts,
        "medium_texts": medium_texts,
        "long_texts": long_texts
    }


def fingerprint_similarity(fp1: Dict, fp2: Dict) -> float:
    """
    Calculate similarity between two slide fingerprints.
    Returns 0.0 to 1.0
    """
    if fp1.get("empty") or fp2.get("empty"):
        return 0.0

    score = 0.0
    max_score = 0.0

    # Text count similarity (weight: 3)
    max_score += 3
    count_diff = abs(fp1["text_count"] - fp2["text_count"])
    if count_diff == 0:
        score += 3
    elif count_diff <= 2:
        score += 2
    elif count_diff <= 5:
        score += 1

    # Total words similarity (weight: 2)
    max_score += 2
    max_words = max(fp1["total_words"], fp2["total_words"])
    if max_words > 0:
        word_ratio = min(fp1["total_words"], fp2["total_words"]) / max_words
        score += 2 * word_ratio

    # Has numbers match (weight: 1)
    max_score += 1
    if fp1["has_numbers"] == fp2["has_numbers"]:
        score += 1

    # Has bullets match (weight: 1)
    max_score += 1
    if fp1["has_bullets"] == fp2["has_bullets"]:
        score += 1

    # Text length distribution similarity (weight: 2)
    max_score += 2
    dist_match = 0
    for key in ["short_texts", "medium_texts", "long_texts"]:
        if abs(fp1[key] - fp2[key]) <= 1:
            dist_match += 1
    score += (dist_match / 3) * 2

    return score / max_score if max_score > 0 else 0.0


# ============================================================================
# IMPROVED SLIDE MATCHING - Search globally with larger offsets
# ============================================================================

def find_best_slide_matches(
    english_slides: Dict[int, List[str]],
    arabic_slides: Dict[int, List[str]],
    max_offset: int = 10
) -> List[Dict]:
    """
    Find the best matching Arabic slide for each English slide.

    Uses fingerprint similarity + LLM validation for top candidates.
    Allows for large offsets (up to max_offset slides apart).
    """
    # Create fingerprints for all slides
    en_fingerprints = {num: get_slide_fingerprint(texts) for num, texts in english_slides.items()}
    ar_fingerprints = {num: get_slide_fingerprint(texts) for num, texts in arabic_slides.items()}

    en_slide_nums = sorted(english_slides.keys())
    ar_slide_nums = sorted(arabic_slides.keys())

    mappings = []
    used_ar_slides = set()

    for en_num in en_slide_nums:
        en_texts = english_slides[en_num]
        en_fp = en_fingerprints[en_num]

        if en_fp.get("empty"):
            continue

        # Score all potential Arabic slides
        candidates = []
        for ar_num in ar_slide_nums:
            if ar_num in used_ar_slides:
                continue

            ar_fp = ar_fingerprints[ar_num]
            if ar_fp.get("empty"):
                continue

            # Calculate fingerprint similarity
            fp_sim = fingerprint_similarity(en_fp, ar_fp)

            # Position bonus: prefer slides at similar positions
            position_diff = abs(en_num - ar_num)
            position_bonus = max(0, 1 - (position_diff / max_offset)) * 0.2

            combined_score = fp_sim + position_bonus

            if combined_score > 0.3:  # Minimum threshold
                candidates.append({
                    "ar_num": ar_num,
                    "fp_sim": fp_sim,
                    "position_diff": position_diff,
                    "combined_score": combined_score
                })

        # Sort by combined score
        candidates.sort(key=lambda x: x["combined_score"], reverse=True)

        # Take top 3 candidates for LLM validation
        best_match = None
        best_confidence = 0

        for cand in candidates[:3]:
            ar_num = cand["ar_num"]
            ar_texts = arabic_slides[ar_num]

            is_match, confidence, reason = validate_slide_correspondence(
                en_num, en_texts, ar_num, ar_texts
            )

            if is_match and confidence > best_confidence:
                best_match = ar_num
                best_confidence = confidence

        if best_match is not None and best_confidence >= 0.3:
            mappings.append({
                "en_slide": en_num,
                "ar_slide": best_match,
                "confidence": best_confidence,
                "en_texts": english_slides[en_num],
                "ar_texts": arabic_slides[best_match]
            })
            used_ar_slides.add(best_match)
            print(f"  Matched: EN slide {en_num} <-> AR slide {best_match} (confidence: {best_confidence:.2f})")

    return mappings


def validate_slide_correspondence(
    en_slide_num: int,
    en_texts: List[str],
    ar_slide_num: int,
    ar_texts: List[str]
) -> Tuple[bool, float, str]:
    """
    Use LLM to check if two slides are likely corresponding translations.
    """
    en_content = "\n".join([f"- {t}" for t in en_texts[:10]])
    ar_content = "\n".join([f"- {t}" for t in ar_texts[:10]])

    system_prompt = """You are an expert at comparing document slides to determine if they are translations of each other.
Analyze the structure and content of both slides and determine if they are likely parallel translations.

Consider:
1. Similar number of text elements
2. Similar structure/layout patterns
3. Content that appears to be translations (even if you can't fully verify the Arabic)
4. Similar formatting patterns (titles, bullet points, etc.)

IMPORTANT: Be lenient - slides don't need to be perfect matches. Accept if they seem to cover the same topic.

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
        # Fallback to fingerprint-based heuristic
        en_fp = get_slide_fingerprint(en_texts)
        ar_fp = get_slide_fingerprint(ar_texts)
        fp_sim = fingerprint_similarity(en_fp, ar_fp)

        if fp_sim >= 0.6:
            return True, 0.5, "Similar structure (fingerprint fallback)"
        return False, 0.2, "Different structure (fingerprint fallback)"

    is_match = "MATCH: yes" in result.lower() or "match:yes" in result.lower().replace(" ", "")

    confidence = 0.5
    if "CONFIDENCE: high" in result.lower() or "confidence:high" in result.lower().replace(" ", ""):
        confidence = 0.9
    elif "CONFIDENCE: medium" in result.lower() or "confidence:medium" in result.lower().replace(" ", ""):
        confidence = 0.6
    elif "CONFIDENCE: low" in result.lower() or "confidence:low" in result.lower().replace(" ", ""):
        confidence = 0.3

    reason = result.split("REASON:")[-1].strip() if "REASON:" in result else result

    return is_match, confidence, reason


# ============================================================================
# IMPROVED SENTENCE MATCHING - Handle out-of-order sentences
# ============================================================================

def match_sentences_within_slides(
    en_texts: List[str],
    ar_texts: List[str]
) -> List[Dict]:
    """
    Use LLM to match sentences between corresponding slides.
    Handles cases where sentences are in different order or some are missing.
    """
    if not en_texts or not ar_texts:
        return []

    # Build numbered lists
    en_numbered = "\n".join([f"{i+1}. {t}" for i, t in enumerate(en_texts)])
    ar_numbered = "\n".join([f"{i+1}. {t}" for i, t in enumerate(ar_texts)])

    system_prompt = """You are an expert at matching English sentences with their Arabic translations.
Given numbered lists of English and Arabic texts from corresponding slides, identify which English sentences match which Arabic sentences.

IMPORTANT:
- Sentences may be in DIFFERENT ORDER in each list
- Not all sentences may have matches (one language may have extra content)
- Match based on MEANING and CONTEXT, not position
- Be generous - match pairs even if the translation is slightly paraphrased
- Only skip pairs that clearly don't match

Respond with a list of matches in this exact format (one per line):
EN:1 -> AR:2 (confidence: high/medium/low)
EN:3 -> AR:1 (confidence: high/medium/low)

If a sentence has no match, don't include it.
If you cannot determine any matches, respond with: NO_MATCHES"""

    user_prompt = f"""English texts:
{en_numbered}

Arabic texts:
{ar_numbered}

Match the sentences (they may be in different order):"""

    result = call_llm(system_prompt, user_prompt)

    matches = []

    if not result or "NO_MATCHES" in result:
        # Fallback: try to match by structural similarity
        return match_by_structure(en_texts, ar_texts)

    # Parse LLM response
    for line in result.strip().split("\n"):
        line = line.strip()
        if "->" not in line:
            continue

        try:
            parts = line.split("->")
            en_part = parts[0].strip()
            ar_part = parts[1].strip()

            # Extract indices
            en_idx = int(re.search(r'\d+', en_part).group()) - 1
            ar_idx = int(re.search(r'\d+', ar_part.split("(")[0]).group()) - 1

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
        except (ValueError, IndexError, AttributeError):
            continue

    return matches


def match_by_structure(en_texts: List[str], ar_texts: List[str]) -> List[Dict]:
    """
    Fallback matching by structural similarity (length, word count).
    """
    matches = []
    used_ar = set()

    for en_idx, en_text in enumerate(en_texts):
        en_words = len(en_text.split())

        best_ar_idx = None
        best_score = 0

        for ar_idx, ar_text in enumerate(ar_texts):
            if ar_idx in used_ar:
                continue

            ar_words = len(ar_text.split())

            # Arabic text is often slightly shorter than English
            # Accept if word count ratio is between 0.5 and 1.5
            if en_words > 0 and ar_words > 0:
                ratio = ar_words / en_words
                if 0.4 <= ratio <= 1.6:
                    score = 1 - abs(ratio - 0.9) / 0.7  # Optimal ratio around 0.9
                    if score > best_score:
                        best_score = score
                        best_ar_idx = ar_idx

        if best_ar_idx is not None and best_score > 0.3:
            matches.append({
                "english": en_texts[en_idx],
                "arabic": ar_texts[best_ar_idx],
                "en_index": en_idx,
                "ar_index": best_ar_idx,
                "confidence": best_score * 0.4,  # Lower confidence for structural match
                "match_method": "structure_fallback"
            })
            used_ar.add(best_ar_idx)

    return matches


# ============================================================================
# PAIR VALIDATION
# ============================================================================

def validate_pair_with_llm(english: str, arabic: str) -> Tuple[bool, str]:
    """
    Use LLM to validate if an English-Arabic pair are true translations.
    """
    if not API_URL or not API_KEY or "......" in API_URL or "......" in API_KEY:
        return False, "API not configured"

    if not english or not arabic or english == "[UNPAIRED]" or arabic == "[UNPAIRED]":
        return False, "Invalid text"

    system_prompt = """You are a translation validation expert. Given an English text and an Arabic text,
determine if they are valid translations of each other (same meaning, same scope).

Respond in this exact format:
VALID: yes/no
REASON: brief explanation

Be LENIENT - accept pairs that:
- Convey the same general meaning
- May have slight paraphrasing
- May have minor additions/omissions

Reject pairs that:
- Have completely different meanings
- Are clearly from different contexts
- One is a title and the other is body text"""

    user_prompt = f'English: "{english}"\nArabic: "{arabic}"'

    result = call_llm(system_prompt, user_prompt)

    if not result:
        return False, "Validation unavailable"

    is_valid = "VALID: yes" in result.lower() or "valid:yes" in result.lower().replace(" ", "")
    reason = result.split("REASON:")[-1].strip() if "REASON:" in result else result

    return is_valid, reason


def validate_candidates(candidates: List[Dict]) -> List[Dict]:
    """Validate candidate pairs using LLM."""
    for candidate in candidates:
        if candidate.get("validated"):
            continue

        english = candidate.get("english", "")
        arabic = candidate.get("arabic", "")

        if not english or not arabic:
            candidate["validated"] = False
            candidate["validation_reason"] = "Empty text"
            continue

        # High-confidence matches can skip detailed validation
        if candidate.get("confidence", 0) >= 0.7:
            candidate["validated"] = True
            candidate["validation_reason"] = "High confidence match"
            continue

        is_valid, reason = validate_pair_with_llm(english, arabic)
        candidate["validated"] = is_valid
        candidate["validation_reason"] = reason

    return candidates


# ============================================================================
# MAIN ENTRY POINTS
# ============================================================================

def align_with_heuristics(english_file: str, arabic_file: str) -> List[Dict]:
    """
    Align texts from two parallel PPTX files using smart heuristics.
    Handles imperfect alignment with large slide offsets.
    """
    print("\n[Alignment] Extracting texts from files...")
    english_slides = extract_texts_by_slide(english_file)
    arabic_slides = extract_texts_by_slide(arabic_file)

    print(f"[Alignment] Found {len(english_slides)} English slides, {len(arabic_slides)} Arabic slides")

    candidates = []

    # Step 1: Find slide mappings (allows large offsets)
    print("[Alignment] Finding slide matches...")
    slide_mappings = find_best_slide_matches(english_slides, arabic_slides, max_offset=10)
    print(f"[Alignment] Matched {len(slide_mappings)} slide pairs")

    # Step 2: Within each mapped slide pair, match sentences
    print("[Alignment] Matching sentences within slides...")
    for mapping in slide_mappings:
        en_slide = mapping["en_slide"]
        ar_slide = mapping["ar_slide"]
        slide_confidence = mapping["confidence"]
        en_texts = mapping["en_texts"]
        ar_texts = mapping["ar_texts"]

        sentence_matches = match_sentences_within_slides(en_texts, ar_texts)

        for match in sentence_matches:
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

    print(f"[Alignment] Found {len(candidates)} candidate pairs")
    return candidates


def build_dictionary_from_parallel_pptx(
    english_file: str,
    arabic_file: str,
    validate: bool = True,
    use_heuristics: bool = True
) -> Dict:
    """
    Build dictionary from parallel PowerPoint files.
    Handles imperfect alignment where slides may be several positions apart.
    """
    print("\n" + "="*60)
    print("DICTIONARY BUILDER")
    print("="*60)

    # Get slide counts
    en_slide_count = get_slide_count(english_file)
    ar_slide_count = get_slide_count(arabic_file)
    print(f"English file: {en_slide_count} slides")
    print(f"Arabic file: {ar_slide_count} slides")

    # Step 1: Align texts
    if use_heuristics:
        candidates = align_with_heuristics(english_file, arabic_file)
    else:
        candidates = align_by_position(english_file, arabic_file)

    # Step 2: Validate with LLM
    if validate and candidates:
        print("[Validation] Validating candidate pairs...")
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
        print(f"[Dictionary] Added {added_count} entries")

    print("\n" + "="*60)
    print(f"Summary: {len(candidates)} candidates -> {len(valid_entries)} validated -> {added_count} added")
    print("="*60 + "\n")

    return {
        "en_slide_count": en_slide_count,
        "ar_slide_count": ar_slide_count,
        "total_candidates": len(candidates),
        "validated_pairs": len(valid_entries),
        "added_to_dictionary": added_count,
        "candidates": candidates
    }


def align_by_position(english_file: str, arabic_file: str) -> List[Dict]:
    """Simple position-based alignment (fallback method)."""
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
