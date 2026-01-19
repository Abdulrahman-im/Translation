"""Translation service with caching using LLM API and semantic dictionary."""

import requests
from typing import Dict, Optional

# =============================================================================
# API CONFIGURATION - Edit these values to match your API
# =============================================================================
API_URL = "https://......"  # <-- Replace with your API URL (--location)
API_KEY = "......"          # <-- Replace with your API key (--header Authorization)
# =============================================================================

# In-memory cache for translations
_translation_cache: Dict[str, str] = {}


def _get_dictionary_context(text: str) -> tuple[Optional[str], str]:
    """
    Get translation context from dictionary.

    Returns:
        Tuple of (exact_match_or_none, context_string)
    """
    try:
        from .dictionary import find_exact_match, find_semantic_matches

        # Check for exact match first
        exact = find_exact_match(text)
        if exact:
            return exact, ""

        # Find semantic matches for context
        matches = find_semantic_matches(text, top_k=5)

        if matches:
            context = "Use these similar translations as reference for style and terminology:\n"
            for m in matches:
                context += f"- \"{m['english']}\" -> \"{m['arabic']}\"\n"
            return None, context

        return None, ""
    except ImportError:
        return None, ""


def call_translation_api(text: str, context: str = "") -> str:
    """
    Call the LLM API to translate English text to Arabic.

    Args:
        text: English text to translate
        context: Optional context with similar translations

    Returns:
        Arabic translation
    """
    if not API_URL or not API_KEY or "......" in API_URL or "......" in API_KEY:
        # Fallback to mock if API not configured
        return f"[AR] {text}"

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {API_KEY}"
    }

    # Build system prompt with context if available
    system_prompt = "You are a professional translator. Translate the following English text to Arabic. Return ONLY the Arabic translation, nothing else. Do not include any explanations or notes."

    if context:
        system_prompt += f"\n\n{context}"

    payload = {
        "model": "llama-3.3-70b-versatile",
        "messages": [
            {
                "role": "system",
                "content": system_prompt
            },
            {
                "role": "user",
                "content": text
            }
        ],
        "temperature": 0.3
    }

    try:
        response = requests.post(API_URL, headers=headers, json=payload, timeout=30)
        response.raise_for_status()

        data = response.json()
        translation = data["choices"][0]["message"]["content"].strip()
        return translation

    except requests.exceptions.RequestException as e:
        print(f"Translation API error: {e}")
        # Fallback to mock on error
        return f"[AR] {text}"
    except (KeyError, IndexError) as e:
        print(f"Error parsing API response: {e}")
        return f"[AR] {text}"


def translate_text(text: str) -> str:
    """
    Translate English text to Arabic using semantic dictionary.

    Translation flow:
    1. Check dictionary for exact match
    2. Check cache
    3. Find semantically similar entries from dictionary
    4. Call LLM API with context
    5. Store in cache
    6. Return translation

    Args:
        text: English text to translate

    Returns:
        Arabic translation
    """
    # Normalize text for lookup
    normalized = text.strip()

    # Check cache first (fastest)
    if normalized in _translation_cache:
        return _translation_cache[normalized]

    # Check dictionary for exact match and get context
    exact_match, context = _get_dictionary_context(normalized)

    if exact_match:
        _translation_cache[normalized] = exact_match
        return exact_match

    # Call translation API with semantic context
    translation = call_translation_api(normalized, context)

    # Store in cache
    _translation_cache[normalized] = translation

    return translation


def clear_cache() -> None:
    """Clear the translation cache."""
    global _translation_cache
    _translation_cache = {}


def get_cache_stats() -> Dict[str, int]:
    """Get cache statistics."""
    try:
        from .dictionary import get_dictionary_stats
        dict_stats = get_dictionary_stats()
    except ImportError:
        dict_stats = {"total_entries": 0}

    return {
        "dictionary_entries": dict_stats.get("total_entries", 0),
        "cached_translations": len(_translation_cache),
    }
