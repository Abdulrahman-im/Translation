"""Dictionary service for semantic translation lookup."""

import json
import requests
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from datetime import datetime

# Path to dictionary file
DICTIONARY_PATH = Path(__file__).resolve().parent.parent.parent / "data" / "dictionary.json"

# Import API config from translator
from .translator import API_URL, API_KEY


def load_dictionary() -> Dict:
    """Load dictionary from JSON file."""
    if not DICTIONARY_PATH.exists():
        return {"entries": [], "metadata": {"version": "1.0", "last_updated": None, "total_entries": 0}}

    with open(DICTIONARY_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_dictionary(data: Dict) -> None:
    """Save dictionary to JSON file."""
    data["metadata"]["last_updated"] = datetime.now().isoformat()
    data["metadata"]["total_entries"] = len(data["entries"])

    DICTIONARY_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(DICTIONARY_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_all_entries() -> List[Dict]:
    """Get all dictionary entries."""
    data = load_dictionary()
    return data.get("entries", [])


def add_entry(english: str, arabic: str, validated: bool = False) -> bool:
    """Add a new entry to the dictionary."""
    data = load_dictionary()

    # Check for duplicate
    for entry in data["entries"]:
        if entry["english"].lower() == english.lower():
            # Update existing entry
            entry["arabic"] = arabic
            entry["validated"] = validated
            save_dictionary(data)
            return True

    # Add new entry
    data["entries"].append({
        "english": english,
        "arabic": arabic,
        "validated": validated
    })
    save_dictionary(data)
    return True


def add_entries_bulk(entries: List[Dict]) -> int:
    """Add multiple entries to the dictionary. Returns count of added entries."""
    data = load_dictionary()
    existing = {e["english"].lower() for e in data["entries"]}
    added = 0

    for entry in entries:
        if entry["english"].lower() not in existing:
            data["entries"].append(entry)
            existing.add(entry["english"].lower())
            added += 1
        else:
            # Update existing
            for e in data["entries"]:
                if e["english"].lower() == entry["english"].lower():
                    e["arabic"] = entry["arabic"]
                    e["validated"] = entry.get("validated", False)
                    break

    save_dictionary(data)
    return added


def find_exact_match(text: str) -> Optional[str]:
    """Find exact match in dictionary."""
    entries = get_all_entries()
    text_lower = text.strip().lower()

    for entry in entries:
        if entry["english"].lower() == text_lower:
            return entry["arabic"]

    return None


def find_semantic_matches(text: str, top_k: int = 5) -> List[Dict]:
    """
    Find semantically similar entries using LLM.
    Returns top_k most relevant entries from the dictionary.
    """
    entries = get_all_entries()

    if not entries:
        return []

    if not API_URL or not API_KEY or "......" in API_URL or "......" in API_KEY:
        # API not configured, return empty
        return []

    # Build a prompt to find similar entries
    entries_text = "\n".join([f"{i+1}. \"{e['english']}\" -> \"{e['arabic']}\""
                              for i, e in enumerate(entries[:50])])  # Limit to 50 for prompt size

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {API_KEY}"
    }

    payload = {
        "model": "llama-3.3-70b-versatile",
        "messages": [
            {
                "role": "system",
                "content": """You are a translation assistant. Given a text to translate and a list of existing translations,
identify which existing translations are most relevant or semantically similar to help translate the new text.
Return ONLY the numbers of the most relevant entries (up to 5), separated by commas.
If none are relevant, return "NONE".
Example response: 1, 5, 12"""
            },
            {
                "role": "user",
                "content": f"Text to translate: \"{text}\"\n\nExisting translations:\n{entries_text}"
            }
        ],
        "temperature": 0.1
    }

    try:
        response = requests.post(API_URL, headers=headers, json=payload, timeout=30)
        response.raise_for_status()

        data = response.json()
        result = data["choices"][0]["message"]["content"].strip()

        if result.upper() == "NONE":
            return []

        # Parse the numbers
        matches = []
        for num_str in result.split(","):
            try:
                idx = int(num_str.strip()) - 1  # Convert to 0-indexed
                if 0 <= idx < len(entries):
                    matches.append(entries[idx])
            except ValueError:
                continue

        return matches[:top_k]

    except Exception as e:
        print(f"Error finding semantic matches: {e}")
        return []


def build_translation_context(text: str) -> str:
    """
    Build context for translation by finding relevant dictionary entries.
    Returns a formatted string of relevant translations to include in the prompt.
    """
    # First check for exact match
    exact = find_exact_match(text)
    if exact:
        return f"EXACT MATCH FOUND: \"{text}\" -> \"{exact}\""

    # Find semantic matches
    matches = find_semantic_matches(text)

    if not matches:
        return ""

    context = "Use these similar translations as reference:\n"
    for m in matches:
        context += f"- \"{m['english']}\" -> \"{m['arabic']}\"\n"

    return context


def get_dictionary_stats() -> Dict:
    """Get dictionary statistics."""
    data = load_dictionary()
    entries = data.get("entries", [])
    validated = sum(1 for e in entries if e.get("validated", False))

    return {
        "total_entries": len(entries),
        "validated_entries": validated,
        "unvalidated_entries": len(entries) - validated,
        "last_updated": data.get("metadata", {}).get("last_updated")
    }
