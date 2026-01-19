from .pptx_parser import extract_text_from_pptx
from .translator import translate_text
from .excel_writer import create_excel_file
from .dictionary import (
    load_dictionary,
    save_dictionary,
    get_all_entries,
    add_entry,
    find_exact_match,
    find_semantic_matches,
    get_dictionary_stats
)
from .alignment import build_dictionary_from_parallel_pptx

__all__ = [
    "extract_text_from_pptx",
    "translate_text",
    "create_excel_file",
    "load_dictionary",
    "save_dictionary",
    "get_all_entries",
    "add_entry",
    "find_exact_match",
    "find_semantic_matches",
    "get_dictionary_stats",
    "build_dictionary_from_parallel_pptx"
]
