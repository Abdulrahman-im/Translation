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
from .pptx_translator import translate_pptx_in_place, translate_pptx_with_options

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
    "build_dictionary_from_parallel_pptx",
    "translate_pptx_in_place",
    "translate_pptx_with_options"
]
