"""PPTX text extraction service."""

from pptx import Presentation
from pptx.shapes.group import GroupShape
from pptx.shapes.base import BaseShape
from typing import List, Tuple


def extract_text_from_shape(shape: BaseShape) -> List[str]:
    """Extract text from a single shape, handling different shape types."""
    texts = []

    # Handle grouped shapes recursively
    if isinstance(shape, GroupShape):
        for child_shape in shape.shapes:
            texts.extend(extract_text_from_shape(child_shape))
        return texts

    # Handle shapes with text frames
    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                text = run.text.strip()
                if text:
                    texts.append(text)

    # Handle tables
    if shape.has_table:
        table = shape.table
        for row in table.rows:
            for cell in row.cells:
                if cell.text_frame:
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            text = run.text.strip()
                            if text:
                                texts.append(text)

    return texts


def extract_text_from_pptx(file_path: str) -> List[Tuple[int, str]]:
    """
    Extract all translatable text from a PowerPoint file.

    Args:
        file_path: Path to the PPTX file

    Returns:
        List of (slide_number, text) tuples
    """
    prs = Presentation(file_path)
    results = []

    for slide_num, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            texts = extract_text_from_shape(shape)
            for text in texts:
                # Skip empty or whitespace-only text
                if text and text.strip():
                    results.append((slide_num, text))

    return results
