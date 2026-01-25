"""PPTX in-place translation and RTL mirroring service."""

import copy
from pptx import Presentation
from pptx.util import Emu
from pptx.shapes.group import GroupShape
from pptx.shapes.base import BaseShape
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from typing import List, Tuple, Optional, Set, Dict
from lxml import etree

from .translator import translate_text
from .pptx_parser import parse_slide_range


def set_paragraph_rtl(paragraph):
    """Set paragraph to RTL direction."""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set(qn('a:rtl'), '1')

    # Also set alignment to right for RTL
    if paragraph.alignment is None or paragraph.alignment == PP_ALIGN.LEFT:
        paragraph.alignment = PP_ALIGN.RIGHT


def set_paragraph_ltr(paragraph):
    """Set paragraph to LTR direction."""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set(qn('a:rtl'), '0')


def get_paragraph_direction(paragraph) -> str:
    """Get paragraph text direction (rtl or ltr)."""
    pPr = paragraph._p.find(qn('a:pPr'))
    if pPr is not None:
        rtl = pPr.get(qn('a:rtl'))
        if rtl == '1':
            return 'rtl'
    return 'ltr'


def set_table_direction_rtl(table):
    """Set table direction to RTL."""
    tbl = table._tbl
    tblPr = tbl.find(qn('a:tblPr'))
    if tblPr is None:
        tblPr = etree.SubElement(tbl, qn('a:tblPr'))
    tblPr.set('rtl', '1')


def mirror_shape_position(shape, slide_width):
    """Mirror shape position horizontally (flip from LTR to RTL layout)."""
    try:
        # Calculate new left position: slideWidth - left - width
        new_left = slide_width - shape.left - shape.width
        shape.left = new_left
    except Exception as e:
        print(f"Could not mirror shape: {e}")


def is_thinkcell_shape(shape) -> bool:
    """Check if a shape is a ThinkCell object."""
    try:
        # ThinkCell shapes typically have specific naming patterns or are OLE objects
        if hasattr(shape, 'name'):
            name_lower = shape.name.lower()
            if 'thinkcell' in name_lower or 'think-cell' in name_lower:
                return True

        # Check for OLE objects which ThinkCell often uses
        if hasattr(shape, 'ole_format'):
            return True

        # Check shape XML for ThinkCell markers
        if hasattr(shape, '_element'):
            xml_str = etree.tostring(shape._element, encoding='unicode')
            if 'thinkcell' in xml_str.lower() or 'think-cell' in xml_str.lower():
                return True
    except:
        pass
    return False


def translate_and_mirror_shape(shape, slide_width, do_mirror: bool = True):
    """
    Translate text in a shape and mirror its position.

    Args:
        shape: The PowerPoint shape to process
        slide_width: Width of the slide for mirroring calculations
        do_mirror: Whether to mirror the position (True for LTR->RTL conversion)

    Returns:
        List of (original_text, translated_text) tuples for logging
    """
    translations = []

    # Skip ThinkCell shapes for now (they have special handling)
    if is_thinkcell_shape(shape):
        # Still mirror position if requested
        if do_mirror:
            mirror_shape_position(shape, slide_width)
        return translations

    # Handle grouped shapes recursively
    if isinstance(shape, GroupShape):
        # For groups, we need to process children but be careful with positioning
        # The group itself should be mirrored, not individual children
        if do_mirror:
            mirror_shape_position(shape, slide_width)

        for child_shape in shape.shapes:
            child_translations = translate_and_mirror_shape(
                child_shape, slide_width, do_mirror=False  # Don't mirror children, group is mirrored
            )
            translations.extend(child_translations)
        return translations

    # Handle shapes with text frames
    if shape.has_text_frame:
        text_frame = shape.text_frame

        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                original_text = run.text.strip()
                if original_text:
                    # Translate the text
                    translated_text = translate_text(original_text)
                    run.text = translated_text
                    translations.append((original_text, translated_text))

            # Set paragraph to RTL
            set_paragraph_rtl(paragraph)

        # Mirror the shape position
        if do_mirror:
            mirror_shape_position(shape, slide_width)

    # Handle tables
    if shape.has_table:
        table = shape.table

        # Set table direction to RTL
        set_table_direction_rtl(table)

        # Process each cell
        for row in table.rows:
            for cell in row.cells:
                if cell.text_frame:
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            original_text = run.text.strip()
                            if original_text:
                                translated_text = translate_text(original_text)
                                run.text = translated_text
                                translations.append((original_text, translated_text))

                        # Set paragraph to RTL
                        set_paragraph_rtl(paragraph)

        # Mirror the table position
        if do_mirror:
            mirror_shape_position(shape, slide_width)

    # For shapes without text, just mirror position
    elif do_mirror and not shape.has_text_frame and not shape.has_table:
        mirror_shape_position(shape, slide_width)

    return translations


def translate_pptx_in_place(
    input_path: str,
    output_path: str,
    slide_range: Optional[str] = None,
    mirror_layout: bool = True
) -> Dict:
    """
    Translate a PowerPoint file in-place and optionally mirror for RTL.

    Args:
        input_path: Path to the input PPTX file
        output_path: Path to save the translated PPTX file
        slide_range: Optional slide range to process (e.g., "1-10", "1,3,5")
        mirror_layout: Whether to mirror the layout for RTL (default True)

    Returns:
        Dictionary with translation results and statistics
    """
    prs = Presentation(input_path)
    total_slides = len(prs.slides)

    # Parse slide range
    slides_to_process = parse_slide_range(slide_range or "", total_slides)

    all_translations = []
    processed_slides = 0

    for slide_num, slide in enumerate(prs.slides, start=1):
        # Skip slides not in the requested range
        if slide_num not in slides_to_process:
            continue

        processed_slides += 1
        slide_width = prs.slide_width

        # Process all shapes on the slide
        for shape in slide.shapes:
            shape_translations = translate_and_mirror_shape(
                shape, slide_width, do_mirror=mirror_layout
            )
            for orig, trans in shape_translations:
                all_translations.append({
                    "slide": slide_num,
                    "original": orig,
                    "translated": trans
                })

    # Save the translated presentation
    prs.save(output_path)

    return {
        "total_slides": total_slides,
        "processed_slides": processed_slides,
        "total_translations": len(all_translations),
        "translations": all_translations
    }


def translate_pptx_with_options(
    input_path: str,
    pptx_output_path: Optional[str] = None,
    excel_output_path: Optional[str] = None,
    slide_range: Optional[str] = None,
    mirror_layout: bool = True,
    output_excel: bool = False
) -> Dict:
    """
    Translate a PowerPoint file with flexible output options.

    Args:
        input_path: Path to the input PPTX file
        pptx_output_path: Path to save the translated PPTX (if desired)
        excel_output_path: Path to save the Excel file (if desired)
        slide_range: Optional slide range to process
        mirror_layout: Whether to mirror the layout for RTL
        output_excel: Whether to also output an Excel file

    Returns:
        Dictionary with results and file paths
    """
    from .excel_writer import create_excel_file

    result = {
        "pptx_generated": False,
        "excel_generated": False,
        "pptx_path": None,
        "excel_path": None
    }

    # Generate translated PPTX
    if pptx_output_path:
        translation_result = translate_pptx_in_place(
            input_path, pptx_output_path, slide_range, mirror_layout
        )
        result["pptx_generated"] = True
        result["pptx_path"] = pptx_output_path
        result["total_slides"] = translation_result["total_slides"]
        result["processed_slides"] = translation_result["processed_slides"]
        result["total_translations"] = translation_result["total_translations"]
        result["translations"] = translation_result["translations"]

    # Generate Excel file if requested
    if output_excel and excel_output_path and "translations" in result:
        # Convert to the format expected by create_excel_file
        excel_data = [
            (t["slide"], t["original"], t["translated"])
            for t in result["translations"]
        ]
        create_excel_file(excel_data, excel_output_path)
        result["excel_generated"] = True
        result["excel_path"] = excel_output_path

    return result
