"""PPTX in-place translation and RTL mirroring service."""

from pptx import Presentation
from pptx.util import Emu
from pptx.shapes.group import GroupShape
from pptx.shapes.base import BaseShape
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
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


def set_table_direction_rtl(table):
    """Set table direction to RTL."""
    try:
        tbl = table._tbl
        tblPr = tbl.find(qn('a:tblPr'))
        if tblPr is None:
            tblPr = etree.SubElement(tbl, qn('a:tblPr'))
        tblPr.set('rtl', '1')
    except Exception as e:
        print(f"Could not set table RTL: {e}")


def mirror_shape_position(shape, slide_width):
    """Mirror shape position horizontally (flip from LTR to RTL layout)."""
    try:
        current_left = shape.left
        current_width = shape.width
        # Calculate new left position: slideWidth - left - width
        new_left = slide_width - current_left - current_width

        # Ensure we don't set negative values
        if new_left >= 0:
            shape.left = new_left
        else:
            # If new position would be negative, place at edge
            shape.left = 0

    except Exception as e:
        print(f"Could not mirror shape '{getattr(shape, 'name', 'unknown')}': {e}")


def is_thinkcell_shape(shape) -> bool:
    """Check if a shape is a ThinkCell object."""
    try:
        if hasattr(shape, 'name'):
            name_lower = shape.name.lower()
            if 'thinkcell' in name_lower or 'think-cell' in name_lower:
                return True

        if hasattr(shape, '_element'):
            xml_str = etree.tostring(shape._element, encoding='unicode')
            if 'thinkcell' in xml_str.lower() or 'think-cell' in xml_str.lower():
                return True
    except:
        pass
    return False


def get_all_shapes_flat(slide) -> List:
    """
    Get all shapes from a slide, ungrouping groups recursively.
    Returns a flat list of all shapes.
    """
    all_shapes = []

    def collect_shapes(shapes):
        for shape in shapes:
            if isinstance(shape, GroupShape):
                # Add the group itself (we'll mirror it)
                all_shapes.append(('group', shape))
                # Also collect children for text processing
                collect_shapes(shape.shapes)
            else:
                all_shapes.append(('shape', shape))

    collect_shapes(slide.shapes)
    return all_shapes


def is_title_shape(shape, slide) -> bool:
    """Check if a shape is the slide title."""
    try:
        # Check if this is the title placeholder
        if hasattr(slide, 'shapes') and hasattr(slide.shapes, 'title'):
            title_shape = slide.shapes.title
            if title_shape is not None and shape == title_shape:
                return True

        # Also check by placeholder type
        if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
            if hasattr(shape, 'placeholder_format'):
                ph_type = shape.placeholder_format.type
                # Title placeholder types
                if ph_type in [1, 3]:  # TITLE = 1, CENTER_TITLE = 3
                    return True
    except:
        pass
    return False


def translate_shape_text(shape, translations_list):
    """Translate all text in a shape and set RTL direction."""
    if not shape.has_text_frame:
        return

    text_frame = shape.text_frame
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            original_text = run.text.strip()
            if original_text:
                translated_text = translate_text(original_text)
                run.text = translated_text
                translations_list.append((original_text, translated_text))

        # Set paragraph to RTL
        set_paragraph_rtl(paragraph)


def translate_table_text(shape, translations_list):
    """Translate all text in a table and set RTL direction."""
    if not shape.has_table:
        return

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
                            translations_list.append((original_text, translated_text))

                    # Set paragraph to RTL
                    set_paragraph_rtl(paragraph)


def process_slide(slide, slide_width, do_mirror: bool = True) -> List[Tuple[str, str]]:
    """
    Process a single slide: translate text and mirror layout.

    Based on VBA logic:
    1. Title shape: only change text direction, don't mirror position
    2. Other text shapes: mirror position AND change text direction
    3. Tables: mirror position, change table direction, change cell text directions
    4. Other shapes: just mirror position

    Returns list of (original, translated) tuples.
    """
    translations = []

    # Process all shapes on the slide
    for shape in list(slide.shapes):
        # Skip ThinkCell shapes
        if is_thinkcell_shape(shape):
            if do_mirror:
                mirror_shape_position(shape, slide_width)
            continue

        # Handle grouped shapes
        if isinstance(shape, GroupShape):
            # Mirror the group position
            if do_mirror:
                mirror_shape_position(shape, slide_width)

            # Process text in children (but don't mirror children - they're relative to group)
            for child in shape.shapes:
                translate_shape_text(child, translations)
                if child.has_table:
                    translate_table_text(child, translations)
            continue

        # Check if this is the title
        is_title = is_title_shape(shape, slide)

        # Handle shapes with text frames
        if shape.has_text_frame:
            # Translate text and set RTL
            translate_shape_text(shape, translations)

            # Mirror position ONLY if not title
            if do_mirror and not is_title:
                mirror_shape_position(shape, slide_width)

        # Handle tables
        elif shape.has_table:
            # Translate table and set RTL
            translate_table_text(shape, translations)

            # Mirror position
            if do_mirror:
                mirror_shape_position(shape, slide_width)

        # Handle other shapes (images, shapes without text, etc.)
        else:
            if do_mirror:
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
    slide_width = prs.slide_width

    # Parse slide range
    slides_to_process = parse_slide_range(slide_range or "", total_slides)

    all_translations = []
    processed_slides = 0

    for slide_num, slide in enumerate(prs.slides, start=1):
        # Skip slides not in the requested range
        if slide_num not in slides_to_process:
            continue

        processed_slides += 1

        # Process the slide
        slide_translations = process_slide(slide, slide_width, do_mirror=mirror_layout)

        for orig, trans in slide_translations:
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
