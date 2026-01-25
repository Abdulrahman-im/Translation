"""PPTX in-place translation and comprehensive RTL conversion service."""

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.shapes.group import GroupShape
from pptx.shapes.base import BaseShape
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from typing import List, Tuple, Optional, Dict
from lxml import etree

from .translator import translate_text
from .pptx_parser import parse_slide_range


# ============================================================================
# RTL TEXT FUNCTIONS
# ============================================================================

def set_paragraph_rtl(paragraph):
    """Set paragraph to RTL with right alignment."""
    try:
        pPr = paragraph._p.get_or_add_pPr()
        pPr.set(qn('a:rtl'), '1')
        paragraph.alignment = PP_ALIGN.RIGHT
    except Exception as e:
        print(f"set_paragraph_rtl error: {e}")


def set_textframe_rtl(text_frame):
    """Set entire text frame to RTL."""
    try:
        for paragraph in text_frame.paragraphs:
            set_paragraph_rtl(paragraph)
    except Exception as e:
        print(f"set_textframe_rtl error: {e}")


def swap_textframe_margins(text_frame):
    """Swap left/right margins in text frame."""
    try:
        txBody = text_frame._txBody
        bodyPr = txBody.find(qn('a:bodyPr'))
        if bodyPr is not None:
            lIns = bodyPr.get('lIns')
            rIns = bodyPr.get('rIns')
            if lIns is not None and rIns is not None:
                bodyPr.set('lIns', rIns)
                bodyPr.set('rIns', lIns)
            elif lIns is not None:
                bodyPr.set('rIns', lIns)
                bodyPr.attrib.pop('lIns', None)
            elif rIns is not None:
                bodyPr.set('lIns', rIns)
                bodyPr.attrib.pop('rIns', None)
    except Exception as e:
        print(f"swap_textframe_margins error: {e}")


def set_table_rtl(shape):
    """Set table to RTL direction."""
    try:
        if not shape.has_table:
            return
        table = shape.table
        tbl = table._tbl

        # Set table RTL
        tblPr = tbl.find(qn('a:tblPr'))
        if tblPr is None:
            tblPr = etree.SubElement(tbl, qn('a:tblPr'))
            tbl.insert(0, tblPr)
        tblPr.set('rtl', '1')

        # Set each cell to RTL
        for row in table.rows:
            for cell in row.cells:
                if cell.text_frame:
                    set_textframe_rtl(cell.text_frame)
    except Exception as e:
        print(f"set_table_rtl error: {e}")


# ============================================================================
# SHAPE MIRRORING - THE CORE FUNCTION
# ============================================================================

def mirror_shape(shape, slide_width):
    """
    Mirror a shape's position horizontally.
    Formula: new_left = slide_width - old_left - shape_width
    """
    try:
        old_left = shape.left
        shape_width = shape.width
        new_left = slide_width - old_left - shape_width

        # Apply the new position
        if new_left >= 0:
            shape.left = int(new_left)
        else:
            # Shape extends beyond slide - place at left edge
            shape.left = 0

        return True
    except Exception as e:
        print(f"mirror_shape error for '{getattr(shape, 'name', 'unknown')}': {e}")
        return False


def flip_shape_horizontal(shape):
    """Apply horizontal flip to a shape (for arrows, etc.)."""
    try:
        # Find the xfrm element
        sp = shape._element
        spPr = sp.find('.//' + qn('p:spPr')) or sp.find('.//' + qn('a:spPr'))
        if spPr is None:
            return

        xfrm = spPr.find(qn('a:xfrm'))
        if xfrm is not None:
            current = xfrm.get('flipH', '0')
            xfrm.set('flipH', '0' if current == '1' else '1')
    except Exception as e:
        print(f"flip_shape_horizontal error: {e}")


# ============================================================================
# SHAPE TYPE DETECTION
# ============================================================================

def is_arrow_or_directional(shape) -> bool:
    """Check if shape is an arrow or directional element."""
    try:
        name = getattr(shape, 'name', '').lower()
        keywords = ['arrow', 'chevron', 'triangle', 'pointer', '>', '<', 'flow']
        return any(kw in name for kw in keywords)
    except:
        return False


def is_thinkcell(shape) -> bool:
    """Check if shape is ThinkCell."""
    try:
        name = getattr(shape, 'name', '').lower()
        if 'thinkcell' in name or 'think-cell' in name:
            return True
        xml = etree.tostring(shape._element, encoding='unicode').lower()
        return 'thinkcell' in xml
    except:
        return False


# ============================================================================
# TRANSLATION
# ============================================================================

def translate_shape_text(shape, translations):
    """Translate text in a shape."""
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            orig = run.text.strip()
            if orig:
                trans = translate_text(orig)
                run.text = trans
                translations.append((orig, trans))


def translate_table_text(shape, translations):
    """Translate text in a table."""
    if not shape.has_table:
        return
    for row in shape.table.rows:
        for cell in row.cells:
            if cell.text_frame:
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        orig = run.text.strip()
                        if orig:
                            trans = translate_text(orig)
                            run.text = trans
                            translations.append((orig, trans))


# ============================================================================
# MAIN SLIDE PROCESSING
# ============================================================================

def process_shape(shape, slide_width, do_mirror, translations):
    """
    Process a single shape:
    1. Translate text
    2. Set RTL direction
    3. Mirror position
    4. Flip directional shapes
    """
    # Skip ThinkCell (but still mirror position)
    if is_thinkcell(shape):
        if do_mirror:
            mirror_shape(shape, slide_width)
        return

    # Handle text shapes
    if shape.has_text_frame:
        translate_shape_text(shape, translations)
        set_textframe_rtl(shape.text_frame)
        swap_textframe_margins(shape.text_frame)

    # Handle tables
    if shape.has_table:
        translate_table_text(shape, translations)
        set_table_rtl(shape)

    # MIRROR POSITION - This is the key for RTL layout
    if do_mirror:
        mirror_shape(shape, slide_width)

    # Flip arrows and directional shapes
    if do_mirror and is_arrow_or_directional(shape):
        flip_shape_horizontal(shape)


def process_group(group, slide_width, do_mirror, translations):
    """
    Process a grouped shape:
    1. Mirror the group's position (children move with it)
    2. Process children for text/RTL only
    """
    # Mirror the entire group
    if do_mirror:
        mirror_shape(group, slide_width)

    # Process children for text and RTL (but don't mirror - they're relative to group)
    for child in group.shapes:
        if isinstance(child, GroupShape):
            # Nested group - recursive but no mirroring
            process_group(child, slide_width, False, translations)
        else:
            # Process text and RTL only
            if child.has_text_frame:
                translate_shape_text(child, translations)
                set_textframe_rtl(child.text_frame)
                swap_textframe_margins(child.text_frame)
            if child.has_table:
                translate_table_text(child, translations)
                set_table_rtl(child)
            # Flip arrows in groups too
            if is_arrow_or_directional(child):
                flip_shape_horizontal(child)


def process_slide_rtl(slide, slide_width, do_mirror=True):
    """
    Process entire slide for RTL conversion.

    IMPORTANT: We mirror ALL shapes except:
    - Shape children inside groups (they're relative to group position)

    This ensures the entire slide layout is flipped.
    """
    translations = []

    # Get all shapes as a list (to avoid modification during iteration)
    shapes = list(slide.shapes)

    for shape in shapes:
        if isinstance(shape, GroupShape):
            process_group(shape, slide_width, do_mirror, translations)
        else:
            process_shape(shape, slide_width, do_mirror, translations)

    return translations


# ============================================================================
# MAIN ENTRY POINTS
# ============================================================================

def translate_pptx_in_place(
    input_path: str,
    output_path: str,
    slide_range: Optional[str] = None,
    mirror_layout: bool = True
) -> Dict:
    """
    Translate PPTX and convert layout to RTL.

    This function:
    1. Translates all text from English to Arabic
    2. Sets text direction to RTL
    3. Right-aligns all text
    4. Mirrors ALL shape positions (flips entire layout)
    5. Flips directional shapes (arrows, chevrons)
    6. Sets tables to RTL direction
    """
    print(f"Opening presentation: {input_path}")
    prs = Presentation(input_path)

    total_slides = len(prs.slides)
    slide_width = prs.slide_width
    print(f"Slide width: {slide_width} EMUs ({slide_width / 914400:.2f} inches)")

    # Parse slide range
    slides_to_process = parse_slide_range(slide_range or "", total_slides)
    print(f"Processing slides: {sorted(slides_to_process)}")

    all_translations = []
    processed_slides = 0

    for slide_num, slide in enumerate(prs.slides, start=1):
        if slide_num not in slides_to_process:
            continue

        processed_slides += 1
        print(f"\n--- Processing slide {slide_num} ---")
        print(f"Number of shapes: {len(slide.shapes)}")

        # Process the slide
        slide_translations = process_slide_rtl(slide, slide_width, do_mirror=mirror_layout)

        print(f"Translated {len(slide_translations)} text items")

        for orig, trans in slide_translations:
            all_translations.append({
                "slide": slide_num,
                "original": orig,
                "translated": trans
            })

    # Save
    print(f"\nSaving to: {output_path}")
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
    """Translate PPTX with flexible output options."""
    from .excel_writer import create_excel_file

    result = {
        "pptx_generated": False,
        "excel_generated": False,
        "pptx_path": None,
        "excel_path": None
    }

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

    if output_excel and excel_output_path and "translations" in result:
        excel_data = [
            (t["slide"], t["original"], t["translated"])
            for t in result["translations"]
        ]
        create_excel_file(excel_data, excel_output_path)
        result["excel_generated"] = True
        result["excel_path"] = excel_output_path

    return result
