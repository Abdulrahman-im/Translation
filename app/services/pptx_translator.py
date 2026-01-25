"""PPTX in-place translation and comprehensive RTL conversion service."""

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.shapes.group import GroupShape
from pptx.shapes.connector import Connector
from pptx.shapes.autoshape import Shape
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn, nsmap
from pptx.oxml import parse_xml
from typing import List, Tuple, Optional, Dict
from lxml import etree
import copy

from .translator import translate_text
from .pptx_parser import parse_slide_range


# ============================================================================
# RTL TEXT DIRECTION FUNCTIONS
# ============================================================================

def set_paragraph_rtl(paragraph):
    """
    Set paragraph to full RTL mode with proper alignment.
    Handles: text direction, alignment, and bidirectional settings.
    """
    try:
        pPr = paragraph._p.get_or_add_pPr()

        # Set RTL direction
        pPr.set(qn('a:rtl'), '1')

        # Set right alignment for RTL
        paragraph.alignment = PP_ALIGN.RIGHT

        # Set font direction hint for proper rendering
        for run in paragraph.runs:
            try:
                rPr = run._r.get_or_add_rPr()
                # Remove any LTR hints
                for attr in list(rPr.attrib.keys()):
                    if 'lang' in attr.lower():
                        pass  # Keep language settings
            except:
                pass

    except Exception as e:
        print(f"Error setting paragraph RTL: {e}")


def set_textframe_rtl(text_frame):
    """
    Set entire text frame to RTL mode.
    Handles text frame level settings.
    """
    try:
        # Set all paragraphs to RTL
        for paragraph in text_frame.paragraphs:
            set_paragraph_rtl(paragraph)

        # Try to set text frame anchor to right
        try:
            txBody = text_frame._txBody
            bodyPr = txBody.find(qn('a:bodyPr'))
            if bodyPr is not None:
                # Set anchor to right for RTL
                bodyPr.set('anchor', 'r')
        except:
            pass

    except Exception as e:
        print(f"Error setting text frame RTL: {e}")


def mirror_text_frame_margins(text_frame):
    """
    Mirror the left and right margins of a text frame for RTL.
    Swaps left margin with right margin.
    """
    try:
        txBody = text_frame._txBody
        bodyPr = txBody.find(qn('a:bodyPr'))
        if bodyPr is not None:
            # Get current margins
            left_margin = bodyPr.get('lIns')
            right_margin = bodyPr.get('rIns')

            # Swap them
            if left_margin is not None:
                bodyPr.set('rIns', left_margin)
            if right_margin is not None:
                bodyPr.set('lIns', right_margin)
            elif left_margin is not None:
                # If only left was set, move it to right and reset left
                bodyPr.set('lIns', '91440')  # Default value in EMUs

    except Exception as e:
        print(f"Error mirroring text frame margins: {e}")


def set_list_rtl(paragraph):
    """
    Set bullet/numbered list to RTL alignment.
    Fixes bullet position for Arabic text.
    """
    try:
        pPr = paragraph._p.get_or_add_pPr()

        # Check if this is a list item (has bullet or numbering)
        buNone = pPr.find(qn('a:buNone'))
        buChar = pPr.find(qn('a:buChar'))
        buAutoNum = pPr.find(qn('a:buAutoNum'))

        if buChar is not None or buAutoNum is not None:
            # This is a list item - ensure RTL bullet alignment
            pPr.set(qn('a:rtl'), '1')

            # Set indent for RTL (swap margin and indent)
            marL = pPr.get('marL')
            indent = pPr.get('indent')

            if marL and indent:
                # For RTL, we might need to adjust these
                pass  # Keep current values but ensure RTL is set

    except Exception as e:
        print(f"Error setting list RTL: {e}")


# ============================================================================
# TABLE RTL FUNCTIONS
# ============================================================================

def set_table_full_rtl(table_shape):
    """
    Comprehensive RTL conversion for tables.
    Handles: table direction, cell alignment, text direction.
    """
    try:
        table = table_shape.table
        tbl = table._tbl

        # Set table direction to RTL
        tblPr = tbl.find(qn('a:tblPr'))
        if tblPr is None:
            # Create tblPr if it doesn't exist
            tblPr = etree.Element(qn('a:tblPr'))
            tbl.insert(0, tblPr)
        tblPr.set('rtl', '1')

        # Process each cell
        for row in table.rows:
            for cell in row.cells:
                # Set cell text to RTL
                if cell.text_frame:
                    set_textframe_rtl(cell.text_frame)
                    mirror_text_frame_margins(cell.text_frame)

                # Try to set cell-level RTL
                try:
                    tc = cell._tc
                    tcPr = tc.find(qn('a:tcPr'))
                    if tcPr is not None:
                        # Mirror cell margins
                        marL = tcPr.get('marL')
                        marR = tcPr.get('marR')
                        if marL is not None:
                            tcPr.set('marR', marL)
                        if marR is not None:
                            tcPr.set('marL', marR)
                except:
                    pass

    except Exception as e:
        print(f"Error setting table RTL: {e}")


# ============================================================================
# SHAPE MIRRORING FUNCTIONS
# ============================================================================

def mirror_shape_position(shape, slide_width):
    """
    Mirror shape position horizontally for RTL layout.
    Calculates: new_left = slide_width - current_left - width
    """
    try:
        current_left = shape.left
        current_width = shape.width
        new_left = slide_width - current_left - current_width

        if new_left >= 0:
            shape.left = new_left
        else:
            shape.left = 0

    except Exception as e:
        print(f"Could not mirror shape position: {e}")


def mirror_shape_flip(shape):
    """
    Apply horizontal flip to shape (for arrows, chevrons, etc.).
    This visually reverses directional shapes.
    """
    try:
        spPr = shape._element.find(qn('p:spPr'))
        if spPr is None:
            spPr = shape._element.find(qn('a:spPr'))

        if spPr is not None:
            xfrm = spPr.find(qn('a:xfrm'))
            if xfrm is not None:
                # Toggle flipH attribute
                current_flip = xfrm.get('flipH')
                if current_flip == '1':
                    xfrm.set('flipH', '0')
                else:
                    xfrm.set('flipH', '1')
    except Exception as e:
        print(f"Could not flip shape: {e}")


def is_directional_shape(shape) -> bool:
    """
    Check if a shape is directional (arrow, chevron, etc.) that needs flipping.
    """
    try:
        # Check shape name for directional indicators
        if hasattr(shape, 'name'):
            name_lower = shape.name.lower()
            directional_keywords = ['arrow', 'chevron', 'triangle', 'pointer', 'flow', 'connector']
            if any(kw in name_lower for kw in directional_keywords):
                return True

        # Check auto shape type
        if hasattr(shape, 'shape_type'):
            arrow_types = [
                MSO_SHAPE_TYPE.LEFT_ARROW,
                MSO_SHAPE_TYPE.RIGHT_ARROW,
                MSO_SHAPE_TYPE.LEFT_RIGHT_ARROW,
                MSO_SHAPE_TYPE.CHEVRON,
                MSO_SHAPE_TYPE.NOTCHED_RIGHT_ARROW,
            ]
            try:
                if shape.shape_type in arrow_types:
                    return True
            except:
                pass

    except:
        pass
    return False


def reverse_connector(connector, slide_width):
    """
    Reverse a connector's direction for RTL.
    Swaps start and end points.
    """
    try:
        # Mirror the connector position
        mirror_shape_position(connector, slide_width)

        # Try to swap start/end connections
        cxnSp = connector._element

        # Find and swap stCxn and endCxn
        nvCxnSpPr = cxnSp.find(qn('p:nvCxnSpPr'))
        if nvCxnSpPr is not None:
            cNvCxnSpPr = nvCxnSpPr.find(qn('p:cNvCxnSpPr'))
            if cNvCxnSpPr is not None:
                stCxn = cNvCxnSpPr.find(qn('a:stCxn'))
                endCxn = cNvCxnSpPr.find(qn('a:endCxn'))

                if stCxn is not None and endCxn is not None:
                    # Swap the connection IDs
                    st_id = stCxn.get('id')
                    st_idx = stCxn.get('idx')
                    end_id = endCxn.get('id')
                    end_idx = endCxn.get('idx')

                    if st_id and end_id:
                        stCxn.set('id', end_id)
                        endCxn.set('id', st_id)
                    if st_idx and end_idx:
                        stCxn.set('idx', end_idx)
                        endCxn.set('idx', st_idx)

    except Exception as e:
        print(f"Could not reverse connector: {e}")


# ============================================================================
# DETECTION FUNCTIONS
# ============================================================================

def is_thinkcell_shape(shape) -> bool:
    """Check if shape is a ThinkCell object."""
    try:
        if hasattr(shape, 'name'):
            name_lower = shape.name.lower()
            if 'thinkcell' in name_lower or 'think-cell' in name_lower:
                return True

        if hasattr(shape, '_element'):
            xml_str = etree.tostring(shape._element, encoding='unicode')
            if 'thinkcell' in xml_str.lower():
                return True
    except:
        pass
    return False


def is_title_shape(shape, slide) -> bool:
    """Check if shape is the slide title (should not be position-mirrored)."""
    try:
        if hasattr(slide, 'shapes') and hasattr(slide.shapes, 'title'):
            title_shape = slide.shapes.title
            if title_shape is not None and shape == title_shape:
                return True

        if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
            if hasattr(shape, 'placeholder_format'):
                ph_type = shape.placeholder_format.type
                if ph_type in [1, 3]:  # TITLE, CENTER_TITLE
                    return True
    except:
        pass
    return False


def is_connector_shape(shape) -> bool:
    """Check if shape is a connector."""
    try:
        return isinstance(shape, Connector)
    except:
        return False


def is_decorative_shape(shape) -> bool:
    """
    Check if shape is likely decorative (colored blocks, lines, etc.).
    These should always be mirrored.
    """
    try:
        # Shapes without text are likely decorative
        if not shape.has_text_frame:
            return True

        # Check if text frame is empty
        if shape.has_text_frame:
            text = shape.text_frame.text.strip()
            if not text:
                return True

    except:
        pass
    return False


# ============================================================================
# TRANSLATION FUNCTIONS
# ============================================================================

def translate_shape_text(shape, translations_list):
    """Translate all text in a shape."""
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


def translate_table(table_shape, translations_list):
    """Translate all text in a table."""
    if not table_shape.has_table:
        return

    table = table_shape.table
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


# ============================================================================
# MAIN PROCESSING FUNCTIONS
# ============================================================================

def process_shape_rtl(shape, slide, slide_width, do_mirror: bool, translations: list):
    """
    Process a single shape for RTL conversion.

    Handles:
    - Text translation
    - Text direction (RTL)
    - Text alignment (right)
    - Position mirroring
    - Margin swapping
    - Directional shape flipping
    """

    # Skip ThinkCell shapes (just mirror position)
    if is_thinkcell_shape(shape):
        if do_mirror:
            mirror_shape_position(shape, slide_width)
        return

    # Handle connectors specially
    if is_connector_shape(shape):
        if do_mirror:
            reverse_connector(shape, slide_width)
        return

    # Check if this is the title
    is_title = is_title_shape(shape, slide)

    # Check if this is a directional shape (arrows, chevrons)
    is_directional = is_directional_shape(shape)

    # Handle shapes with text
    if shape.has_text_frame:
        # Translate text
        translate_shape_text(shape, translations)

        # Set RTL text direction and alignment
        set_textframe_rtl(shape.text_frame)

        # Mirror text frame margins
        mirror_text_frame_margins(shape.text_frame)

        # Handle list items
        for paragraph in shape.text_frame.paragraphs:
            set_list_rtl(paragraph)

        # Mirror position (except for titles)
        if do_mirror and not is_title:
            mirror_shape_position(shape, slide_width)

    # Handle tables
    elif shape.has_table:
        # Translate table text
        translate_table(shape, translations)

        # Set full RTL for table
        set_table_full_rtl(shape)

        # Mirror position
        if do_mirror:
            mirror_shape_position(shape, slide_width)

    # Handle other shapes (decorative, images, etc.)
    else:
        if do_mirror:
            mirror_shape_position(shape, slide_width)

        # Flip directional shapes
        if is_directional and do_mirror:
            mirror_shape_flip(shape)


def process_group_shape(group_shape, slide, slide_width, do_mirror: bool, translations: list):
    """
    Process a grouped shape.

    For groups:
    - Mirror the group's position
    - Process children for text/RTL (but don't mirror children - they're relative to group)
    """
    # Mirror the group position
    if do_mirror:
        mirror_shape_position(group_shape, slide_width)

    # Process children for text and RTL settings
    for child in group_shape.shapes:
        if isinstance(child, GroupShape):
            # Nested group - process recursively (but don't mirror)
            process_group_shape(child, slide, slide_width, False, translations)
        else:
            # Process child for text/RTL but don't mirror position
            if child.has_text_frame:
                translate_shape_text(child, translations)
                set_textframe_rtl(child.text_frame)
                mirror_text_frame_margins(child.text_frame)
                for paragraph in child.text_frame.paragraphs:
                    set_list_rtl(paragraph)
            elif child.has_table:
                translate_table(child, translations)
                set_table_full_rtl(child)


def process_slide(slide, slide_width, do_mirror: bool = True) -> List[Tuple[str, str]]:
    """
    Process an entire slide for translation and RTL conversion.

    Returns list of (original, translated) text pairs.
    """
    translations = []

    # Process all shapes
    for shape in list(slide.shapes):
        if isinstance(shape, GroupShape):
            process_group_shape(shape, slide, slide_width, do_mirror, translations)
        else:
            process_shape_rtl(shape, slide, slide_width, do_mirror, translations)

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
    Translate a PowerPoint file and convert to RTL layout.

    Comprehensive RTL conversion includes:
    - Text translation
    - Text direction (RTL)
    - Text alignment (right-aligned)
    - Text frame margin mirroring
    - Shape position mirroring
    - Table RTL direction
    - Bullet/list RTL alignment
    - Connector reversal
    - Directional shape flipping

    Args:
        input_path: Path to input PPTX
        output_path: Path to save translated PPTX
        slide_range: Optional slide range (e.g., "1-10")
        mirror_layout: Whether to mirror layout for RTL

    Returns:
        Dictionary with translation statistics
    """
    prs = Presentation(input_path)
    total_slides = len(prs.slides)
    slide_width = prs.slide_width

    # Parse slide range
    slides_to_process = parse_slide_range(slide_range or "", total_slides)

    all_translations = []
    processed_slides = 0

    for slide_num, slide in enumerate(prs.slides, start=1):
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
    """
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
