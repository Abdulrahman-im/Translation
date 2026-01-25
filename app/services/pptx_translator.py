"""PPTX in-place translation and RTL conversion service."""

from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.shapes.group import GroupShape
from pptx.shapes.placeholder import PlaceholderPicture, PlaceholderGraphicFrame
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml.ns import qn
from pptx.oxml import register_element_cls
from typing import List, Tuple, Optional, Dict
from lxml import etree
from copy import deepcopy

from .translator import translate_text
from .pptx_parser import parse_slide_range


# ============================================================================
# BULLET / LIST RTL FUNCTIONS
# ============================================================================

def set_paragraph_full_rtl(paragraph):
    """
    Set paragraph to full RTL mode including bullet alignment.

    In RTL mode:
    - Text flows right-to-left
    - Bullets appear on the RIGHT side of text
    - Margins are interpreted from the RIGHT
    """
    try:
        pPr = paragraph._p.get_or_add_pPr()

        # Set RTL direction - this is the key for bullet position
        pPr.set(qn('a:rtl'), '1')

        # Set right alignment
        pPr.set('algn', 'r')

        # Handle margins for RTL
        # In RTL, marL becomes the RIGHT margin visually
        # We need to swap marL and marR
        marL = pPr.get('marL')
        marR = pPr.get('marR')

        if marL is not None and marR is None:
            # Move left margin to right margin position
            pPr.set('marR', marL)
            # Don't remove marL, just set it to 0
            pPr.set('marL', '0')
        elif marL is not None and marR is not None:
            # Swap them
            pPr.set('marL', marR)
            pPr.set('marR', marL)

        # Handle indent (for bullet offset)
        indent = pPr.get('indent')
        if indent is not None:
            # Indent in RTL should work from right side
            # Keep indent as-is, RTL flag will handle positioning
            pass

    except Exception as e:
        print(f"set_paragraph_full_rtl error: {e}")


def set_textframe_full_rtl(text_frame):
    """Set entire text frame to RTL with proper bullet handling."""
    try:
        # Set RTL on all paragraphs
        for paragraph in text_frame.paragraphs:
            set_paragraph_full_rtl(paragraph)

        # Also set text frame level properties
        try:
            txBody = text_frame._txBody
            bodyPr = txBody.find(qn('a:bodyPr'))
            if bodyPr is not None:
                # Swap left/right insets
                lIns = bodyPr.get('lIns')
                rIns = bodyPr.get('rIns')
                if lIns is not None and rIns is not None:
                    bodyPr.set('lIns', rIns)
                    bodyPr.set('rIns', lIns)
                elif lIns is not None:
                    bodyPr.set('rIns', lIns)
                    bodyPr.set('lIns', '91440')  # Default
        except:
            pass

    except Exception as e:
        print(f"set_textframe_full_rtl error: {e}")


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
                    set_textframe_full_rtl(cell.text_frame)
    except Exception as e:
        print(f"set_table_rtl error: {e}")


# ============================================================================
# SHAPE POSITION MIRRORING
# ============================================================================

def is_placeholder(shape) -> bool:
    """Check if shape is a placeholder."""
    try:
        return hasattr(shape, 'is_placeholder') and shape.is_placeholder
    except:
        return False


def get_shape_position(shape):
    """Get shape position safely."""
    try:
        return shape.left, shape.top, shape.width, shape.height
    except:
        return None, None, None, None


def set_shape_position(shape, left, top=None, width=None, height=None):
    """Set shape position, handling placeholders specially."""
    try:
        # For placeholders, we need to modify the XML directly
        if is_placeholder(shape):
            # Get the spPr element
            sp = shape._element
            spPr = sp.find(qn('p:spPr'))
            if spPr is None:
                spPr = etree.SubElement(sp, qn('p:spPr'))

            # Get or create xfrm element
            xfrm = spPr.find(qn('a:xfrm'))
            if xfrm is None:
                xfrm = etree.SubElement(spPr, qn('a:xfrm'))

            # Get or create off (offset) element
            off = xfrm.find(qn('a:off'))
            if off is None:
                off = etree.SubElement(xfrm, qn('a:off'))

            # Set the position
            off.set('x', str(int(left)))
            if top is not None:
                off.set('y', str(int(top)))

            # Also set via shape properties as backup
            shape.left = int(left)
        else:
            # Regular shapes - just set directly
            shape.left = int(left)

        return True
    except Exception as e:
        print(f"set_shape_position error: {e}")
        return False


def mirror_shape_position(shape, slide_width):
    """
    Mirror shape position horizontally.
    new_left = slide_width - old_left - shape_width
    """
    try:
        old_left, top, width, height = get_shape_position(shape)
        if old_left is None or width is None:
            print(f"  Could not get position for shape: {getattr(shape, 'name', 'unknown')}")
            return False

        new_left = slide_width - old_left - width

        if new_left < 0:
            new_left = 0

        print(f"  Mirroring '{getattr(shape, 'name', 'unknown')}': {old_left/914400:.2f}in -> {new_left/914400:.2f}in")

        return set_shape_position(shape, new_left, top)

    except Exception as e:
        print(f"mirror_shape_position error: {e}")
        return False


def flip_shape_horizontal(shape):
    """Apply horizontal flip to directional shapes."""
    try:
        sp = shape._element
        spPr = sp.find(qn('p:spPr'))
        if spPr is None:
            spPr = sp.find(qn('a:spPr'))
        if spPr is None:
            return

        xfrm = spPr.find(qn('a:xfrm'))
        if xfrm is not None:
            current = xfrm.get('flipH', '0')
            xfrm.set('flipH', '0' if current == '1' else '1')
    except Exception as e:
        print(f"flip_shape_horizontal error: {e}")


# ============================================================================
# SHAPE DETECTION
# ============================================================================

def is_arrow_shape(shape) -> bool:
    """Check if shape is directional."""
    try:
        name = getattr(shape, 'name', '').lower()
        keywords = ['arrow', 'chevron', 'triangle', 'pointer', 'flow', 'connector']
        return any(kw in name for kw in keywords)
    except:
        return False


def is_thinkcell(shape) -> bool:
    """Check if shape is ThinkCell."""
    try:
        name = getattr(shape, 'name', '').lower()
        return 'thinkcell' in name or 'think-cell' in name
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
# MAIN PROCESSING
# ============================================================================

def process_shape(shape, slide_width, do_mirror, translations):
    """Process a single shape for RTL conversion."""

    shape_name = getattr(shape, 'name', 'unknown')
    is_ph = is_placeholder(shape)

    print(f"  Processing: '{shape_name}' (placeholder={is_ph})")

    # Handle ThinkCell
    if is_thinkcell(shape):
        if do_mirror:
            mirror_shape_position(shape, slide_width)
        return

    # Translate and set RTL for text
    if shape.has_text_frame:
        translate_shape_text(shape, translations)
        set_textframe_full_rtl(shape.text_frame)

    # Handle tables
    if shape.has_table:
        translate_table_text(shape, translations)
        set_table_rtl(shape)

    # Mirror position
    if do_mirror:
        mirror_shape_position(shape, slide_width)

    # Flip directional shapes
    if do_mirror and is_arrow_shape(shape):
        flip_shape_horizontal(shape)


def process_group(group, slide_width, do_mirror, translations):
    """Process a grouped shape."""

    print(f"  Processing group: '{getattr(group, 'name', 'unknown')}'")

    # Mirror the entire group
    if do_mirror:
        mirror_shape_position(group, slide_width)

    # Process children for text/RTL only (don't mirror - relative to group)
    for child in group.shapes:
        if isinstance(child, GroupShape):
            process_group(child, slide_width, False, translations)
        else:
            if child.has_text_frame:
                translate_shape_text(child, translations)
                set_textframe_full_rtl(child.text_frame)
            if child.has_table:
                translate_table_text(child, translations)
                set_table_rtl(child)
            if is_arrow_shape(child):
                flip_shape_horizontal(child)


def process_slide(slide, slide_width, do_mirror=True):
    """Process entire slide for RTL conversion."""
    translations = []
    shapes = list(slide.shapes)

    print(f"  Total shapes: {len(shapes)}")

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
    Translate PPTX and convert to RTL layout.

    This handles:
    1. Text translation (English -> Arabic)
    2. RTL text direction with proper bullet alignment
    3. Shape position mirroring (full layout flip)
    4. Table RTL direction
    5. Directional shape flipping (arrows, etc.)
    """
    print(f"\n{'='*60}")
    print(f"PPTX RTL Translator")
    print(f"{'='*60}")
    print(f"Input: {input_path}")
    print(f"Output: {output_path}")
    print(f"Mirror layout: {mirror_layout}")

    prs = Presentation(input_path)
    total_slides = len(prs.slides)
    slide_width = prs.slide_width

    print(f"\nPresentation info:")
    print(f"  Slides: {total_slides}")
    print(f"  Width: {slide_width} EMUs ({slide_width/914400:.2f} inches)")

    slides_to_process = parse_slide_range(slide_range or "", total_slides)
    print(f"  Processing: {sorted(slides_to_process)}")

    all_translations = []
    processed_slides = 0

    for slide_num, slide in enumerate(prs.slides, start=1):
        if slide_num not in slides_to_process:
            continue

        processed_slides += 1
        print(f"\n{'='*40}")
        print(f"SLIDE {slide_num}")
        print(f"{'='*40}")

        slide_translations = process_slide(slide, slide_width, do_mirror=mirror_layout)

        for orig, trans in slide_translations:
            all_translations.append({
                "slide": slide_num,
                "original": orig,
                "translated": trans
            })

    print(f"\n{'='*60}")
    print(f"Saving presentation...")
    prs.save(output_path)
    print(f"Done! Translated {len(all_translations)} text items.")
    print(f"{'='*60}\n")

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
