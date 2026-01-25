"""PPTX in-place translation and RTL conversion service."""

from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.shapes.group import GroupShape
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from typing import List, Tuple, Optional, Dict
from lxml import etree
from copy import deepcopy

from .translator import translate_text
from .pptx_parser import parse_slide_range


# ============================================================================
# UNGROUP SHAPES (Like VBA code does)
# ============================================================================

def ungroup_all_shapes(slide):
    """
    Ungroup all grouped shapes on a slide, recursively.
    This matches the VBA behavior:

    Do While (groupsExist = True)
      groupsExist = False
      For Each shp In sld.Shapes
          If shp.Type = msoGroup Then
              shp.Ungroup
              groupsExist = True
          End If
      Next shp
    Loop
    """
    max_iterations = 20  # Safety limit
    iteration = 0

    while iteration < max_iterations:
        iteration += 1
        found_group = False

        for shape in list(slide.shapes):
            if isinstance(shape, GroupShape):
                print(f"  Ungrouping: '{getattr(shape, 'name', 'unknown')}'")
                try:
                    ungroup_shape(shape, slide)
                    found_group = True
                    break  # Restart loop after ungrouping
                except Exception as e:
                    print(f"  Could not ungroup: {e}")

        if not found_group:
            break

    print(f"  Ungrouping complete after {iteration} iterations")


def ungroup_shape(group_shape, slide):
    """
    Ungroup a single group shape, adding children directly to the slide.
    Children positions are converted from group-relative to slide-absolute.
    """
    # Get group position
    group_left = group_shape.left
    group_top = group_shape.top

    # Get the shapes tree
    spTree = slide.shapes._spTree

    # Get group element and its children
    grpSp = group_shape._element

    # Find all child shapes in the group
    for child_elem in list(grpSp):
        tag = etree.QName(child_elem.tag).localname

        if tag in ['sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp']:
            # Clone the element
            new_elem = deepcopy(child_elem)

            # Adjust position (convert from group-relative to absolute)
            adjust_element_position(new_elem, group_left, group_top)

            # Add to slide's shape tree
            spTree.append(new_elem)

    # Remove the group from the slide
    grpSp.getparent().remove(grpSp)


def adjust_element_position(elem, offset_x, offset_y):
    """Adjust element position by adding offsets."""
    # Find xfrm in spPr or grpSpPr
    for spPr_tag in ['p:spPr', 'p:grpSpPr', 'a:spPr']:
        spPr = elem.find('.//' + qn(spPr_tag))
        if spPr is not None:
            xfrm = spPr.find(qn('a:xfrm'))
            if xfrm is not None:
                off = xfrm.find(qn('a:off'))
                if off is not None:
                    x = int(off.get('x', 0))
                    y = int(off.get('y', 0))
                    off.set('x', str(x + offset_x))
                    off.set('y', str(y + offset_y))
                    return


# ============================================================================
# RTL TEXT DIRECTION (Bullet on RIGHT of text)
# ============================================================================

def set_paragraph_rtl(paragraph):
    """
    Set paragraph to RTL mode.

    In RTL:
    - Bullet appears on the RIGHT side of text
    - Text flows to the LEFT of the bullet
    - This is controlled by a:rtl='1' on pPr

    We do NOT set explicit alignment - let RTL handle it.
    """
    try:
        pPr = paragraph._p.get_or_add_pPr()

        # This is THE key setting for bullet position
        # When rtl='1', bullet moves to right side of text
        pPr.set(qn('a:rtl'), '1')

        # Remove any explicit alignment - let RTL default handle it
        # In RTL mode, default alignment is RIGHT (start of line)
        if 'algn' in pPr.attrib:
            del pPr.attrib['algn']

    except Exception as e:
        print(f"set_paragraph_rtl error: {e}")


def set_textframe_rtl(text_frame):
    """Set entire text frame to RTL."""
    try:
        for paragraph in text_frame.paragraphs:
            set_paragraph_rtl(paragraph)
    except Exception as e:
        print(f"set_textframe_rtl error: {e}")


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
# SHAPE POSITION MIRRORING
# ============================================================================

def mirror_shape_position(shape, slide_width):
    """
    Mirror shape position: new_left = slide_width - old_left - width
    """
    try:
        old_left = shape.left
        width = shape.width
        new_left = slide_width - old_left - width

        if new_left < 0:
            new_left = 0

        shape_name = getattr(shape, 'name', 'unknown')
        print(f"  Mirror '{shape_name}': {old_left/914400:.2f}\" -> {new_left/914400:.2f}\"")

        shape.left = int(new_left)
        return True
    except Exception as e:
        print(f"  Mirror error: {e}")
        return False


def flip_shape_horizontal(shape):
    """Apply horizontal flip to shape."""
    try:
        sp = shape._element
        spPr = sp.find(qn('p:spPr')) or sp.find(qn('a:spPr'))
        if spPr is None:
            return

        xfrm = spPr.find(qn('a:xfrm'))
        if xfrm is not None:
            current = xfrm.get('flipH', '0')
            xfrm.set('flipH', '0' if current == '1' else '1')
    except Exception as e:
        print(f"flip error: {e}")


# ============================================================================
# SHAPE DETECTION
# ============================================================================

def is_title_shape(shape, slide):
    """Check if shape is the slide title."""
    try:
        if hasattr(slide.shapes, 'title') and slide.shapes.title is not None:
            if shape == slide.shapes.title:
                return True
            # Also check by text content
            if shape.has_text_frame and slide.shapes.title.has_text_frame:
                if shape.text_frame.text == slide.shapes.title.text_frame.text:
                    return True
    except:
        pass
    return False


def is_arrow_shape(shape) -> bool:
    """Check if shape is directional."""
    name = getattr(shape, 'name', '').lower()
    keywords = ['arrow', 'chevron', 'triangle', 'pointer', 'flow']
    return any(kw in name for kw in keywords)


def is_thinkcell(shape) -> bool:
    """Check if shape is ThinkCell."""
    name = getattr(shape, 'name', '').lower()
    return 'thinkcell' in name or 'think-cell' in name


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

def process_slide(slide, slide_width, do_mirror=True):
    """
    Process slide for RTL conversion.

    Following VBA logic:
    1. Ungroup all groups first
    2. For title: only change text direction (no position mirror)
    3. For other shapes with text: mirror position + change text direction
    4. For tables: mirror position + change table direction
    5. For shapes without text: just mirror position
    """
    translations = []

    # Step 1: Ungroup all groups (like VBA does)
    if do_mirror:
        print("  Step 1: Ungrouping shapes...")
        ungroup_all_shapes(slide)

    # Step 2: Get the title text for comparison
    title_text = None
    try:
        if hasattr(slide.shapes, 'title') and slide.shapes.title is not None:
            if slide.shapes.title.has_text_frame:
                title_text = slide.shapes.title.text_frame.text
                print(f"  Title text: '{title_text[:50]}...' " if len(title_text) > 50 else f"  Title text: '{title_text}'")
    except:
        pass

    # Step 3: Process all shapes
    print(f"  Step 2: Processing {len(slide.shapes)} shapes...")

    for shape in list(slide.shapes):
        shape_name = getattr(shape, 'name', 'unknown')

        # Skip groups (should be ungrouped by now, but just in case)
        if isinstance(shape, GroupShape):
            print(f"  Skipping remaining group: '{shape_name}'")
            continue

        # Check if this is the title
        is_title = False
        if title_text and shape.has_text_frame:
            if shape.text_frame.text == title_text:
                is_title = True

        # Handle based on shape type (following VBA logic)
        if shape.has_text_frame:
            # Translate text
            translate_shape_text(shape, translations)

            # Set RTL text direction
            set_textframe_rtl(shape.text_frame)

            if is_title:
                # Title: only change text direction, no position mirror
                print(f"  Title '{shape_name}': RTL text only (no mirror)")
            else:
                # Other text shapes: mirror position
                if do_mirror:
                    mirror_shape_position(shape, slide_width)

        elif shape.has_table:
            # Tables: translate, set RTL, mirror position
            translate_table_text(shape, translations)
            set_table_rtl(shape)
            if do_mirror:
                mirror_shape_position(shape, slide_width)

        else:
            # Shapes without text: just mirror position
            if do_mirror:
                mirror_shape_position(shape, slide_width)

        # Flip directional shapes
        if do_mirror and is_arrow_shape(shape):
            flip_shape_horizontal(shape)

    return translations


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

def translate_pptx_in_place(
    input_path: str,
    output_path: str,
    slide_range: Optional[str] = None,
    mirror_layout: bool = True
) -> Dict:
    """
    Translate PPTX and convert to RTL layout.

    Process:
    1. Ungroup all grouped shapes (like VBA code)
    2. For each shape:
       - If title: change text direction only
       - If text shape: mirror position + change text direction
       - If table: mirror position + change table direction
       - If other: mirror position only
    3. Flip directional shapes (arrows, etc.)
    """
    print(f"\n{'='*60}")
    print("PPTX RTL TRANSLATOR")
    print(f"{'='*60}")
    print(f"Input: {input_path}")
    print(f"Output: {output_path}")
    print(f"Mirror: {mirror_layout}")

    prs = Presentation(input_path)
    total_slides = len(prs.slides)
    slide_width = prs.slide_width

    print(f"\nSlide width: {slide_width/914400:.2f} inches")

    slides_to_process = parse_slide_range(slide_range or "", total_slides)

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
    print("Saving...")
    prs.save(output_path)
    print(f"Done! {len(all_translations)} items translated.")
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
