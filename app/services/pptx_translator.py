"""PPTX translation with VBA-style RTL mirroring."""

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
# STEP 1: UNGROUP ALL SHAPES (Exact VBA logic)
# ============================================================================

def ungroup_all_shapes(slide):
    """
    Ungroup all grouped shapes on a slide, recursively.
    Matches VBA:
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
    max_iterations = 50
    iteration = 0

    while iteration < max_iterations:
        iteration += 1
        groups_exist = False

        for shape in list(slide.shapes):
            if isinstance(shape, GroupShape):
                print(f"    Ungrouping: '{getattr(shape, 'name', 'unknown')}'")
                try:
                    ungroup_shape(shape, slide)
                    groups_exist = True
                    break  # Restart loop after ungrouping
                except Exception as e:
                    print(f"    Could not ungroup: {e}")

        if not groups_exist:
            break

    print(f"    Ungrouping complete ({iteration} iterations)")


def ungroup_shape(group_shape, slide):
    """Ungroup a single group shape, adding children directly to the slide."""
    group_left = group_shape.left
    group_top = group_shape.top

    spTree = slide.shapes._spTree
    grpSp = group_shape._element

    for child_elem in list(grpSp):
        tag = etree.QName(child_elem.tag).localname

        if tag in ['sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp']:
            new_elem = deepcopy(child_elem)
            adjust_element_position(new_elem, group_left, group_top)
            spTree.append(new_elem)

    grpSp.getparent().remove(grpSp)


def adjust_element_position(elem, offset_x, offset_y):
    """Adjust element position by adding offsets."""
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
# STEP 2: MIRROR POSITION (VBA: shp.Left = slideWidth - shp.Left - shp.Width)
# ============================================================================

def mirror_shape_position(shape, slide_width):
    """Mirror shape position horizontally."""
    try:
        old_left = shape.left
        width = shape.width
        new_left = slide_width - old_left - width

        if new_left < 0:
            new_left = 0

        shape.left = int(new_left)
        return True
    except Exception as e:
        print(f"    Mirror error: {e}")
        return False


# ============================================================================
# STEP 3: TEXT DIRECTION (VBA: TextFrame.TextRange.ParagraphFormat.TextDirection)
# ============================================================================

def set_text_direction_rtl(text_frame):
    """
    Set text direction to RTL for all paragraphs.
    VBA equivalent: TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionRightToLeft
    """
    try:
        for paragraph in text_frame.paragraphs:
            pPr = paragraph._p.get_or_add_pPr()
            pPr.set(qn('a:rtl'), '1')
    except Exception as e:
        print(f"    Text direction error: {e}")


def set_table_direction_rtl(shape):
    """
    Set table direction to RTL.
    VBA equivalent: shp.Table.TableDirection = ppDirectionRightToLeft
    """
    try:
        if not shape.has_table:
            return

        table = shape.table
        tbl = table._tbl

        # Set table RTL (VBA: shp.Table.TableDirection = ppDirectionRightToLeft)
        tblPr = tbl.find(qn('a:tblPr'))
        if tblPr is None:
            tblPr = etree.SubElement(tbl, qn('a:tblPr'))
            tbl.insert(0, tblPr)
        tblPr.set('rtl', '1')

        # Set each cell text direction (VBA: cell.Shape.TextFrame.TextRange.ParagraphFormat.TextDirection)
        for row in table.rows:
            for cell in row.cells:
                if cell.text_frame:
                    set_text_direction_rtl(cell.text_frame)

    except Exception as e:
        print(f"    Table direction error: {e}")


# ============================================================================
# STEP 4: APPLY VBA MIRRORING LOGIC (BEFORE TRANSLATION)
# ============================================================================

def apply_vba_mirror_logic(slide, slide_width):
    """
    Apply the exact VBA mirroring logic:
    1. Ungroup all groups
    2. Get slide title text
    3. For each shape:
       - If title: only change text direction (no mirror)
       - If has text: mirror + change text direction
       - If table: mirror + change table direction + cell text direction
       - Else: just mirror
    """
    print("  [VBA Mirror] Step 1: Ungrouping all shapes...")
    ungroup_all_shapes(slide)

    # Get slide title text for comparison
    slide_title = None
    try:
        if hasattr(slide.shapes, 'title') and slide.shapes.title is not None:
            if slide.shapes.title.has_text_frame:
                slide_title = slide.shapes.title.text_frame.text
                print(f"  [VBA Mirror] Title: '{slide_title[:50]}...'" if len(str(slide_title)) > 50 else f"  [VBA Mirror] Title: '{slide_title}'")
    except:
        pass

    print(f"  [VBA Mirror] Step 2: Processing {len(slide.shapes)} shapes...")

    for shape in list(slide.shapes):
        shape_name = getattr(shape, 'name', 'unknown')

        # Skip any remaining groups
        if isinstance(shape, GroupShape):
            print(f"    Skipping group: '{shape_name}'")
            continue

        # Check if this shape is the title
        is_title = False
        if slide_title and shape.has_text_frame:
            try:
                if shape.text_frame.text == slide_title:
                    is_title = True
            except:
                pass

        # VBA Logic:
        if shape.has_text_frame:
            if is_title:
                # Title: only change text direction, NO mirror
                print(f"    '{shape_name}' (TITLE): text direction only")
                set_text_direction_rtl(shape.text_frame)
            else:
                # Has text but not title: mirror + text direction
                print(f"    '{shape_name}' (TEXT): mirror + text direction")
                mirror_shape_position(shape, slide_width)
                set_text_direction_rtl(shape.text_frame)

        elif shape.has_table:
            # Table: mirror + table direction + cell text direction
            print(f"    '{shape_name}' (TABLE): mirror + table direction")
            mirror_shape_position(shape, slide_width)
            set_table_direction_rtl(shape)

        else:
            # No text: just mirror
            print(f"    '{shape_name}' (OTHER): mirror only")
            mirror_shape_position(shape, slide_width)


# ============================================================================
# STEP 5: TRANSLATE TEXT (AFTER MIRRORING)
# ============================================================================

def translate_slide_text(slide):
    """Translate all text in the slide AFTER mirroring is done."""
    translations = []

    for shape in list(slide.shapes):
        if isinstance(shape, GroupShape):
            continue

        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    orig = run.text.strip()
                    if orig:
                        trans = translate_text(orig)
                        run.text = trans
                        translations.append((orig, trans))

        elif shape.has_table:
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
    Translate PPTX with VBA-style RTL mirroring.

    Process (matching VBA exactly):
    1. FIRST: Apply VBA mirror logic (ungroup, mirror positions, set text direction)
    2. THEN: Translate all text
    """
    print(f"\n{'='*60}")
    print("PPTX TRANSLATOR (VBA Mirror Logic)")
    print(f"{'='*60}")
    print(f"Input: {input_path}")
    print(f"Output: {output_path}")
    print(f"Mirror: {mirror_layout}")

    prs = Presentation(input_path)
    total_slides = len(prs.slides)
    slide_width = prs.slide_width

    print(f"Slide width: {slide_width/914400:.2f} inches ({slide_width} EMUs)")

    slides_to_process = parse_slide_range(slide_range or "", total_slides)
    print(f"Processing slides: {slides_to_process}")

    all_translations = []
    processed_slides = 0

    for slide_num, slide in enumerate(prs.slides, start=1):
        if slide_num not in slides_to_process:
            continue

        processed_slides += 1
        print(f"\n{'='*40}")
        print(f"SLIDE {slide_num}")
        print(f"{'='*40}")

        # STEP 1: Apply VBA mirror logic FIRST
        if mirror_layout:
            apply_vba_mirror_logic(slide, slide_width)

        # STEP 2: Translate text AFTER mirroring
        print(f"  [Translate] Translating text...")
        slide_translations = translate_slide_text(slide)

        for orig, trans in slide_translations:
            all_translations.append({
                "slide": slide_num,
                "original": orig,
                "translated": trans
            })

        print(f"  [Translate] {len(slide_translations)} items translated")

    print(f"\n{'='*60}")
    print("Saving...")
    prs.save(output_path)
    print(f"Done! {len(all_translations)} total items translated.")
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
