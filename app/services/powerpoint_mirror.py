"""
RTL mirroring using PowerPoint directly (same effect as VBA macro).

- Windows: PowerPoint COM (pywin32) – full control
- Mac: AppleScript – drives Microsoft PowerPoint for Mac

Both use PowerPoint's native object model; no python-pptx for layout.
"""

import os
import platform
import shutil
import subprocess

# PowerPoint COM constants (same as VBA)
MsoShapeType = type("MsoShapeType", (), {"msoGroup": 6, "msoTable": 19})()
MsoTextDirection = type("MsoTextDirection", (), {"msoTextDirectionLeftToRight": 1, "msoTextDirectionRightToLeft": 2})()
PpDirection = type("PpDirection", (), {"ppDirectionLeftToRight": 1, "ppDirectionRightToLeft": 2})()


def check_powerpoint_available():
    """Check if we can use PowerPoint (Windows COM or Mac AppleScript)."""
    if platform.system() == "Windows":
        try:
            import win32com.client
            app = win32com.client.Dispatch("PowerPoint.Application")
            app.Quit()
            return True
        except Exception:
            return False
    if platform.system() == "Darwin":
        return os.path.exists("/Applications/Microsoft PowerPoint.app")
    return False


def _mirror_slide_via_com(sld, slide_width):
    """
    Mirror one slide: ungroup, then mirror positions and toggle text/table direction.
    Mirrors the VBA logic of MirrorTextAndTablesBasedOnAlignment.
    """
    # Ungroup all groups (repeat until no groups left)
    groups_exist = True
    while groups_exist:
        groups_exist = False
        for i in range(sld.Shapes.Count, 0, -1):
            shp = sld.Shapes(i)
            if shp.Type == MsoShapeType.msoGroup:
                shp.Ungroup()
                groups_exist = True
                break

    # Get slide title text (from title placeholder)
    slide_title = ""
    try:
        title_shape = sld.Shapes.Title
        if title_shape.HasTextFrame:
            slide_title = title_shape.TextFrame.TextRange.Text or ""
    except Exception:
        pass

    # Process each shape
    for i in range(1, sld.Shapes.Count + 1):
        shp = sld.Shapes(i)
        try:
            if shp.HasTextFrame:
                shp_text = (shp.TextFrame.TextRange.Text or "").strip()
                if slide_title and shp_text == slide_title:
                    # Title: only toggle text direction, do not move
                    try:
                        pdir = shp.TextFrame.TextRange.ParagraphFormat.TextDirection
                        if pdir == MsoTextDirection.msoTextDirectionLeftToRight:
                            shp.TextFrame.TextRange.ParagraphFormat.TextDirection = MsoTextDirection.msoTextDirectionRightToLeft
                        elif pdir == MsoTextDirection.msoTextDirectionRightToLeft:
                            shp.TextFrame.TextRange.ParagraphFormat.TextDirection = MsoTextDirection.msoTextDirectionLeftToRight
                    except Exception:
                        pass
                else:
                    # Text shape (not title): mirror position + toggle direction
                    shp.LockAspectRatio = -1  # msoTrue
                    shp.Left = slide_width - shp.Left - shp.Width
                    try:
                        pdir = shp.TextFrame.TextRange.ParagraphFormat.TextDirection
                        if pdir == MsoTextDirection.msoTextDirectionLeftToRight:
                            shp.TextFrame.TextRange.ParagraphFormat.TextDirection = MsoTextDirection.msoTextDirectionRightToLeft
                        elif pdir == MsoTextDirection.msoTextDirectionRightToLeft:
                            shp.TextFrame.TextRange.ParagraphFormat.TextDirection = MsoTextDirection.msoTextDirectionLeftToRight
                    except Exception:
                        pass
            elif shp.Type == MsoShapeType.msoTable:
                # Table: mirror position, table direction, and each cell's text direction
                shp.LockAspectRatio = -1  # msoTrue
                shp.Left = slide_width - shp.Left - shp.Width
                try:
                    tbl = shp.Table
                    if tbl.TableDirection == PpDirection.ppDirectionLeftToRight:
                        tbl.TableDirection = PpDirection.ppDirectionRightToLeft
                    elif tbl.TableDirection == PpDirection.ppDirectionRightToLeft:
                        tbl.TableDirection = PpDirection.ppDirectionLeftToRight
                except Exception:
                    pass
                try:
                    for r in range(1, tbl.Rows.Count + 1):
                        for c in range(1, tbl.Columns.Count + 1):
                            cell = tbl.Cell(r, c)
                            if cell.Shape.HasTextFrame:
                                try:
                                    pdir = cell.Shape.TextFrame.TextRange.ParagraphFormat.TextDirection
                                    if pdir == MsoTextDirection.msoTextDirectionLeftToRight:
                                        cell.Shape.TextFrame.TextRange.ParagraphFormat.TextDirection = MsoTextDirection.msoTextDirectionRightToLeft
                                    elif pdir == MsoTextDirection.msoTextDirectionRightToLeft:
                                        cell.Shape.TextFrame.TextRange.ParagraphFormat.TextDirection = MsoTextDirection.msoTextDirectionLeftToRight
                                except Exception:
                                    pass
                except Exception:
                    pass
            else:
                # Other shape (no text frame): mirror position only
                shp.LockAspectRatio = -1  # msoTrue
                shp.Left = slide_width - shp.Left - shp.Width
        except Exception:
            continue


def _mirror_with_applescript(abs_path: str, slide_numbers: list = None, timeout: int = 600) -> None:
    """Mirror via AppleScript (Mac). Opens PowerPoint, ungroups, mirrors positions, toggles text/table direction."""
    applescript = '''
tell application "Microsoft PowerPoint"
    activate
    open (POSIX file "%s")
    delay 2

    set thePres to active presentation
    set slideW to width of page setup of thePres
    set totalSlides to count of slides of thePres

    repeat with sldIdx from 1 to totalSlides
        set sld to slide sldIdx of thePres

        -- UNGROUP ALL GROUPS
        set groupsExist to true
        set maxIter to 50
        set iter to 0
        repeat while groupsExist and iter < maxIter
            set iter to iter + 1
            set groupsExist to false
            try
                set shpCount to count of shapes of sld
                repeat with i from 1 to shpCount
                    try
                        set shp to shape i of sld
                        set shpType to shape type of shp
                        if shpType = 6 then
                            ungroup shp
                            set groupsExist to true
                            exit repeat
                        end if
                    end try
                end repeat
            end try
        end repeat

        -- GET SLIDE TITLE TEXT (placeholder type 1 = title)
        set slideTitle to ""
        try
            repeat with i from 1 to (count of shapes of sld)
                set shp to shape i of sld
                try
                    set pType to placeholder type of shp
                    if pType = 1 then
                        if has text frame of shp then
                            set slideTitle to content of text range of text frame of shp
                            exit repeat
                        end if
                    end if
                end try
            end repeat
        end try

        -- PROCESS ALL SHAPES
        set shpCount to count of shapes of sld
        set msoTable to 19
        repeat with i from 1 to shpCount
            try
                set shp to shape i of sld
                set shpType to shape type of shp

                set shpText to ""
                set hasText to false
                try
                    if has text frame of shp then
                        set hasText to true
                        set shpText to content of text range of text frame of shp
                    end if
                end try

                set isTitle to false
                if hasText and slideTitle is not "" then
                    if shpText = slideTitle then
                        set isTitle to true
                    end if
                end if

                if shpType = msoTable then
                    try
                        set oldLeft to left position of shp
                        set shpW to width of shp
                        set newLeft to slideW - oldLeft - shpW
                        if newLeft < 0 then set newLeft to 0
                        set left position of shp to newLeft
                    end try
                    try
                        set tblDir to table direction of table of shp
                        if tblDir = 1 then
                            set table direction of table of shp to 2
                        else if tblDir = 2 then
                            set table direction of table of shp to 1
                        end if
                    end try
                    try
                        repeat with rowIdx from 1 to (count of rows of table of shp)
                            set theRow to row rowIdx of table of shp
                            repeat with cellIdx from 1 to (count of cells of theRow)
                                set theCell to cell cellIdx of theRow
                                if has text frame of shape of theCell then
                                    set pDir to text direction of paragraph format of text range of text frame of shape of theCell
                                    if pDir = 1 then
                                        set text direction of paragraph format of text range of text frame of shape of theCell to 2
                                    else if pDir = 2 then
                                        set text direction of paragraph format of text range of text frame of shape of theCell to 1
                                    end if
                                end if
                            end repeat
                        end repeat
                    end try
                else
                    if not isTitle then
                        try
                            set oldLeft to left position of shp
                            set shpW to width of shp
                            set newLeft to slideW - oldLeft - shpW
                            if newLeft < 0 then set newLeft to 0
                            set left position of shp to newLeft
                        end try
                    end if
                    if hasText then
                        try
                            set pDir to text direction of paragraph format of text range of text frame of shp
                            if pDir = 1 then
                                set text direction of paragraph format of text range of text frame of shp to 2
                            else if pDir = 2 then
                                set text direction of paragraph format of text range of text frame of shp to 1
                            end if
                        end try
                    end if
                end if
            end try
        end repeat
    end repeat

    save thePres
    delay 1
    close thePres saving no
end tell
return "SUCCESS"
''' % abs_path

    result = subprocess.run(
        ["osascript", "-e", applescript],
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    if result.returncode != 0:
        err = result.stderr.strip() or result.stdout.strip()
        raise RuntimeError(f"PowerPoint mirroring failed: {err}")


def mirror_with_powerpoint(input_path: str, output_path: str, slide_numbers: list = None) -> bool:
    """
    Mirror slide layouts using PowerPoint. Same effect as the VBA macro.
    - Windows: COM (pywin32)
    - Mac: AppleScript (Microsoft PowerPoint for Mac)
    """
    sys = platform.system()
    if sys == "Windows":
        try:
            import win32com.client
        except ImportError:
            raise RuntimeError("pywin32 is required on Windows. Install with: pip install pywin32")

        shutil.copy2(input_path, output_path)
        abs_path = os.path.abspath(output_path)

        app = None
        try:
            app = win32com.client.Dispatch("PowerPoint.Application")
            app.Visible = 0
            pres = app.Presentations.Open(abs_path, WithWindow=False)

            try:
                slide_width = pres.PageSetup.SlideWidth
            except Exception:
                slide_width = pres.SlideMaster.Width

            total = pres.Slides.Count
            for idx in range(1, total + 1):
                if slide_numbers is not None and idx not in slide_numbers:
                    continue
                sld = pres.Slides(idx)
                _mirror_slide_via_com(sld, slide_width)

            pres.Save()
            pres.Close()
        finally:
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
        return True

    if sys == "Darwin":
        if not check_powerpoint_available():
            raise RuntimeError("Microsoft PowerPoint for Mac is not installed.")
        shutil.copy2(input_path, output_path)
        abs_path = os.path.abspath(output_path)
        _mirror_with_applescript(abs_path, slide_numbers=slide_numbers)
        return True

    raise RuntimeError("PowerPoint mirroring is only supported on Windows or macOS.")
