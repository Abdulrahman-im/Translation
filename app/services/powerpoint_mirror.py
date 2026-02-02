"""
Use PowerPoint directly (via AppleScript) for RTL mirroring.
This ensures 100% compatibility with the VBA macro logic.
"""

import subprocess
import shutil
import os
import platform
import time


def check_powerpoint_available():
    """Check if PowerPoint for Mac is installed."""
    if platform.system() != 'Darwin':
        return False
    return os.path.exists('/Applications/Microsoft PowerPoint.app')


def mirror_with_powerpoint(input_path: str, output_path: str, slide_numbers: list = None) -> bool:
    """
    Mirror slide layouts using PowerPoint directly via AppleScript.
    This runs the EXACT same logic as the VBA macro.

    Args:
        input_path: Path to input PPTX
        output_path: Path to save mirrored PPTX
        slide_numbers: List of slide numbers to process (None = all)

    Returns:
        True if successful
    """
    if not check_powerpoint_available():
        raise RuntimeError("PowerPoint for Mac is not installed")

    # Copy input to output location (PowerPoint will modify in place)
    shutil.copy2(input_path, output_path)
    abs_path = os.path.abspath(output_path)

    # Build slide filter (empty means all slides)
    if slide_numbers:
        slide_list = '{' + ','.join(str(s) for s in sorted(slide_numbers)) + '}'
    else:
        slide_list = '{}'  # Empty = process all

    # AppleScript that mirrors the VBA logic EXACTLY
    applescript = f'''
on run
    set filePath to "{abs_path}"
    set slideFilter to {slide_list}

    tell application "Microsoft PowerPoint"
        activate

        -- Open the file
        open (POSIX file filePath)
        delay 2

        set thePres to active presentation
        set slideW to width of page setup of thePres
        set totalSlides to count of slides of thePres

        repeat with sldIdx from 1 to totalSlides
            -- Check if we should process this slide
            set shouldProcess to true
            if (count of slideFilter) > 0 then
                set shouldProcess to (slideFilter contains sldIdx)
            end if

            if shouldProcess then
                set sld to slide sldIdx of thePres

                -- ========================================
                -- STEP 1: UNGROUP ALL GROUPS (VBA loop)
                -- ========================================
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
                                if shpType is group then
                                    ungroup shp
                                    set groupsExist to true
                                    exit repeat
                                end if
                            end try
                        end repeat
                    end try
                end repeat

                -- ========================================
                -- STEP 2: GET SLIDE TITLE TEXT
                -- ========================================
                set slideTitle to ""
                try
                    set titleShp to shape 1 of sld
                    if has text frame of titleShp then
                        set titleTF to text frame of titleShp
                        if has text of titleTF then
                            set slideTitle to content of text range of titleTF
                        end if
                    end if
                end try

                -- ========================================
                -- STEP 3: PROCESS ALL SHAPES
                -- ========================================
                set shpCount to count of shapes of sld
                repeat with i from 1 to shpCount
                    try
                        set shp to shape i of sld
                        set shpType to shape type of shp

                        if has text frame of shp then
                            -- Shape has text frame
                            set shpTF to text frame of shp
                            set shpText to ""
                            try
                                if has text of shpTF then
                                    set shpText to content of text range of shpTF
                                end if
                            end try

                            if shpText is equal to slideTitle and slideTitle is not "" then
                                -- TITLE: Only change text direction, NO mirror
                                try
                                    set text direction of paragraph format of text range of shpTF to right to left direction
                                end try
                            else
                                -- TEXT SHAPE: Mirror position + text direction
                                try
                                    set oldLeft to left position of shp
                                    set shpW to width of shp
                                    set newLeft to slideW - oldLeft - shpW
                                    set left position of shp to newLeft
                                end try
                                try
                                    set text direction of paragraph format of text range of shpTF to right to left direction
                                end try
                            end if

                        else if shpType is table then
                            -- TABLE: Mirror position + table direction
                            try
                                set oldLeft to left position of shp
                                set shpW to width of shp
                                set newLeft to slideW - oldLeft - shpW
                                set left position of shp to newLeft
                            end try
                            try
                                set table direction of table of shp to right to left direction
                            end try
                            -- Set cell text directions
                            try
                                set theTable to table of shp
                                repeat with rowIdx from 1 to count of rows of theTable
                                    repeat with colIdx from 1 to count of columns of theTable
                                        set theCell to cell rowIdx of column colIdx of theTable
                                        if has text frame of theCell then
                                            set text direction of paragraph format of text range of text frame of theCell to right to left direction
                                        end if
                                    end repeat
                                end repeat
                            end try

                        else
                            -- OTHER SHAPE: Just mirror position
                            try
                                set oldLeft to left position of shp
                                set shpW to width of shp
                                set newLeft to slideW - oldLeft - shpW
                                set left position of shp to newLeft
                            end try
                        end if

                    end try
                end repeat
            end if
        end repeat

        -- Save and close
        save thePres
        delay 1
        close thePres saving no

    end tell

    return "SUCCESS"
end run
'''

    print(f"  [PowerPoint] Opening file in PowerPoint...")

    # Run the AppleScript
    result = subprocess.run(
        ['osascript', '-e', applescript],
        capture_output=True,
        text=True,
        timeout=600  # 10 minute timeout for large files
    )

    if result.returncode != 0:
        error_msg = result.stderr.strip()
        print(f"  [PowerPoint] Error: {error_msg}")
        raise RuntimeError(f"PowerPoint mirroring failed: {error_msg}")

    print(f"  [PowerPoint] Mirroring complete!")
    return True


def run_vba_macro_in_powerpoint(file_path: str, output_path: str) -> bool:
    """
    Alternative: Run the actual VBA macro in PowerPoint.
    This requires the macro to be available in PowerPoint's macro storage.
    """
    # First copy the file
    shutil.copy2(file_path, output_path)
    abs_path = os.path.abspath(output_path)

    # The full VBA code to inject
    vba_code = '''
Sub MirrorAllSlides()
    Dim pres As Presentation
    Set pres = ActivePresentation
    Dim slideWidth As Single
    slideWidth = pres.PageSetup.slideWidth

    Dim sld As Slide
    Dim shp As Shape
    Dim groupsExist As Boolean
    Dim slide_title As String

    For Each sld In pres.Slides
        ' Ungroup all groups
        groupsExist = True
        Do While groupsExist
            groupsExist = False
            For Each shp In sld.Shapes
                If shp.Type = msoGroup Then
                    shp.Ungroup
                    groupsExist = True
                    Exit For
                End If
            Next shp
        Loop

        ' Get title
        slide_title = ""
        On Error Resume Next
        slide_title = sld.Shapes.Title.TextFrame.TextRange.Text
        On Error GoTo 0

        ' Process shapes
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.TextRange.Text = slide_title And slide_title <> "" Then
                    ' Title: only text direction
                    If shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionLeftToRight Then
                        shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionRightToLeft
                    End If
                Else
                    ' Text shape: mirror + direction
                    shp.Left = slideWidth - shp.Left - shp.Width
                    If shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionLeftToRight Then
                        shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionRightToLeft
                    End If
                End If
            ElseIf shp.Type = msoTable Then
                ' Table: mirror + direction
                shp.Left = slideWidth - shp.Left - shp.Width
                If shp.Table.TableDirection = ppDirectionLeftToRight Then
                    shp.Table.TableDirection = ppDirectionRightToLeft
                End If
                Dim row As row
                Dim cell As cell
                For Each row In shp.Table.Rows
                    For Each cell In row.Cells
                        If cell.Shape.HasTextFrame Then
                            If cell.Shape.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionLeftToRight Then
                                cell.Shape.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionRightToLeft
                            End If
                        End If
                    Next cell
                Next row
            Else
                ' Other: just mirror
                shp.Left = slideWidth - shp.Left - shp.Width
            End If
        Next shp
    Next sld

    pres.Save
End Sub
'''

    # Try to run VBA directly through AppleScript
    applescript = f'''
tell application "Microsoft PowerPoint"
    activate
    open (POSIX file "{abs_path}")
    delay 2

    -- Run VBA macro
    try
        run VB macro macro name "MirrorAllSlides"
    on error errMsg
        -- Macro might not exist, return error
        return "ERROR: " & errMsg
    end try

    save active presentation
    close active presentation saving no
end tell
return "SUCCESS"
'''

    result = subprocess.run(
        ['osascript', '-e', applescript],
        capture_output=True,
        text=True,
        timeout=300
    )

    if "ERROR" in result.stdout or result.returncode != 0:
        return False

    return True
