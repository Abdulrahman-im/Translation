"""
Use PowerPoint directly (via AppleScript) for RTL mirroring.
This ensures 100% compatibility with the VBA macro logic.
"""

import subprocess
import shutil
import os
import platform


def check_powerpoint_available():
    """Check if PowerPoint for Mac is installed."""
    if platform.system() != 'Darwin':
        return False
    return os.path.exists('/Applications/Microsoft PowerPoint.app')


def mirror_with_powerpoint(input_path: str, output_path: str, slide_numbers: list = None) -> bool:
    """
    Mirror slide layouts using PowerPoint directly via AppleScript.
    This runs the EXACT same logic as the VBA macro.
    """
    if not check_powerpoint_available():
        raise RuntimeError("PowerPoint for Mac is not installed")

    # Copy input to output location (PowerPoint will modify in place)
    shutil.copy2(input_path, output_path)
    abs_path = os.path.abspath(output_path)

    # Simpler AppleScript that just runs the VBA macro
    # First, we try to use AppleScript to do basic mirroring
    applescript = f'''
tell application "Microsoft PowerPoint"
    activate

    open (POSIX file "{abs_path}")
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
                        -- Check if it's a group (type 6 in PowerPoint)
                        if shpType is 6 then
                            ungroup shp
                            set groupsExist to true
                            exit repeat
                        end if
                    end try
                end repeat
            end try
        end repeat

        -- GET SLIDE TITLE TEXT
        set slideTitle to ""
        try
            repeat with i from 1 to (count of shapes of sld)
                set shp to shape i of sld
                try
                    if placeholder type of shp is title placeholder then
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
        repeat with i from 1 to shpCount
            try
                set shp to shape i of sld

                -- Get shape text if it has text
                set shpText to ""
                set hasText to false
                try
                    if has text frame of shp then
                        set hasText to true
                        set shpText to content of text range of text frame of shp
                    end if
                end try

                -- Check if it's a table (type 19)
                set shpType to shape type of shp
                set isTable to (shpType is 19)

                if hasText then
                    if shpText is equal to slideTitle and slideTitle is not "" then
                        -- TITLE: Only change text direction, NO mirror
                        try
                            tell text frame of shp
                                set paragraph direction to right to left
                            end tell
                        end try
                    else
                        -- TEXT SHAPE: Mirror position + text direction
                        try
                            set oldLeft to left position of shp
                            set shpW to width of shp
                            set newLeft to slideW - oldLeft - shpW
                            if newLeft < 0 then set newLeft to 0
                            set left position of shp to newLeft
                        end try
                        try
                            tell text frame of shp
                                set paragraph direction to right to left
                            end tell
                        end try
                    end if

                else if isTable then
                    -- TABLE: Mirror position + table direction
                    try
                        set oldLeft to left position of shp
                        set shpW to width of shp
                        set newLeft to slideW - oldLeft - shpW
                        if newLeft < 0 then set newLeft to 0
                        set left position of shp to newLeft
                    end try
                    try
                        set direction of table of shp to right to left
                    end try

                else
                    -- OTHER SHAPE: Just mirror position
                    try
                        set oldLeft to left position of shp
                        set shpW to width of shp
                        set newLeft to slideW - oldLeft - shpW
                        if newLeft < 0 then set newLeft to 0
                        set left position of shp to newLeft
                    end try
                end if

            end try
        end repeat
    end repeat

    save thePres
    delay 1
    close thePres saving no

end tell

return "SUCCESS"
'''

    print(f"  [PowerPoint] Opening file in PowerPoint...")

    # Run the AppleScript
    result = subprocess.run(
        ['osascript', '-e', applescript],
        capture_output=True,
        text=True,
        timeout=600
    )

    if result.returncode != 0:
        error_msg = result.stderr.strip()
        print(f"  [PowerPoint] Error: {error_msg}")
        raise RuntimeError(f"PowerPoint mirroring failed: {error_msg}")

    print(f"  [PowerPoint] Mirroring complete!")
    return True
