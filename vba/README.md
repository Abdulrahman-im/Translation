# VBA Macro for RTL Mirroring (reference)

The macro `MirrorTextAndTablesBasedOnAlignment` ungroups shapes, mirrors positions (except the slide title), and toggles text and table direction for RTL.

**This project uses Windows COM (pywin32)** to do the same logic in Python—no macro installation needed. The code in `app/services/powerpoint_mirror.py` replicates this macro via the PowerPoint object model, so deployment only needs Windows + PowerPoint + `pip install pywin32`.

You can still use this VBA file to:

- **Run the macro manually** in PowerPoint (Windows or Mac) on a presentation.
- **Compare or port** the logic; the Python COM implementation mirrors it.
