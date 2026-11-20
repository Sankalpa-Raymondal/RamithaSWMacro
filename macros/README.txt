DimensionsLightBlue.swp  -  SOLIDWORKS drawing macro
====================================================

What it does (to the drawing that is open and active, all sheets):
  - Puts every dimension on a layer called "Dimensions" (creates it in light blue,
    RGB 83,169,255, if the drawing doesn't have it).
  - Colours every dimension light blue.
  - Sets 1 decimal place for dimension values and their tolerances (lengths and angles).
  - Changes Document Properties > Dimensions so that NEW dimensions also go on the
    "Dimensions" layer with 1 decimal place.

It does not change notes, labels, weld symbols, centre lines or the title block,
and it does not save the drawing.

How to run:
  1. Open the drawing in SOLIDWORKS.
  2. Tools > Macro > Run...
  3. Pick DimensionsLightBlue.swp (or DimensionsLightBlue.swb: set the file type to "*.swp; *.swb"), Open.
  4. A message says how many dimensions were changed. Save the drawing to keep it.

Optional: add it as a toolbar button via Tools > Customize > Commands > Macro.

Tested on SOLIDWORKS 2026 (Rail guide drawing, 12 dimensions).
