# RamithaSWMacro

SOLIDWORKS macros for drawings.

| File | What it is |
|---|---|
| `macros/DimensionsLightBlue.swp` | The macro, ready to run in SOLIDWORKS |
| `macros/DimensionsLightBlue.swb` | The same macro as plain text (SOLIDWORKS can also run this) |
| `macros/DimensionsLightBlue.bas` | Source export, so the code can be read and reviewed on GitHub |

## DimensionsLightBlue

For the active drawing (all sheets), the macro:

- puts every dimension on a layer called **Dimensions**, creating it in light blue (RGB 83, 169, 255) if it is missing;
- colours every dimension light blue;
- sets **1 decimal place** for dimension values and tolerances (lengths and angles);
- sets Document Properties → Dimensions so **new** dimensions also use the Dimensions layer and 1 decimal place.

Notes, labels, weld symbols, centre lines and the title block are not changed. The drawing is **not saved**: close without saving to undo.

## How to run

1. Download `macros/DimensionsLightBlue.swp`.
2. Open a drawing in SOLIDWORKS.
3. **Tools → Macro → Run…**, pick `DimensionsLightBlue.swp`.
4. A message says how many dimensions were changed. Save the drawing to keep the changes.

To make it a toolbar button: **Tools → Customize → Commands → Macro**, drag *New Macro Button* onto a toolbar and point it at the `.swp`.

Tested on SOLIDWORKS 2026.
