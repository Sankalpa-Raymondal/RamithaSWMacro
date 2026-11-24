# RamithaSWMacro

SOLIDWORKS macros for engineering drawings.

| Macro | What it does |
|---|---|
| [DimensionsLightBlue](#dimensionslightblue) | Puts every dimension on a light-blue "Dimensions" layer with 1 decimal place |
| [SavePDFs](#savepdfs) | Saves or updates PDFs of the drawings you select, with the same names as the drawings |
| [ManufacturingTables](#manufacturingtables) | Makes one manufacturing table (.xlsx) per part, filled with the part name, drawing name and material |
| [SymmetryCenterlines](#symmetrycenterlines) | Finds symmetric views and adds a black centre line with the centre-line symbol |
| [UpdateSheetNumber](#updatesheetnumber) | Writes the sheet number (01, 02, ...) into the title block of every sheet |
| [ViewLabels](#viewlabels) | Makes every view name belong to its own view (so it moves with it) and adds missing names |
| [FillTitleBlock](#filltitleblock) | Fills the title block from a template file (`TitleBlockTemplate.txt`) |

Each macro comes as three files in the `macros` folder:

| File | What it is |
|---|---|
| `.swp` | The macro, ready to run in SOLIDWORKS. **Download this one.** |
| `.swb` | The same macro as plain text. SOLIDWORKS can run this too. |
| `.bas` | Source export, so the code can be read and reviewed on GitHub |

## DimensionsLightBlue

For the active drawing (all sheets), the macro:

- puts every dimension on a layer called **Dimensions**, creating it in light blue (RGB 83, 169, 255) if it is missing;
- colours every dimension light blue;
- sets **1 decimal place** for dimension values and tolerances (lengths and angles);
- sets Document Properties → Dimensions so **new** dimensions also use the Dimensions layer and 1 decimal place.

Notes, labels, weld symbols, centre lines and the title block are not changed. The drawing is **not saved**: close without saving to undo.

Tested on SOLIDWORKS 2026.

## SavePDFs

Saves (or updates) PDFs of engineering drawings.

1. Run the macro. A file dialog opens. You don't need a drawing open first.
2. Go to the folder and select one or more drawings (`.SLDDRW`). **Ctrl+click** picks several, **Shift+click** picks a range, **Ctrl+A** picks all. Click **Open**.
3. Each drawing is saved as a PDF with **all its sheets**, **next to the drawing** and with **the same name** (`Rail guide.SLDDRW` → `Rail guide.pdf`). An existing PDF with that name is **overwritten**, so the PDFs always match the latest drawings.
4. A message lists the PDFs that were saved and any that failed.

The drawings are opened read-only and are never saved or changed. Drawings that were already open stay open; the others are closed again. If a PDF fails, it's usually open in a PDF viewer: close it and run the macro again.

Tested on SOLIDWORKS 2026.

## ManufacturingTables

Makes **one manufacturing table (`.xlsx`) per part**. Uses **LibreOffice** (free), so Microsoft Excel isn't needed.

1. Run the macro. A file dialog opens: select one or more parts (`.SLDPRT`), **Ctrl+click** for several, then click **Open**.
2. For each part the macro looks for its drawing in the same folder: a drawing with the same name, or one whose name contains the part name.
3. It writes `<Part name> - Manufacturing.xlsx` next to the part, in the same style as the project BOM:

| Assembly / Part | Drawing Name | Quantity | Manufacturing Method Phase 1 | Manufacturing Method Phase 2 | Manufacturing Method Phase 3 | Material | Special Notes |
|---|---|---|---|---|---|---|---|
| part name | drawing name | 1 | | | | from the drawing's title block | |

- **Material** is copied from the drawing's title block. If there's no drawing, the part's SOLIDWORKS material is used.
- **Fill in yourself:** Quantity, Phase 1–3 (e.g. *Lathe/CNC Milling*, *Laser Cut*, *Sheet Metal Bending*, *Welding*, or anything else) and Special Notes.

**Running it again** rebuilds the file with the latest drawing data and **keeps your Quantity, Phase 1–3 and Special Notes**. Close the file in LibreOffice before running the macro.

Parts and drawings are opened read-only and are never saved or changed.

Tested on SOLIDWORKS 2026 with LibreOffice.

## SymmetryCenterlines

Finds symmetric drawing views and adds a centre line to each symmetry axis.

1. Open a drawing. Select the views to check (Ctrl+click), or select nothing to check every view on the sheet.
2. Run the macro. For each view it checks left/right and top/bottom symmetry: every line end and arc/hole centre must have a mirror partner.
3. For each symmetry axis it draws a **centre line** across the whole view (Center line style, **black**, on the **Text** layer) and puts the **centre-line symbol** at its end. Both belong to the view, so they move with it.

- Axes that already have a centre line (sketch line or Centerline annotation) are skipped, so it's safe to run again.
- Isometric and other non-straight-on views are skipped; rotated views (e.g. a flat pattern turned 90°) are handled.
- A message lists each view and what was done. The drawing is not saved.

Tested on SOLIDWORKS 2026.

## UpdateSheetNumber

Writes the sheet number into the title block of every sheet: sheet 1 → `01`, sheet 2 → `02`, ... in the order of the sheet tabs.

1. Open a drawing and run the macro.
2. On each sheet it finds the **Sheet:** label in the title block and fills the box next to it.
3. A message lists each sheet and its number.

Run it again after adding, deleting or reordering sheets. The drawing is not saved.

Tested on SOLIDWORKS 2026.

## ViewLabels

> The `.swp` for this macro isn't in the repository yet. Until it is, run `ViewLabels.swb` (Tools → Macro → Run, file type *SW VBA Macros (\*.swp; \*.swb)*).

Makes every view name belong to **its own view**, so dragging a view moves its name with it (and nothing else).

1. Open a drawing and run the macro (it works on the current sheet).
2. **Existing names** (notes whose first line contains "VIEW") are matched to the nearest view. A name that belongs to another view or to the sheet is re-created inside its own view with the same text, font, position and layer, and the old one is deleted.
3. **Views without a name** get one, centred above the view (below it if there's no room above):
   - **FRONT VIEW** for the main view; **TOP / BOTTOM / SIDE / LEFT SIDE VIEW** for projected views, from where they sit next to the main view;
   - **ISOMETRIC**, **EXPLODED ISOMETRIC**, **FLAT PATTERN** or **AUXILIARY VIEW**;
   - plus a `SCALE 1:2` line when the view scale differs from the sheet scale;
   - Century Gothic 13 pt, black, on the **Text** layer.

Section and detail views are skipped (SOLIDWORKS names them itself). Safe to run again. The drawing is not saved.

Tested on SOLIDWORKS 2026.

## FillTitleBlock

> The `.swp` for this macro isn't in the repository yet. Until it is, run `FillTitleBlock.swb` (Tools → Macro → Run, file type *SW VBA Macros (\*.swp; \*.swb)*).

Fills in the title block (bottom right table) of every sheet from a **template file**.

1. Put [`TitleBlockTemplate.txt`](macros/TitleBlockTemplate.txt) in the drawing's folder or next to the macro, and edit it in Notepad. One line per box:

   ```
   Description = {partname}
   Material = {material}
   Approx Weight = {mass}
   Drawn by = Ramitha D. & Sankalpa
   Drawn Date = {date}
   Sheet = {sheet}
   ```

   The label is the heading printed in the title block (without the colon). The value is fixed text and/or fields filled in from SOLIDWORKS: `{partname}`, `{material}`, `{mass}` (kg), `{date}` (yyyy/mm/dd), `{sheet}` (01, 02, ...), `{sheets}`, `{scale}`, `{sheetsize}`, `{filename}`.
2. Open a drawing and run the macro. If no template is found, it asks you to pick one.
3. For each label it finds the heading in the title block and fills the box next to it (right of it or below it), keeping that box's font. If the text is too long for the box, it is made smaller until it fits (also for text that was already there).

With `REMOVE PROJECTION = yes` in the template it also clears the **THIRD ANGLE** text and the projection symbol from the Projection Method box.

By default it **only fills empty boxes**, so text you typed is kept. To replace existing values too, set `OVERWRITE = yes` in the template. A message lists every box and what was done. The drawing is not saved.

Tested on SOLIDWORKS 2026.

## Installing and running a macro in SOLIDWORKS

First, download the `.swp` file from the `macros` folder and keep it somewhere permanent, for example `Documents\SOLIDWORKS Macros\`. A toolbar button points at the file's location, so the button stops working if the file is moved or deleted.

The steps below use DimensionsLightBlue as the example; they're the same for every macro (e.g. pick `SavePDFs.swp` and tooltip `Save PDFs`).

### Option 1: run it once

1. Open a drawing.
2. **Tools → Macro → Run…**
3. Select `DimensionsLightBlue.swp` and click **Open**.
4. Click **OK** on the message, then save the drawing if you want to keep the changes.

SOLIDWORKS remembers recently used macros, so next time you can pick it from the bottom of the **Tools → Macro** menu.

### Option 2: make it a toolbar button (best for everyday use)

1. Open a drawing, then **Tools → Customize…** and go to the **Commands** tab.
2. Under *Toolbars* (the category list), choose **Macro**.
3. Drag the **New Macro Button** icon (the last icon in the row) onto a toolbar, for example the **Annotation** tab of the CommandManager.
4. In the *Customize Macro Button* window:
   - **Macro:** click **…** and select `DimensionsLightBlue.swp`.
   - **Method:** choose **main**.
   - **Tooltip:** e.g. `Dimensions light blue, 1 decimal`.
   - **Icon:** optional. Leave it blank for the default icon.
5. Click **OK**, then **OK** again to close Customize.

From then on, one click runs the macro.

### Option 3: keyboard shortcut

After making the button (Option 2), go to **Tools → Customize… → Keyboard**, search for the tooltip text (e.g. `Dimensions light blue`), click in the **Shortcut(s)** column and press the keys you want, e.g. **Ctrl+Shift+D**. Then click **OK**.

### Troubleshooting

- **"Macros are disabled":** go to **Tools → Options → System Options → General** and make sure *Disable VBA macros* is unticked.
- **The button does nothing or can't find the macro:** the `.swp` was moved or renamed. Edit the button again (Option 2, step 4) and point it at the new location.
- **Can't open the `.swp`:** run the `.swb` instead. It's the same code as plain text. In the Run dialog, set the file type to *SW VBA Macros (\*.swp; \*.swb)*.

## Why the `.swp` files are made by hand

The code is written and tested as a plain-text `.swb`, but each `.swp` is created manually in SOLIDWORKS. That's because a `.swp` isn't a text file:

- **It's a binary VBA project.** A `.swp` is a Microsoft compound file (like an old `.doc`) holding the VBA project: the source code in compressed form, compiled code for the VBA version that saved it, and project records that SOLIDWORKS checks when it opens the file.
- **Only the VBA editor writes it reliably.** SOLIDWORKS has no API command to create or save a `.swp`. While a macro is running, the VBA editor's Save command is disabled, so a macro can't save one either. A `.swp` written directly by a script, without the editor, was rejected by SOLIDWORKS ("Cannot open").
- **Git can't show it.** Because the file is binary, GitHub can't display or compare its contents. That's why every macro also has a `.swb` (runnable text) and a `.bas` (source export) next to it, so the code can be read and reviewed.

**How a `.swp` is made (about 30 seconds):**

1. In SOLIDWORKS: **Tools → Macro → New…**, save it in the `macros` folder as e.g. `SavePDFs.swp`. The VBA editor opens.
2. Open the matching **`.swb`** in Notepad (**not the `.bas`**: its first line, `Attribute VB_Name = ...`, is not valid in the editor and stops the macro from running), select everything (**Ctrl+A**) and copy it (**Ctrl+C**).
3. In the VBA editor, click in the code window, select everything (**Ctrl+A**) and paste (**Ctrl+V**).
4. Press **Ctrl+S** and close the VBA editor.

Each `.swp` is tested in SOLIDWORKS before it's added to the repository.
