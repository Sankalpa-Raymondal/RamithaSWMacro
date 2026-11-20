# RamithaSWMacro

SOLIDWORKS macros for engineering drawings.

| Macro | What it does |
|---|---|
| [DimensionsLightBlue](#dimensionslightblue) | Puts every dimension on a light-blue "Dimensions" layer with 1 decimal place |
| [SavePDFs](#savepdfs) | Saves or updates PDFs of the drawings you select, with the same names as the drawings |
| [ManufacturingTables](#manufacturingtables) | Makes one manufacturing table (.xlsx) per part, filled with the part name, drawing name and material |

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
