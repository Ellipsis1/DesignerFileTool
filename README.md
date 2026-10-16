# Patient Admin Tool

A Windows desktop app for dental lab designers. It reads a lab Rx PDF, fills in the case information, and builds the patient folder: CaseNotes, the Rx, and the case files (STL, DCM, screenshots), each renamed and filed to a fixed convention.

It replaces a legacy Excel sheet, and the names and CaseNotes text it produces match that sheet exactly.

## What it does

A case is handled in two steps.

1. **Start the case.** Drop the Rx PDF onto the window. The form is pre-filled from the PDF; correct anything it missed, then click **Create Initial Folder + CaseNotes**. A working folder appears on the Desktop holding fresh CaseNotes and the Rx PDF (the PDF is moved in). Save models and screenshots there and edit the CaseNotes by hand as the design progresses.
2. **Finish the case.** Drop the working folder's CaseNotes file onto the window. The form is filled from it and the folder's case files are listed, each with a guessed role you can change. Click **Create Folder + Notes + Files** to build the dated patient folder, with every file copied into the right subfolder under its standard name.

Dropping an Rx PDF, adding the case files and clicking **Create Folder + Notes + Files** does both steps at once.

### Output

For patient `T. Est 1234-QWER` at the `Chicago` center:

```
Desktop\
  T. EST 1234-QWER Chicago\                 working folder (step 1)
    !1234-QWER_T_EST_CaseNotes.txt
    1234-QWER_T_EST_Rx.pdf
    ...the designer's own files

  10.2.2026 T. EST 1234-QWER Chicago\       finished folder (step 2), dated by the Rx date
    1234-QWER_T_EST_CaseNotes.txt
    1234-QWER_T_EST_Rx.pdf
    1234-QWER STL\
      1234-QWER_MX_T_EST_Chicago.stl
      1234-QWER_MDL_MX_T_EST_Chicago.stl
    3D Viewer\
      1234-QWER_MX_T_EST.dcm
      1234-QWER_WRK_MX_T_EST.stl
      1234-QWER_MDL_MX_T_EST_Chicago.stl
    Design Screenshots\
      ...screenshots, original names kept
```

### Revisions

Step 2 is repeated for each round: the first send to the doctor, every revision, and the final submission once the STL files and zip have been added to the working folder.

For a revision, save the new files in a subfolder of the working folder named `Revision 1` (`Rev 2`, `Redesign 3` and the like work too), or put the mark in the file name (`upper Rev 1.dcm`). Files in the working folder itself, or in an `Initial Design` subfolder, are the initial design. The **Design** column shows what was detected and can be changed per file.

Once a case has a revision, `3D Viewer` and `Design Screenshots` get a subfolder per round so the doctor can compare them; file names are the same standard names in each. The STL folder and the patient folder itself stay flat.

```
    3D Viewer\
      Initial Design\
        1234-QWER_MX_T_EST.dcm
      Revision 1\
        1234-QWER_MX_T_EST.dcm
    Design Screenshots\
      Initial Design\
      Revision 1\
```

A case with no revision is laid out flat, as shown under Output. If the folder from the first send is still on the Desktop when a revision is added, its flat files stay where they are beside the new subfolders.

### Safeguards

- Files are copied into the finished folder, never moved; the working folder stays complete.
- Existing case files are never overwritten. A name that is already taken gets `_02`, `_03`, ... appended.
- A file that is already in the finished folder, unchanged, is not copied again, so step 2 can be re-run for each round.
- CaseNotes loaded from a file are not rebuilt, so hand edits survive. If the form differs from the file, the changes are listed and you are asked before anything is written.
- Only STL files belong in the STL folder. If any other file type is headed there, a warning lists the files and asks whether to continue.
- Zip files are copied into the patient folder under their own name. If a zip's name has neither the unique ID (or its first or last four characters) nor the patient name in it, a warning lists it as possibly belonging to another case and asks whether to continue.
- Every file needs a role before the folder is created; files the app cannot identify are flagged for you to choose.

## Running it

### From the exe

Put `PatientAdminTool.exe` and `lists.json` in the same folder and run the exe. No Python install is needed.

### From source

Requires Windows and Python 3 with Tkinter (developed on 3.13).

```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python PatientAdminTool.py
```

`pdfplumber` (PDF reading) and `tkinterdnd2` (drag and drop) are optional. Without them the app still runs, with manual entry and click-to-browse instead.

## Dropdown lists (`lists.json`)

The designer, center and arch-type dropdowns are read from a `lists.json` beside the app. The file is not in this repository; without it the app runs on short placeholder lists and the title bar shows "built-in lists".

```json
{
  "version": "2026-10-02",
  "designers": ["Designer One", "Designer Two"],
  "centers": ["Chicago", "Denver"],
  "upper_arch": ["UAO4", "UAO4 - Try-in"],
  "lower_arch": ["LAO4", "LAO4 - Try-in"],
  "double_arch": ["DAO4"],
  "other_arch": ["MODEL ONLY"]
}
```

- Any key that is missing or empty keeps its built-in list.
- The arch group a type sits in decides whether the case is treated as upper, lower or both.
- Change `version` whenever you edit the file; it is shown in the title bar.
- **File → Reload Lists** re-reads the file without restarting the app.

## Building the exe

```
pyinstaller --noconfirm --clean --onefile --windowed --name PatientAdminTool --collect-all tkinterdnd2 PatientAdminTool.py
```

The exe is written to `dist\`. Copy `lists.json` beside it before handing it out.

## Development

```
python -m unittest -v
```

| File | Contents |
|---|---|
| `patient_core.py` | All logic: naming, CaseNotes text, Rx PDF parsing, file roles and copy planning. No GUI imports. |
| `PatientAdminTool.py` | The Tkinter window only. |
| `test_patient_core.py` | Unit tests for the core module. |
| `CLAUDE.md` | Architecture notes and the invariants the code relies on. |

Rx PDF parsing is best-effort and tied to one lab Rx layout; always check the pre-filled form. Real Rx PDFs contain patient data and must not be committed: keep them in the gitignored `sample_rx\` folder.
