---
name: libreoffice-xlsx
description: "Use this skill any time a spreadsheet file is the primary input or output using LibreOffice Python SDK / UNO socket mode. This means any task where the user wants to: open, read, edit, or fix an existing .xlsx, .xlsm, .csv, or .tsv file (e.g., adding columns, computing formulas, formatting, charting, cleaning messy data); create a new spreadsheet from scratch or from other data sources; or convert between tabular file formats. Trigger especially when the user references a spreadsheet file by name or path — even casually (like \"the xlsx in my downloads\") — and wants something done to it or produced from it. Also trigger for cleaning or restructuring messy tabular data files (malformed rows, misplaced headers, junk data) into proper spreadsheets. The deliverable must be a spreadsheet file. Do NOT trigger when the primary deliverable is a Word document, HTML report, standalone Python script, database pipeline, or Google Sheets API integration, even if tabular data is involved."
license: Proprietary. LICENSE.txt has complete terms
---

# Requirements for Outputs

## All Excel files

### Professional Font
- Use a consistent, professional font (e.g., Arial, Times New Roman) for all deliverables unless otherwise instructed by the user

### Zero Formula Errors
- Every Excel model MUST be delivered with ZERO formula errors (#REF!, #DIV/0!, #VALUE!, #N/A, #NAME?)

### Preserve Existing Templates (when updating templates)
- Study and EXACTLY match existing format, style, and conventions when modifying files
- Never impose standardized formatting on files with established patterns
- Existing template conventions ALWAYS override these guidelines

## Financial models

### Color Coding Standards
Unless otherwise stated by the user or existing template

#### Industry-Standard Color Conventions
- **Blue text (RGB: 0,0,255)**: Hardcoded inputs, and numbers users will change for scenarios
- **Black text (RGB: 0,0,0)**: ALL formulas and calculations
- **Green text (RGB: 0,128,0)**: Links pulling from other worksheets within same workbook
- **Red text (RGB: 255,0,0)**: External links to other files
- **Yellow background (RGB: 255,255,0)**: Key assumptions needing attention or cells that need to be updated

### Number Formatting Standards

#### Required Format Rules
- **Years**: Format as text strings (e.g., "2024" not "2,024")
- **Currency**: Use $#,##0 format; ALWAYS specify units in headers ("Revenue ($mm)")
- **Zeros**: Use number formatting to make all zeros "-", including percentages (e.g., "$#,##0;($#,##0);-")
- **Percentages**: Default to 0.0% format (one decimal)
- **Multiples**: Format as 0.0x for valuation multiples (EV/EBITDA, P/E)
- **Negative numbers**: Use parentheses (123) not minus -123

### Formula Construction Rules

#### Assumptions Placement
- Place ALL assumptions (growth rates, margins, multiples, etc.) in separate assumption cells
- Use cell references instead of hardcoded values in formulas
- Example: Use =B5*(1+$B$6) instead of =B5*1.05

#### Formula Error Prevention
- Verify all cell references are correct
- Check for off-by-one errors in ranges
- Ensure consistent formulas across all projection periods
- Test with edge cases (zero values, negative numbers)
- Verify no unintended circular references

#### Documentation Requirements for Hardcodes
- Comment or in cells beside (if end of table). Format: "Source: [System/Document], [Date], [Specific Reference], [URL if applicable]"
- Examples:
  - "Source: Company 10-K, FY2024, Page 45, Revenue Note, [SEC EDGAR URL]"
  - "Source: Company 10-Q, Q2 2025, Exhibit 99.1, [SEC EDGAR URL]"
  - "Source: Bloomberg Terminal, 8/15/2025, AAPL US Equity"
  - "Source: FactSet, 8/20/2025, Consensus Estimates Screen"

# XLSX creation, editing, and analysis

## Scope and engine

Use LibreOffice Python SDK / pyuno / UNO API as the primary engine for spreadsheet reading, editing, saving, recalculation, and validation.

Use `scripts/uno_api.py` (documented as `uno-api` below) for common operations before writing one-off Python UNO scripts. All commands emit structured JSON so Codex and shell scripts can parse results reliably.

Use the user's current operating system and shell. Provide commands for the shell in use (for example PowerShell, cmd.exe, Bash, or Zsh) and keep the UNO socket connection string the same across shells.

Only use UNO socket listener mode. Do not use pipe mode and do not install LibreOffice Basic macros for recalculation.

Do not use `pandas` or `openpyxl` as the default workbook I/O path for reading, editing, or saving spreadsheets.

## LibreOffice socket workflow

First try to use the existing default UNO connection:

```bash
python scripts/uno_api.py status --host 127.0.0.1 --port 2002
```

If that direct UNO connection succeeds, use the existing LibreOffice session and active workbook. If it fails because `127.0.0.1:2002` is not accepting UNO connections, then connect to or start LibreOffice through the CLI wrapper:

```bash
python scripts/uno_api.py connect
```

If `connect` fails, start a new LibreOffice instance:

```bash
python scripts/uno_api.py start
```

Then open the workbook:

```bash
python scripts/uno_api.py open workbook.xlsx
```

`connect` attempts a real UNO connection to the target socket and only succeeds if LibreOffice is already running; do not use a raw TCP socket probe as the decision point. `start` always launches a new LibreOffice process with a UNO socket listener. With no explicit `--host` or `--port`, the target socket is always the default `127.0.0.1:2002`; do not let an old `.uno-api/session.json` change the target. Use `--headless` only when a non-interactive run is required, and use `--isolated-profile` when the user's existing LibreOffice profile conflicts with automation.

The default socket is:

```text
uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext
```

The command writes `.uno-api/session.json`; later non-`connect` commands use that host/port by default. Override with `--host` and `--port` when needed.

If `uno-api start` cannot start LibreOffice, start LibreOffice from the user's current shell. Use the command form that matches the environment:

```powershell
"$env:LIBRE_OFFICE_HOME\soffice.exe" '--accept=socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext' --norestore --nofirststartwizard
```

```cmd
"%LIBRE_OFFICE_HOME%\soffice.exe" ^
  --accept="socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext" ^
  --norestore ^
  --nofirststartwizard
```

```bash
soffice \
  --accept="socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext" \
  --norestore \
  --nofirststartwizard
```

If Python reports `couldn't connect to socket`, do not attempt alternate connection modes. Run `python scripts/uno_api.py start`, or fully close LibreOffice and reopen it with the socket listener command.

## LIBRE_OFFICE_HOME

Support `LIBRE_OFFICE_HOME` for locating the LibreOffice `program` directory and UNO Python files.

Set `LIBRE_OFFICE_HOME` using the user's current shell:

```powershell
$env:LIBRE_OFFICE_HOME = "<LibreOffice program directory>"
$env:PATH = "$env:LIBRE_OFFICE_HOME;$env:PATH"
```

```cmd
set "LIBRE_OFFICE_HOME=<LibreOffice program directory>"
set "PATH=%LIBRE_OFFICE_HOME%;%PATH%"
```

```bash
export LIBRE_OFFICE_HOME="/path/to/libreoffice/program"
export PATH="${LIBRE_OFFICE_HOME}:${PATH}"
```

Set `PATH` this way before running Python scripts so standard Python can load LibreOffice UNO DLLs/shared libraries and find `soffice`.

Expected files:

```text
Windows: %LIBRE_OFFICE_HOME%\soffice.exe
Windows: %LIBRE_OFFICE_HOME%\uno.py
Linux:   $LIBRE_OFFICE_HOME/soffice
Linux:   $LIBRE_OFFICE_HOME/uno.py
macOS:   $LIBRE_OFFICE_HOME/soffice
```

If `LIBRE_OFFICE_HOME` is unset, check:

```text
Windows LibreOffice program directory
Linux LibreOffice program directory
macOS LibreOffice MacOS directory
PATH entries for soffice or libreoffice
```

For lower-level UNO API examples, read `references/libreoffice-python-sdk.md`.

## Standard workflow

1. Try a direct UNO connection to `127.0.0.1:2002` with `python scripts/uno_api.py status --host 127.0.0.1 --port 2002`.
2. If the direct connection succeeds, use the existing LibreOffice session and active workbook.
3. If the direct connection fails, try `python scripts/uno_api.py connect` to reuse a running LibreOffice. If `connect` fails, run `python scripts/uno_api.py start` to start LibreOffice. Then `python scripts/uno_api.py open <workbook>` to open the workbook.
4. Use `python scripts/uno_api.py status` to confirm a Calc workbook and active sheet are available.
5. Prefer `uno-api` commands for common read/write, sheet, save, recalc, and formula validation operations.
6. Use spreadsheet formulas instead of Python-calculated hardcoded outputs unless the user explicitly asks for static values.
7. Call `python scripts/uno_api.py recalc` after formula edits if the command did not already calculate.
8. Save with `python scripts/uno_api.py save` or `python scripts/uno_api.py saveAs output.xlsx`.
9. Run `python scripts/uno_api.py formulaErrors`; fix reported errors and rerun until `status` is `success`.
10. Write direct Python UNO scripts only for operations not covered by `uno-api` (formatting, charts, advanced sheet management, etc.).

## uno-api commands

### Connect, start, status, open, save

```bash
python scripts/uno_api.py connect
python scripts/uno_api.py start
python scripts/uno_api.py start --headless --isolated-profile
python scripts/uno_api.py status
python scripts/uno_api.py open workbook.xlsx
python scripts/uno_api.py open workbook.xlsx --hidden
python scripts/uno_api.py recalc
python scripts/uno_api.py save
python scripts/uno_api.py saveAs output.xlsx
```

### Cells and ranges

```bash
python scripts/uno_api.py getCell A1 --sheet Sheet1 --type all
python scripts/uno_api.py setCell A1 --text "Revenue"
python scripts/uno_api.py setCell B2 --value 100
python scripts/uno_api.py setCell C3 --formula "=SUM(B2:B10)"
python scripts/uno_api.py getRange A1:D20 --mode data
python scripts/uno_api.py setRange A1:B2 --json '[[1,2],[3,4]]'
```

`setCell --formula` and `setRange` calculate by default. Use `--no-calc` only when batching many writes and run `recalc` afterward.

### Sheets and validation

```bash
python scripts/uno_api.py listSheets
python scripts/uno_api.py activeSheet
python scripts/uno_api.py selectSheet Sheet1
python scripts/uno_api.py usedRange --sheet Sheet1
python scripts/uno_api.py formulaErrors
```

`formulaErrors` scans every sheet's used range. It checks UNO `cell.Error` and common displayed error text such as `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, and `#NAME?`.

## JSON output contract

Successful commands return:

```json
{
  "status": "success",
  "data": {}
}
```

Failed commands return:

```json
{
  "status": "error",
  "error": {
    "code": "UNO_SOCKET_CONNECT_FAILED",
    "message": "Could not connect to LibreOffice UNO socket at 127.0.0.1:2002.",
    "hint": "Run: uno-api start"
  }
}
```

Common error codes include `LIBREOFFICE_NOT_FOUND`, `UNO_IMPORT_FAILED`, `UNO_SOCKET_CONNECT_FAILED`, `NO_DOCUMENT_OPEN`, `NOT_SPREADSHEET`, `SHEET_NOT_FOUND`, `CELL_ADDRESS_INVALID`, `RANGE_ADDRESS_INVALID`, `SAVE_FAILED`, and `FORMULA_ERRORS_FOUND`.

## Direct UNO fallback

When `uno-api` does not cover the required operation, use a concise Python script with the shared helpers instead of rewriting connection boilerplate:

```python
from scripts.office.uno_session import connect_desktop, get_current_spreadsheet

desktop = connect_desktop()
document = get_current_spreadsheet(desktop)
sheet = document.CurrentController.ActiveSheet
sheet.getCellRangeByName("C3").Formula = "=SUM(B2:B10)"
document.calculateAll()
document.store()
```

## CRITICAL: use formulas, not hardcoded calculated values

Always use spreadsheet formulas instead of calculating values in Python and hardcoding them. This ensures the spreadsheet remains dynamic and updateable.

### Wrong - hardcoding calculated values

```python
# Bad: calculating in Python and hardcoding the result.
total = sum(values)
sheet.getCellRangeByName("B10").Value = total
```

### Correct - using spreadsheet formulas

```python
sheet.getCellRangeByName("B10").Formula = "=SUM(B2:B9)"
sheet.getCellRangeByName("C5").Formula = "=(C4-C2)/C2"
sheet.getCellRangeByName("D20").Formula = "=AVERAGE(D2:D19)"
```

This applies to totals, percentages, ratios, differences, lookups, model links, and all other calculations.

## Reading and editing cells directly

```python
text = sheet.getCellRangeByName("A1").String
number = sheet.getCellRangeByName("B2").Value
formula = sheet.getCellRangeByName("C3").Formula
data = sheet.getCellRangeByName("A1:D20").getDataArray()

sheet.getCellRangeByName("A1").String = "Revenue"
sheet.getCellRangeByName("B2").Value = 100
sheet.getCellRangeByName("C3").Formula = "=SUM(B2:B2)"
```

Use `.String` for displayed text, `.Value` for numeric values, and `.Formula` for formula text.

## Code style guidelines

When generating Python code for spreadsheet operations:

- Write minimal, concise Python code without unnecessary comments.
- Avoid verbose variable names and redundant operations.
- Avoid unnecessary print statements.
- Add comments only for non-obvious UNO behavior or important business assumptions.

For Excel files themselves:

- Add comments to cells with complex formulas or important assumptions.
- Document data sources for hardcoded values.
- Include notes for key calculations and model sections.
