# LibreOffice Python SDK for UNO Socket Mode

## Scope

Use this reference for spreadsheet automation through LibreOffice UNO socket mode on the user's current operating system and shell.

Do not use pipe mode. Do not auto-start a headless LibreOffice process. Connect only to an already-running LibreOffice socket listener on `127.0.0.1:2002`.

## Start LibreOffice socket listener

Use the command form that matches the user's shell.

```powershell
$env:LIBRE_OFFICE_HOME = "<LibreOffice program directory>"
$env:PATH = "$env:LIBRE_OFFICE_HOME;$env:PATH"

"$env:LIBRE_OFFICE_HOME\soffice.exe" '--accept=socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext' --norestore --nofirststartwizard
```

```cmd
set "LIBRE_OFFICE_HOME=<LibreOffice program directory>"
set "PATH=%LIBRE_OFFICE_HOME%;%PATH%"
"%LIBRE_OFFICE_HOME%\soffice.exe" ^
  --accept="socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext" ^
  --norestore ^
  --nofirststartwizard
```

```bash
export LIBRE_OFFICE_HOME="/path/to/libreoffice/program"
export PATH="${LIBRE_OFFICE_HOME}:${PATH}"
"${LIBRE_OFFICE_HOME}/soffice" \
  --accept="socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext" \
  --norestore \
  --nofirststartwizard
```

```bash
soffice \
  --accept="socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext" \
  --norestore \
  --nofirststartwizard
```

If LibreOffice is already open and Python cannot connect, fully close LibreOffice and reopen it with the command above.

## Configure Python for UNO

`LIBRE_OFFICE_HOME` should point to the LibreOffice `program` directory when standard Python cannot import `uno` directly. Common values:

```text
Windows LibreOffice program directory
Linux LibreOffice program directory
macOS LibreOffice MacOS directory
```

Set `PATH` before running Python scripts so standard Python can load LibreOffice UNO DLLs/shared libraries and find `soffice`. The Python helper also adds `LIBRE_OFFICE_HOME` directly to `PATH` and `sys.path` for the current Python process.

Expected files:

```text
%LIBRE_OFFICE_HOME%\soffice.exe
%LIBRE_OFFICE_HOME%\uno.py
$LIBRE_OFFICE_HOME/soffice
$LIBRE_OFFICE_HOME/uno.py
```

If `LIBRE_OFFICE_HOME` is unset, check these defaults:

```text
Windows LibreOffice program directory
Linux LibreOffice program directory
macOS LibreOffice MacOS directory
PATH entries for soffice or libreoffice
```

## Connect

```python
import uno

local_ctx = uno.getComponentContext()

resolver = local_ctx.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver",
    local_ctx,
)

ctx = resolver.resolve(
    "uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext"
)

smgr = ctx.ServiceManager

desktop = smgr.createInstanceWithContext(
    "com.sun.star.frame.Desktop",
    ctx,
)
```

## Get current spreadsheet

```python
document = desktop.getCurrentComponent()

if document is None:
    raise RuntimeError("No LibreOffice document is open.")

if not hasattr(document, "Sheets"):
    raise RuntimeError("The current LibreOffice document is not a spreadsheet.")

sheet = document.CurrentController.ActiveSheet
```

## Open a workbook through the existing socket

Use this only after LibreOffice is already running with the socket listener.

```python
from pathlib import Path
import uno

def path_to_url(path):
    return uno.systemPathToFileUrl(str(Path(path).resolve()))

props = []
hidden = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
hidden.Name = "Hidden"
hidden.Value = False
props.append(hidden)

document = desktop.loadComponentFromURL(
    path_to_url(r"C:\path\to\workbook.xlsx"),
    "_blank",
    0,
    tuple(props),
)
```

## Common cell APIs

```python
cell = sheet.getCellRangeByName("A1")

cell.String = "Text"
cell.Value = 123.45
cell.Formula = "=SUM(B2:B10)"

display_text = cell.String
number_value = cell.Value
formula = cell.Formula
```

Use `.String` for displayed text, `.Value` for numeric values, and `.Formula` for formulas.

## Ranges

```python
rng = sheet.getCellRangeByName("A1:D20")
data = rng.getDataArray()
```

`getDataArray()` returns a tuple of row tuples. It is useful for extracting data into Python for analysis.

## Active sheet and named sheets

```python
active = document.CurrentController.ActiveSheet
named = document.Sheets.getByName("Sheet1")
```

## Zero-based cell addressing

UNO position APIs are zero-based:

```python
# A1
cell = sheet.getCellByPosition(0, 0)

# B3
cell = sheet.getCellByPosition(1, 2)
```

Spreadsheet formulas and A1 references remain one-based:

```python
sheet.getCellRangeByName("C3").Formula = "=SUM(A1:B2)"
```

## Recalculate and save

```python
document.calculateAll()
document.store()
```

Always recalculate and save after editing formulas.

## Formula error scanning

Scan the used range of every sheet after recalculation. Count formulas and inspect each cell for known spreadsheet errors.

Prefer UNO error information when available:

```python
err = getattr(cell, "Error", 0)
```

Also inspect displayed text as a fallback:

```python
text = cell.String
```

Treat these displayed values as errors:

```text
#VALUE!
#DIV/0!
#REF!
#NAME?
#NULL!
#NUM!
#N/A
Err:
```

## Connection failure

If the exception message contains `couldn't connect to socket`, tell the user:

1. Fully close LibreOffice.
2. Reopen it from their current shell with the socket listener command.
3. Open the target spreadsheet in LibreOffice.
4. Retry the Python script.

Do not switch to pipe mode or auto-start headless LibreOffice.
