"""Calc workbook operations used by the uno-api CLI."""

from __future__ import annotations

import re
from typing import Any

from .json_result import UnoApiError

CELL_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*$")
RANGE_RE = re.compile(rf"^({CELL_RE.pattern[1:-1]}):({CELL_RE.pattern[1:-1]})$")
ERROR_TEXTS = ("#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NULL!", "#NUM!")


def normalize_cell(address: str) -> str:
    value = address.replace("$", "").upper()
    if not CELL_RE.match(address):
        raise UnoApiError("CELL_ADDRESS_INVALID", f"Invalid cell address: {address}")
    return value


def normalize_range(address: str) -> str:
    value = address.replace("$", "").upper()
    if not RANGE_RE.match(address):
        raise UnoApiError("RANGE_ADDRESS_INVALID", f"Invalid range address: {address}")
    start, end = value.split(":", 1)
    start_col, start_row = split_cell(start)
    end_col, end_row = split_cell(end)
    if col_to_index(start_col) > col_to_index(end_col) or start_row > end_row:
        raise UnoApiError("RANGE_ADDRESS_INVALID", f"Invalid range address: {address}")
    return f"{start}:{end}"


def split_cell(address: str) -> tuple[str, int]:
    match = re.match(r"^([A-Z]+)([0-9]+)$", normalize_cell(address))
    if not match:
        raise UnoApiError("CELL_ADDRESS_INVALID", f"Invalid cell address: {address}")
    return match.group(1), int(match.group(2))


def col_to_index(column: str) -> int:
    index = 0
    for char in column.upper():
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def index_to_col(index: int) -> str:
    result = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        result = chr(ord("A") + rem) + result
    return result


def address_from_position(col: int, row: int) -> str:
    return f"{index_to_col(col)}{row + 1}"


def range_from_address(range_address: Any) -> str:
    start = address_from_position(range_address.StartColumn, range_address.StartRow)
    end = address_from_position(range_address.EndColumn, range_address.EndRow)
    return f"{start}:{end}"


def get_sheet(document: Any, name: str | None = None) -> Any:
    sheets = document.Sheets
    if name:
        if not sheets.hasByName(name):
            raise UnoApiError("SHEET_NOT_FOUND", f"Sheet not found: {name}")
        return sheets.getByName(name)
    return document.CurrentController.ActiveSheet


def list_sheet_names(document: Any) -> list[str]:
    return list(document.Sheets.getElementNames())


def get_sheet_name(sheet: Any) -> str:
    return str(sheet.Name)


def select_sheet(document: Any, name: str) -> Any:
    sheet = get_sheet(document, name)
    document.CurrentController.setActiveSheet(sheet)
    return sheet


def get_cell(sheet: Any, address: str) -> dict[str, Any]:
    address = normalize_cell(address)
    cell = sheet.getCellRangeByName(address)
    return {
        "sheet": get_sheet_name(sheet),
        "address": address,
        "display": cell.String,
        "value": cell.Value,
        "formula": cell.Formula,
        "error": cell.Error,
    }


def set_cell(
    document: Any,
    sheet: Any,
    address: str,
    *,
    text: str | None = None,
    value: float | None = None,
    formula: str | None = None,
    calculate: bool = True,
) -> dict[str, Any]:
    supplied = [item is not None for item in (text, value, formula)].count(True)
    if supplied != 1:
        raise UnoApiError(
            "CELL_VALUE_INVALID",
            "Provide exactly one of --text, --value, or --formula.",
        )

    address = normalize_cell(address)
    cell = sheet.getCellRangeByName(address)
    if text is not None:
        cell.String = text
    elif value is not None:
        cell.Value = value
    elif formula is not None:
        cell.Formula = formula
        if calculate:
            document.calculateAll()
    return get_cell(sheet, address)


def get_range(sheet: Any, address: str, mode: str = "data") -> dict[str, Any]:
    address = normalize_range(address)
    cell_range = sheet.getCellRangeByName(address)
    if mode == "data":
        values = [list(row) for row in cell_range.getDataArray()]
    elif mode in {"display", "formula"}:
        values = _range_grid(cell_range, mode)
    else:
        raise UnoApiError("RANGE_MODE_INVALID", f"Invalid range mode: {mode}")
    return {
        "sheet": get_sheet_name(sheet),
        "range": address,
        "mode": mode,
        "values": values,
    }


def set_range(
    document: Any,
    sheet: Any,
    address: str,
    values: list[list[Any]],
    *,
    calculate: bool = True,
) -> dict[str, Any]:
    address = normalize_range(address)
    if not values or not all(isinstance(row, list) for row in values):
        raise UnoApiError("RANGE_DATA_INVALID", "--json must be a non-empty 2D array.")
    width = len(values[0])
    if width == 0 or any(len(row) != width for row in values):
        raise UnoApiError("RANGE_DATA_INVALID", "--json rows must have equal length.")

    cell_range = sheet.getCellRangeByName(address)
    range_address = cell_range.getRangeAddress()
    expected_rows = range_address.EndRow - range_address.StartRow + 1
    expected_cols = range_address.EndColumn - range_address.StartColumn + 1
    if len(values) != expected_rows or width != expected_cols:
        raise UnoApiError(
            "RANGE_DATA_INVALID",
            f"Range {address} expects {expected_rows}x{expected_cols} values.",
        )

    cell_range.setDataArray(tuple(tuple(row) for row in values))
    if calculate:
        document.calculateAll()
    return get_range(sheet, address, "data")


def used_range(sheet: Any) -> dict[str, Any]:
    cursor = sheet.createCursor()
    cursor.gotoStartOfUsedArea(False)
    cursor.gotoEndOfUsedArea(True)
    address = cursor.getRangeAddress()
    return {
        "sheet": get_sheet_name(sheet),
        "range": range_from_address(address),
        "startColumn": address.StartColumn,
        "startRow": address.StartRow,
        "endColumn": address.EndColumn,
        "endRow": address.EndRow,
    }


def formula_errors(document: Any) -> dict[str, Any]:
    errors: list[dict[str, Any]] = []
    for sheet_name in list_sheet_names(document):
        sheet = get_sheet(document, sheet_name)
        used = sheet.createCursor()
        used.gotoStartOfUsedArea(False)
        used.gotoEndOfUsedArea(True)
        area = used.getRangeAddress()
        for row in range(area.StartRow, area.EndRow + 1):
            for col in range(area.StartColumn, area.EndColumn + 1):
                cell = sheet.getCellByPosition(col, row)
                display = cell.String
                error = int(cell.Error)
                if error or any(token in display for token in ERROR_TEXTS):
                    errors.append(
                        {
                            "sheet": sheet_name,
                            "address": address_from_position(col, row),
                            "formula": cell.Formula,
                            "display": display,
                            "error": error,
                        }
                    )
    return {"count": len(errors), "errors": errors}


def _range_grid(cell_range: Any, mode: str) -> list[list[Any]]:
    address = cell_range.getRangeAddress()
    rows: list[list[Any]] = []
    for row in range(address.EndRow - address.StartRow + 1):
        values: list[Any] = []
        for col in range(address.EndColumn - address.StartColumn + 1):
            cell = cell_range.getCellByPosition(col, row)
            values.append(cell.String if mode == "display" else cell.Formula)
        rows.append(values)
    return rows
