"""UNO socket helpers for the libreoffice-xlsx skill."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from .soffice import DEFAULT_HOST, DEFAULT_PORT, configure_uno_pythonpath, connection_help


def import_uno():
    configure_uno_pythonpath()
    import uno  # type: ignore

    return uno


def connect_desktop(host: str = DEFAULT_HOST, port: int = DEFAULT_PORT) -> Any:
    uno = import_uno()
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver",
        local_ctx,
    )

    try:
        ctx = resolver.resolve(
            f"uno:socket,host={host},port={port};urp;StarOffice.ComponentContext"
        )
    except Exception as exc:  # pyuno raises bridge-specific exceptions.
        if _looks_like_socket_failure(exc):
            raise RuntimeError(connection_help(host, port)) from exc
        raise

    smgr = ctx.ServiceManager
    return smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def get_current_spreadsheet(desktop: Any) -> Any:
    document = desktop.getCurrentComponent()
    if document is None:
        raise RuntimeError(
            "LibreOffice is connected, but no document is open. Open the target "
            "spreadsheet in LibreOffice and retry."
        )
    return ensure_spreadsheet(document)


def ensure_spreadsheet(document: Any) -> Any:
    try:
        document.Sheets
    except Exception as exc:
        raise RuntimeError(
            "The current LibreOffice document is not a spreadsheet. Activate a "
            "Calc workbook and retry."
        ) from exc
    return document


def open_spreadsheet(desktop: Any, path: str | Path) -> Any:
    uno = import_uno()
    file_url = uno.systemPathToFileUrl(str(Path(path).resolve()))
    document = desktop.loadComponentFromURL(file_url, "_blank", 0, ())
    if document is None:
        raise RuntimeError(f"LibreOffice could not open spreadsheet: {path}")
    return ensure_spreadsheet(document)


def save_as(document: Any, path: str | Path) -> None:
    uno = import_uno()
    output = Path(path).resolve()
    suffix = output.suffix.lower()
    filter_name = {
        ".xlsx": "Calc MS Excel 2007 XML",
        ".xls": "MS Excel 97",
        ".csv": "Text - txt - csv (StarCalc)",
        ".tsv": "Text - txt - csv (StarCalc)",
        ".ods": "calc8",
    }.get(suffix)
    props = []
    if filter_name:
        props.append(make_property("FilterName", filter_name))
    document.storeAsURL(uno.systemPathToFileUrl(str(output)), tuple(props))


def make_property(name: str, value: Any) -> Any:
    uno = import_uno()
    prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    prop.Name = name
    prop.Value = value
    return prop


def _looks_like_socket_failure(exc: Exception) -> bool:
    text = f"{type(exc).__name__}: {exc}".lower()
    return "couldn't connect to socket" in text or "noconnectexception" in text
