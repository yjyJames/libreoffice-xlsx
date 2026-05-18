#!/usr/bin/env python3
"""Command-line UNO helpers for Calc workbooks."""

from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from office import calc_ops
from office.json_result import UnoApiError, print_error, print_success
from office.soffice import (
    DEFAULT_HOST,
    DEFAULT_PORT,
    DEFAULT_TIMEOUT,
    create_isolated_profile,
    get_python_path,
    start_soffice,
)
from office.uno_session import connect_desktop, get_current_spreadsheet, open_spreadsheet, save_as

SESSION_DIR = ".uno-api"
SESSION_FILE = "session.json"


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        data = args.func(args)
        print_success(data)
        return 0
    except UnoApiError as exc:
        print_error(exc)
        return exc.exit_code
    except ImportError as exc:
        error = UnoApiError(
            "UNO_IMPORT_FAILED",
            f"Could not import LibreOffice UNO Python module: {exc}",
            import_uno_hint(),
        )
        print_error(error)
        return error.exit_code
    except RuntimeError as exc:
        error = map_runtime_error(exc)
        print_error(error)
        return error.exit_code
    except Exception as exc:
        error = UnoApiError("UNO_API_FAILED", f"{type(exc).__name__}: {exc}")
        print_error(error)
        return error.exit_code


def build_parser() -> argparse.ArgumentParser:
    parser = JsonArgumentParser(prog="uno-api")
    subparsers = parser.add_subparsers(dest="command", required=True)

    def add_connection_options(subparser: argparse.ArgumentParser) -> None:
        subparser.add_argument("--host", default=None)
        subparser.add_argument("--port", type=int, default=None)

    connect = subparsers.add_parser("connect", help="Connect to a running LibreOffice UNO listener (does not start one).")
    add_connection_options(connect)
    connect.set_defaults(func=cmd_connect)

    start = subparsers.add_parser("start", help="Start LibreOffice with UNO socket listener.")
    add_connection_options(start)
    start.add_argument("--headless", action="store_true")
    start.add_argument("--isolated-profile", action="store_true")
    start.add_argument("--timeout", type=float, default=DEFAULT_TIMEOUT)
    start.set_defaults(func=cmd_start)

    status = subparsers.add_parser("status", help="Show listener and document status.")
    add_connection_options(status)
    status.set_defaults(func=cmd_status)

    open_cmd = subparsers.add_parser("open", help="Open a spreadsheet.")
    add_connection_options(open_cmd)
    open_cmd.add_argument("file")
    open_cmd.set_defaults(func=with_document(cmd_open, needs_document=False))

    get_cell = subparsers.add_parser("getCell", help="Read a cell.")
    add_connection_options(get_cell)
    get_cell.add_argument("address")
    get_cell.add_argument("--sheet")
    get_cell.add_argument("--type", choices=("display", "value", "formula", "all"), default="all")
    get_cell.set_defaults(func=with_document(cmd_get_cell))

    set_cell = subparsers.add_parser("setCell", help="Write a cell.")
    add_connection_options(set_cell)
    set_cell.add_argument("address")
    set_cell.add_argument("--sheet")
    group = set_cell.add_mutually_exclusive_group(required=True)
    group.add_argument("--text")
    group.add_argument("--value", type=float)
    group.add_argument("--formula")
    set_cell.add_argument("--no-calc", action="store_true")
    set_cell.set_defaults(func=with_document(cmd_set_cell))

    get_range = subparsers.add_parser("getRange", help="Read a cell range.")
    add_connection_options(get_range)
    get_range.add_argument("address")
    get_range.add_argument("--sheet")
    get_range.add_argument("--mode", choices=("data", "display", "formula"), default="data")
    get_range.set_defaults(func=with_document(cmd_get_range))

    set_range = subparsers.add_parser("setRange", help="Write a 2D JSON array to a range.")
    add_connection_options(set_range)
    set_range.add_argument("address")
    set_range.add_argument("--sheet")
    set_range.add_argument("--json", required=True)
    set_range.add_argument("--no-calc", action="store_true")
    set_range.set_defaults(func=with_document(cmd_set_range))

    for name, handler in (
        ("listSheets", cmd_list_sheets),
        ("activeSheet", cmd_active_sheet),
        ("recalc", cmd_recalc),
        ("save", cmd_save),
        ("formulaErrors", cmd_formula_errors),
    ):
        subparser = subparsers.add_parser(name)
        add_connection_options(subparser)
        subparser.set_defaults(func=with_document(handler))

    select_sheet = subparsers.add_parser("selectSheet", help="Activate a sheet.")
    add_connection_options(select_sheet)
    select_sheet.add_argument("sheet")
    select_sheet.set_defaults(func=with_document(cmd_select_sheet))

    used_range = subparsers.add_parser("usedRange", help="Return the used range.")
    add_connection_options(used_range)
    used_range.add_argument("--sheet")
    used_range.set_defaults(func=with_document(cmd_used_range))

    save_as_cmd = subparsers.add_parser("saveAs", help="Save to a new path.")
    add_connection_options(save_as_cmd)
    save_as_cmd.add_argument("file")
    save_as_cmd.set_defaults(func=with_document(cmd_save_as))

    return parser


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        payload = {
            "status": "error",
            "error": {
                "code": "ARGUMENT_ERROR",
                "message": message,
                "hint": f"Run: {self.prog} --help",
            },
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2), file=sys.stderr)
        raise SystemExit(2)


def import_uno_hint() -> str:
    python = get_python_path()
    if python:
        return f'Run this script with LibreOffice Python: "{python}" scripts/uno_api.py ...'
    return (
        "Set LIBRE_OFFICE_HOME to the LibreOffice installation directory and run this "
        "script with LibreOffice's program/python executable."
    )


def cmd_connect(args: argparse.Namespace) -> dict[str, Any]:
    host, port = connection_target(args, use_session=False)
    desktop = connect_existing_desktop(host, port)

    if desktop is None:
        raise UnoApiError(
            "UNO_SOCKET_CONNECT_FAILED",
            f"Could not connect to LibreOffice UNO socket at {host}:{port}.",
            "Run: uno-api start",
        )

    document = None
    try:
        document = get_current_spreadsheet(desktop)
    except RuntimeError:
        document = None

    session = {
        "host": host,
        "port": port,
        "pid": None,
        "profile": None,
        "started_by_uno_api": False,
        "created_at": datetime.now(timezone.utc).astimezone().isoformat(),
    }
    write_session(session)
    return {
        "connected": True,
        "started": False,
        "host": host,
        "port": port,
        "pid": None,
        "profile": None,
        "document": document_info(document) if document else None,
    }


def cmd_start(args: argparse.Namespace) -> dict[str, Any]:
    host, port = connection_target(args, use_session=False)
    profile = None
    if args.isolated_profile:
        profile = create_isolated_profile()
    try:
        process = start_soffice(
            host,
            port,
            headless=args.headless,
            profile=profile,
        )
    except FileNotFoundError as exc:
        raise UnoApiError(
            "LIBREOFFICE_NOT_FOUND",
            str(exc),
            "Set LIBRE_OFFICE_HOME to the LibreOffice installation directory.",
        ) from exc
    desktop = wait_for_uno_desktop(host, port, args.timeout)
    if desktop is None:
        raise UnoApiError(
            "UNO_SOCKET_CONNECT_FAILED",
            f"Timed out waiting for LibreOffice UNO socket at {host}:{port}.",
        )

    document = None
    try:
        document = get_current_spreadsheet(desktop)
    except RuntimeError:
        document = None

    session = {
        "host": host,
        "port": port,
        "pid": process.pid,
        "profile": str(profile) if profile else None,
        "started_by_uno_api": True,
        "created_at": datetime.now(timezone.utc).astimezone().isoformat(),
    }
    write_session(session)
    return {
        "connected": True,
        "started": True,
        "host": host,
        "port": port,
        "pid": session["pid"],
        "profile": session["profile"],
        "document": document_info(document) if document else None,
    }


def connect_existing_desktop(host: str, port: int) -> Any | None:
    try:
        return connect_desktop(host, port)
    except RuntimeError as exc:
        if is_socket_connect_failure(exc):
            return None
        raise


def wait_for_uno_desktop(host: str, port: int, timeout: float) -> Any | None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        desktop = connect_existing_desktop(host, port)
        if desktop is not None:
            return desktop
        time.sleep(0.25)
    return connect_existing_desktop(host, port)


def cmd_status(args: argparse.Namespace) -> dict[str, Any]:
    host, port = connection_target(args)
    data: dict[str, Any] = {
        "host": host,
        "port": port,
        "socketOpen": False,
        "connected": False,
        "document": None,
    }
    desktop = connect_existing_desktop(host, port)
    if desktop is None:
        return data
    data["socketOpen"] = True
    data["connected"] = True
    try:
        document = get_current_spreadsheet(desktop)
    except RuntimeError:
        return data
    data["document"] = document_info(document)
    data["activeSheet"] = calc_ops.get_sheet_name(document.CurrentController.ActiveSheet)
    data["sheetCount"] = document.Sheets.getCount()
    return data


def cmd_open(args: argparse.Namespace, desktop: Any, document: Any | None) -> dict[str, Any]:
    opened = open_spreadsheet(desktop, args.file)
    return document_info(opened)


def cmd_get_cell(args: argparse.Namespace, desktop: Any, document: Any) -> dict[str, Any]:
    cell = calc_ops.get_cell(calc_ops.get_sheet(document, args.sheet), args.address)
    if args.type == "all":
        return cell
    return {key: cell[key] for key in ("sheet", "address", args.type)}


def cmd_set_cell(args: argparse.Namespace, desktop: Any, document: Any) -> dict[str, Any]:
    sheet = calc_ops.get_sheet(document, args.sheet)
    return calc_ops.set_cell(
        document,
        sheet,
        args.address,
        text=args.text,
        value=args.value,
        formula=args.formula,
        calculate=not args.no_calc,
    )


def cmd_get_range(args: argparse.Namespace, desktop: Any, document: Any) -> dict[str, Any]:
    return calc_ops.get_range(calc_ops.get_sheet(document, args.sheet), args.address, args.mode)


def cmd_set_range(args: argparse.Namespace, desktop: Any, document: Any) -> dict[str, Any]:
    try:
        values = json.loads(args.json)
    except json.JSONDecodeError as exc:
        raise UnoApiError("RANGE_DATA_INVALID", f"Invalid JSON: {exc}") from exc
    return calc_ops.set_range(
        document,
        calc_ops.get_sheet(document, args.sheet),
        args.address,
        values,
        calculate=not args.no_calc,
    )


def cmd_list_sheets(args: argparse.Namespace, desktop: Any, document: Any) -> dict[str, Any]:
    return {"sheets": calc_ops.list_sheet_names(document)}


def cmd_active_sheet(args: argparse.Namespace, desktop: Any, document: Any) -> dict[str, Any]:
    sheet = document.CurrentController.ActiveSheet
    return {"sheet": calc_ops.get_sheet_name(sheet)}


def cmd_select_sheet(args: argparse.Namespace, desktop: Any, document: Any) -> dict[str, Any]:
    sheet = calc_ops.select_sheet(document, args.sheet)
    return {"sheet": calc_ops.get_sheet_name(sheet)}


def cmd_recalc(args: argparse.Namespace, desktop: Any, document: Any) -> dict[str, Any]:
    document.calculateAll()
    return {"calculated": True}


def cmd_save(args: argparse.Namespace, desktop: Any, document: Any) -> dict[str, Any]:
    try:
        document.store()
    except Exception as exc:
        raise UnoApiError("SAVE_FAILED", f"Could not save document: {exc}") from exc
    return {"saved": True, "document": document_info(document)}


def cmd_save_as(args: argparse.Namespace, desktop: Any, document: Any) -> dict[str, Any]:
    try:
        save_as(document, args.file)
    except Exception as exc:
        raise UnoApiError("SAVE_FAILED", f"Could not save document as {args.file}: {exc}") from exc
    return {"saved": True, "document": document_info(document)}


def cmd_used_range(args: argparse.Namespace, desktop: Any, document: Any) -> dict[str, Any]:
    return calc_ops.used_range(calc_ops.get_sheet(document, args.sheet))


def cmd_formula_errors(args: argparse.Namespace, desktop: Any, document: Any) -> dict[str, Any]:
    result = calc_ops.formula_errors(document)
    if result["count"]:
        raise UnoApiError(
            "FORMULA_ERRORS_FOUND",
            f"Found {result['count']} formula error(s).",
            data=result,
        )
    return result


def with_document(
    handler: Callable[..., dict[str, Any]],
    *,
    needs_document: bool = True,
) -> Callable[[argparse.Namespace], dict[str, Any]]:
    def wrapper(args: argparse.Namespace) -> dict[str, Any]:
        host, port = connection_target(args)
        desktop = connect_desktop(host, port)
        document = get_current_spreadsheet(desktop) if needs_document else None
        return handler(args, desktop, document)

    return wrapper


def connection_target(
    args: argparse.Namespace,
    *,
    use_session: bool = True,
) -> tuple[str, int]:
    session = read_session() if use_session else {}
    host = args.host or session.get("host") or DEFAULT_HOST
    port = args.port or session.get("port") or DEFAULT_PORT
    return str(host), int(port)


def session_path() -> Path:
    return Path.cwd() / SESSION_DIR / SESSION_FILE


def read_session() -> dict[str, Any]:
    path = session_path()
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def write_session(session: dict[str, Any]) -> None:
    path = session_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(session, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def document_info(document: Any) -> dict[str, Any]:
    return {
        "url": getattr(document, "URL", ""),
        "title": getattr(document, "Title", ""),
        "isSpreadsheet": True,
        "activeSheet": calc_ops.get_sheet_name(document.CurrentController.ActiveSheet),
        "sheetCount": document.Sheets.getCount(),
    }


def map_runtime_error(exc: RuntimeError) -> UnoApiError:
    text = str(exc)
    lowered = text.lower()
    if "no document is open" in lowered:
        return UnoApiError("NO_DOCUMENT_OPEN", text, "Run: uno-api open <workbook.xlsx>")
    if "not a spreadsheet" in lowered:
        return UnoApiError("NOT_SPREADSHEET", text)
    if "could not connect" in lowered or "socket" in lowered:
        return UnoApiError("UNO_SOCKET_CONNECT_FAILED", text, "Run: uno-api start")
    return UnoApiError("UNO_API_FAILED", text)


def is_socket_connect_failure(exc: RuntimeError) -> bool:
    lowered = str(exc).lower()
    return "could not connect" in lowered or "socket" in lowered


if __name__ == "__main__":
    raise SystemExit(main())
