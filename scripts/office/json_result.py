"""Small JSON result helpers for uno-api commands."""

from __future__ import annotations

import json
import sys
from dataclasses import dataclass
from typing import Any


@dataclass
class UnoApiError(Exception):
    code: str
    message: str
    hint: str | None = None
    data: Any | None = None
    exit_code: int = 1


def print_success(data: Any | None = None) -> None:
    _print({"status": "success", "data": {} if data is None else data})


def print_error(error: UnoApiError) -> None:
    payload: dict[str, Any] = {
        "status": "error",
        "error": {
            "code": error.code,
            "message": error.message,
        },
    }
    if error.hint:
        payload["error"]["hint"] = error.hint
    if error.data is not None:
        payload["data"] = error.data
    _print(payload, stderr=True)


def _print(payload: dict[str, Any], *, stderr: bool = False) -> None:
    stream = sys.stderr if stderr else sys.stdout
    print(json.dumps(payload, ensure_ascii=False, indent=2), file=stream)
