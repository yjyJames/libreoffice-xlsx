"""LibreOffice discovery and process helpers for UNO socket mode."""

from __future__ import annotations

import os
import socket
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 2002
DEFAULT_TIMEOUT = 20.0


def candidate_homes() -> list[Path]:
    homes: list[Path] = []
    env_home = os.environ.get("LIBRE_OFFICE_HOME")
    if env_home:
        homes.append(Path(env_home))
    return _dedupe(homes)


def find_libreoffice_home() -> Path | None:
    for home in candidate_homes():
        if _program_dir(home):
            return home
    return None


def get_program_dir() -> Path | None:
    home = find_libreoffice_home()
    if not home:
        return None
    return _program_dir(home)


def configure_uno_pythonpath() -> Path | None:
    """Add LibreOffice's program directory so standard Python can import uno."""
    program = get_program_dir()
    if not program:
        return None

    program_str = str(program)
    if program_str not in sys.path:
        sys.path.insert(0, program_str)

    path = os.environ.get("PATH", "")
    parts = path.split(os.pathsep) if path else []
    if program_str not in parts:
        os.environ["PATH"] = program_str + (os.pathsep + path if path else "")

    return program


def get_soffice_path() -> Path | None:
    program = get_program_dir()
    if program:
        for name in ("soffice.exe", "soffice"):
            soffice = program / name
            if soffice.exists():
                return soffice

    found = shutil.which("soffice") or shutil.which("libreoffice")
    return Path(found) if found else None


def is_socket_open(
    host: str = DEFAULT_HOST,
    port: int = DEFAULT_PORT,
    timeout: float = 1.0,
) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def build_soffice_args(
    host: str = DEFAULT_HOST,
    port: int = DEFAULT_PORT,
    *,
    headless: bool = False,
    profile: str | Path | None = None,
) -> list[str]:
    soffice = get_soffice_path()
    if not soffice:
        raise FileNotFoundError("LibreOffice soffice executable was not found.")

    args = [
        str(soffice),
        f"--accept=socket,host={host},port={port};urp;StarOffice.ComponentContext",
        "--norestore",
        "--nofirststartwizard",
    ]
    if headless:
        args.append("--headless")
    if profile:
        args.append(f"--env:UserInstallation={Path(profile).resolve().as_uri()}")
    return args


def create_isolated_profile() -> Path:
    return Path(tempfile.mkdtemp(prefix="libreoffice-xlsx-profile-"))


def start_soffice(
    host: str = DEFAULT_HOST,
    port: int = DEFAULT_PORT,
    *,
    headless: bool = False,
    profile: str | Path | None = None,
) -> subprocess.Popen[bytes]:
    args = build_soffice_args(host, port, headless=headless, profile=profile)
    return subprocess.Popen(
        args,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        close_fds=os.name != "nt",
    )


def wait_for_socket(
    host: str = DEFAULT_HOST,
    port: int = DEFAULT_PORT,
    timeout: float = DEFAULT_TIMEOUT,
) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if is_socket_open(host, port, timeout=0.5):
            return True
        time.sleep(0.25)
    return is_socket_open(host, port, timeout=0.5)


def socket_start_command(host: str = DEFAULT_HOST, port: int = DEFAULT_PORT) -> str:
    accept = f'--accept="socket,host={host},port={port};urp;StarOffice.ComponentContext"'
    ps_accept = f"'--accept=socket,host={host},port={port};urp;StarOffice.ComponentContext'"
    return "\n".join(
        (
            'PowerShell: "$env:LIBRE_OFFICE_HOME\\soffice.exe" '
            f"{ps_accept} --norestore --nofirststartwizard",
            'cmd.exe/.bat: "%LIBRE_OFFICE_HOME%\\soffice.exe" ^\n'
            f"  {accept} ^\n"
            "  --norestore ^\n"
            "  --nofirststartwizard",
            "POSIX shell: "
            f'"${{LIBRE_OFFICE_HOME}}/soffice" \\\n'
            f"  {accept} \\\n"
            "  --norestore \\\n"
            "  --nofirststartwizard",
        )
    )


def connection_help(host: str = DEFAULT_HOST, port: int = DEFAULT_PORT) -> str:
    return (
        f"Could not connect to LibreOffice UNO socket at {host}:{port}.\n"
        "Open LibreOffice from your current shell with a UNO socket listener, "
        "using the matching command form:\n\n"
        f"{socket_start_command(host, port)}\n\n"
        "If LibreOffice is already open, fully close it and reopen it with "
        "the command above."
    )


def _program_dir(home: Path) -> Path | None:
    candidates = (
        home,
        home / "program",
        home / "Contents" / "MacOS",
        home / "MacOS",
    )
    for program in candidates:
        if any((program / name).exists() for name in ("soffice.exe", "soffice", "uno.py")):
            return program
    return None


def _dedupe(paths: list[Path]) -> list[Path]:
    seen: set[str] = set()
    result: list[Path] = []
    for path in paths:
        key = str(path).lower()
        if key not in seen:
            seen.add(key)
            result.append(path)
    return result


if __name__ == "__main__":
    home = find_libreoffice_home()
    print(home if home else "LibreOffice not found")
    print(socket_start_command())
