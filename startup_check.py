import argparse
import importlib
import json
import os
import platform
import subprocess
import sys
import traceback
from pathlib import Path
from typing import Any, Dict, List


BASE_DIR = Path(__file__).resolve().parent
REQUIRED_IMPORTS = ["fastapi", "uvicorn", "jinja2", "openpyxl", "PIL"]
REQUIRED_ROUTES = [
    "/submit",
    "/query",
    "/admin/login",
    "/admin/dashboard",
    "/admin",
    "/admin/schedules",
    "/admin/employees",
    "/admin/shift-types",
    "/__version",
    "/healthz",
]


def git_commit() -> str:
    candidates = [
        "git",
        r"C:\Program Files\Git\cmd\git.exe",
        r"C:\Program Files\Git\bin\git.exe",
    ]
    for git_exe in candidates:
        try:
            result = subprocess.run(
                [git_exe, "rev-parse", "--short", "HEAD"],
                cwd=BASE_DIR,
                capture_output=True,
                text=True,
                timeout=5,
                check=False,
            )
        except (OSError, subprocess.TimeoutExpired):
            continue
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip()
    return "unknown"


def port_owners(port: int = 8701) -> List[str]:
    try:
        result = subprocess.run(
            ["netstat", "-ano"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return []
    owners: List[str] = []
    needle = f":{port}"
    for line in result.stdout.splitlines():
        if needle in line and "LISTENING" in line.upper():
            owners.append(" ".join(line.split()))
    return owners


def check_startup() -> Dict[str, Any]:
    result: Dict[str, Any] = {
        "ok": False,
        "cwd": str(Path.cwd()),
        "base_dir": str(BASE_DIR),
        "python_executable": sys.executable,
        "python_version": platform.python_version(),
        "git_commit": git_commit(),
        "main_file": str(BASE_DIR / "main.py"),
        "required_imports": {},
        "main_imported": False,
        "app_exists": False,
        "route_count": 0,
        "missing_routes": list(REQUIRED_ROUTES),
        "port_8701": port_owners(8701),
        "errors": [],
    }

    if sys.version_info < (3, 10):
        result["errors"].append("Python 3.10 or newer is required.")

    for module_name in REQUIRED_IMPORTS:
        try:
            importlib.import_module(module_name)
            result["required_imports"][module_name] = True
        except Exception:
            result["required_imports"][module_name] = False
            result["errors"].append(f"Cannot import dependency: {module_name}")
            result.setdefault("tracebacks", {})[module_name] = traceback.format_exc()

    try:
        if str(BASE_DIR) not in sys.path:
            sys.path.insert(0, str(BASE_DIR))
        main = importlib.import_module("main")
        result["main_imported"] = True
    except Exception:
        result["errors"].append("Cannot import main.py")
        result["main_traceback"] = traceback.format_exc()
        return result

    app = getattr(main, "app", None)
    result["app_exists"] = app is not None
    if app is None:
        result["errors"].append("main.app is missing.")
        return result

    routes = {getattr(route, "path", "") for route in getattr(app, "routes", [])}
    result["route_count"] = len(getattr(app, "routes", []))
    result["missing_routes"] = [path for path in REQUIRED_ROUTES if path not in routes]
    if result["missing_routes"]:
        result["errors"].append("Required routes are missing.")

    result["ok"] = not result["errors"]
    return result


def print_text(result: Dict[str, Any]) -> None:
    print("STARTUP_CHECK")
    print(f"OK={result['ok']}")
    print(f"CURRENT_DIR={result['cwd']}")
    print(f"BASE_DIR={result['base_dir']}")
    print(f"PYTHON_EXE={result['python_executable']}")
    print(f"PYTHON_VERSION={result['python_version']}")
    print(f"GIT_COMMIT={result['git_commit']}")
    print(f"MAIN_FILE={result['main_file']}")
    print(f"ROUTE_COUNT={result['route_count']}")
    print(f"MISSING_ROUTES={result['missing_routes']}")
    print(f"PORT_8701={result['port_8701']}")
    print("DEPENDENCIES=" + json.dumps(result["required_imports"], ensure_ascii=False, sort_keys=True))
    if result["errors"]:
        print("ERRORS:")
        for error in result["errors"]:
            print(f"- {error}")
    for module_name, module_traceback in result.get("tracebacks", {}).items():
        print(f"TRACEBACK dependency {module_name}:")
        print(module_traceback)
    if result.get("main_traceback"):
        print("TRACEBACK main.py:")
        print(result["main_traceback"])


def main() -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except AttributeError:
        pass
    parser = argparse.ArgumentParser(description="Run startup diagnostics for the local store request service.")
    parser.add_argument("--format", choices=["text", "json"], default="text")
    args = parser.parse_args()

    os.environ.setdefault("PYTHONUTF8", "1")
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    result = check_startup()
    if args.format == "json":
        print(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True))
    else:
        print_text(result)
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
