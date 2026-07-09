from __future__ import annotations

import argparse
import base64
import hashlib
import hmac
import html
from html.parser import HTMLParser
import importlib
import json
import os
from pathlib import Path
import re
import secrets
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timedelta
from typing import Any
from urllib.parse import urljoin, urlsplit
from urllib.parse import urlencode
from urllib.error import HTTPError
from urllib.request import build_opener, HTTPRedirectHandler, Request


PROJECT_DIR = Path(__file__).resolve().parents[1]
DOCS_DIR = PROJECT_DIR / "docs"
MAIN_FILE = PROJECT_DIR / "main.py"
ENV_FILE = PROJECT_DIR / ".env"
REAL_DB_PATH = PROJECT_DIR / "data" / "tickets.db"
DEFAULT_BASE_URL = "http://127.0.0.1:8701"
SESSION_COOKIE_NAME = "admin_session"
DEFAULT_SESSION_SECRET = "store-request-tool-local-dev-session-secret"
EXPECTED_ORG_ROLES = ["系统管理员", "总部管理层", "采购", "财务", "设计", "运营经理", "区域经理", "店长", "店员", "兼职"]
LEGACY_COMPAT_ROLE_NAMES = ["运营管理", "商品采购", "排班管理员", "只读账号"]

ERROR_MARKERS = [
    "UndefinedError",
    "TemplateNotFound",
    "Traceback",
    "Internal Server Error",
    "Jinja",
    "detail not found",
]
READ_ONLY_STATIC_GET_PREFIXES = (
    "/admin/export",
    "/admin/archive/export",
    "/admin/schedules/export",
    "/admin/permission-overview/role-checklist/export",
    "/admin/files/",
    "/admin/uploads/",
    "/admin/embed-content/",
)


class NoRedirectHandler(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):  # type: ignore[no-untyped-def]
        return None


class SimpleResponse:
    def __init__(self, status_code: int, headers: dict[str, str], body: bytes, url: str) -> None:
        self.status_code = status_code
        self.headers = headers
        self.content = body
        self.url = url
        self.encoding = "utf-8"
        content_type = headers.get("content-type", headers.get("Content-Type", ""))
        match = re.search(r"charset=([\w-]+)", content_type, re.IGNORECASE)
        if match:
            self.encoding = match.group(1)

    @property
    def text(self) -> str:
        return self.content.decode(self.encoding or "utf-8", errors="replace")

    def json(self) -> Any:
        return json.loads(self.text)


class SimpleClient:
    def __init__(self, timeout: float = 15.0, headers: dict[str, str] | None = None) -> None:
        self.timeout = timeout
        self.headers = dict(headers or {})
        self.cookies: dict[str, str] = {}
        self._opener = build_opener(NoRedirectHandler)

    def __enter__(self) -> "SimpleClient":
        return self

    def __exit__(self, exc_type, exc, traceback) -> None:  # type: ignore[no-untyped-def]
        return None

    def _request(self, method: str, url: str, data: dict[str, str] | None = None) -> SimpleResponse:
        body = None
        request_headers = dict(self.headers)
        if self.cookies:
            request_headers["Cookie"] = "; ".join(f"{key}={value}" for key, value in self.cookies.items())
        if data is not None:
            body = urlencode(data).encode("utf-8")
            request_headers["Content-Type"] = "application/x-www-form-urlencoded"
        request = Request(url, data=body, headers=request_headers, method=method)
        try:
            response = self._opener.open(request, timeout=self.timeout)
            headers = {key.lower(): value for key, value in response.headers.items()}
            return SimpleResponse(int(response.status), headers, response.read(), url)
        except HTTPError as exc:
            headers = {key.lower(): value for key, value in exc.headers.items()}
            return SimpleResponse(int(exc.code), headers, exc.read(), url)

    def get(self, url: str) -> SimpleResponse:
        return self._request("GET", url)

    def post(self, url: str, data: dict[str, str]) -> SimpleResponse:
        return self._request("POST", url, data=data)

PUBLIC_PAGES = ["/", "/submit", "/query", "/schedule"]
UNAUTH_ADMIN_PAGES = [
    "/admin/dashboard",
    "/admin/account",
    "/admin/roles",
    "/admin/settings",
    "/admin/schedules",
    "/admin/system-check",
    "/admin/permission-overview",
]
AUTH_PAGES = [
    "/admin/dashboard",
    "/admin",
    "/admin/my-work",
    "/admin/archive",
    "/admin/trash",
    "/admin/account",
    "/admin/personnel-governance",
    "/admin/personnel-governance?tab=historical",
    "/admin/employees",
    "/admin/shift-types",
    "/admin/schedules",
    "/admin/settings",
    "/admin/ticket-rules",
    "/admin/embedded-pages",
    "/admin/roles",
    "/admin/audit-logs",
    "/admin/permission-overview",
    "/admin/system-check",
    "/admin/system",
    "/admin/route-health",
]
SCAN_PAGES = [
    "/admin/dashboard",
    "/admin",
    "/admin/account",
    "/admin/personnel-governance",
    "/admin/employees",
    "/admin/roles",
    "/admin/schedules",
    "/admin/settings",
    "/admin/ticket-rules",
    "/admin/permission-overview",
    "/admin/system-check",
]

TEXT_CHECKS = {
    "P0 人员权限": [
        ("/admin/account", "人员与账号管理", ["人员与账号管理"]),
        ("/admin/account", "允许登录后台", ["允许登录后台"]),
        ("/admin/account", "参与排班", ["参与排班"]),
        ("/admin/account", "可作为处理人", ["可作为处理人"]),
        ("/admin/account", "组织角色或系统权限角色选择", ["系统权限角色", "角色"]),
        ("/admin/roles", "权限矩阵", ["权限矩阵"]),
        ("/admin/roles", "模块筛选", ["模块筛选"]),
        ("/admin/roles", "关键词搜索", ["关键词搜索"]),
        ("/admin/roles", "恢复默认权限", ["恢复默认权限"]),
        ("/admin/personnel-governance?tab=historical", "历史员工异常修复", ["历史员工异常修复"]),
    ],
    "P1 工单": [
        ("/submit", "建议填写", ["建议填写"]),
        ("/submit", "商品信息建议填写", ["商品信息", "建议填写"]),
        ("/admin/ticket-rules", "自动分派规则", ["自动分派规则"]),
        ("/admin/ticket-rules", "SLA", ["SLA"]),
        ("/admin/ticket-rules", "工单类型模板", ["工单类型模板"]),
    ],
    "P2 排班": [
        ("/admin/schedules", "复制排班", ["复制排班"]),
        ("/admin/schedules", "冲突或异常", ["冲突", "异常"]),
        ("/admin/schedules", "覆盖冲突权限入口", ["忽略冲突", "override_conflict"]),
    ],
    "P3 配置": [
        ("/admin/settings", "门店配置", ["门店"]),
        ("/admin/settings", "品牌配置", ["品牌"]),
        ("/admin/settings", "节假日配置", ["节假日"]),
        ("/admin/settings", "需求类型配置", ["需求类型"]),
    ],
    "P4 看板治理": [
        ("/admin/dashboard", "工单运营看板", ["工单"]),
        ("/admin/dashboard", "排班运营看板", ["排班"]),
        ("/admin/permission-overview", "高风险 POST 未接入为 0", ["高风险 POST 未接入：0"]),
        ("/admin/audit-logs", "审计日志", ["审计日志"]),
    ],
    "P5 系统治理": [
        ("/admin/system-check", "系统检查/体检", ["系统正式使用检查"]),
        ("/admin/system", "系统信息", ["系统"]),
        ("/admin/route-health", "路由健康", ["路由"]),
    ],
}


class LinkFormParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.hrefs: list[str] = []
        self.form_actions: list[dict[str, str]] = []
        self._form_stack: list[dict[str, str]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr_map = {key.lower(): value or "" for key, value in attrs}
        if "href" in attr_map:
            self.hrefs.append(attr_map["href"])
        if tag.lower() == "form":
            form = {
                "method": (attr_map.get("method") or "GET").upper(),
                "action": attr_map.get("action") or "",
            }
            self._form_stack.append(form)
            if form["action"]:
                self.form_actions.append(dict(form))
        if "formaction" in attr_map:
            method = (attr_map.get("formmethod") or "").upper()
            if not method and self._form_stack:
                method = self._form_stack[-1].get("method", "GET")
            self.form_actions.append(
                {
                    "method": method or "POST",
                    "action": attr_map["formaction"],
                }
            )

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() == "form" and self._form_stack:
            self._form_stack.pop()


def now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M")


def parse_env_file() -> dict[str, str]:
    values: dict[str, str] = {}
    if not ENV_FILE.exists():
        return values
    for raw_line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip().strip("\"'")
    return values


def first_admin_credential(env_values: dict[str, str]) -> tuple[str, str, str] | None:
    raw_users = env_values.get("ADMIN_USERS", "")
    for raw_item in raw_users.split(","):
        username, separator, password = raw_item.partition(":")
        if separator and username.strip() and password:
            return username.strip(), password.strip(), "ADMIN_USERS[0]"
    username = env_values.get("ADMIN_USERNAME", "").strip()
    password = env_values.get("ADMIN_PASSWORD", "")
    if username and password:
        return username, password, "ADMIN_USERNAME/ADMIN_PASSWORD"
    return None


def base64url_encode(raw: bytes) -> str:
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def create_readonly_admin_session(username: str, secret: str, max_age_seconds: int = 3600) -> str:
    issued_at = int(time.time())
    payload = base64url_encode(
        json.dumps(
            {
                "username": username,
                "issued_at": issued_at,
                "expires_at": issued_at + max_age_seconds,
                "csrf_token": secrets.token_urlsafe(32),
            },
            ensure_ascii=False,
            separators=(",", ":"),
        ).encode("utf-8")
    )
    signature = hmac.new(secret.encode("utf-8"), payload.encode("ascii"), hashlib.sha256).digest()
    return f"{payload}.{base64url_encode(signature)}"


def git_commit(short: bool = False) -> str:
    git_candidates = [
        "git",
        r"C:\Users\liuhao\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\git\cmd\git.exe",
        r"C:\Program Files\Git\cmd\git.exe",
    ]
    args = ["rev-parse", "--short" if short else "HEAD"]
    for git_exe in git_candidates:
        try:
            result = subprocess.run(
                [git_exe, *args],
                cwd=PROJECT_DIR,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=5,
                check=False,
            )
        except (OSError, subprocess.TimeoutExpired):
            continue
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip()
    head_file = PROJECT_DIR / ".git" / "HEAD"
    try:
        head = head_file.read_text(encoding="utf-8").strip()
        if head.startswith("ref:"):
            ref_file = PROJECT_DIR / ".git" / head.removeprefix("ref:").strip()
            commit = ref_file.read_text(encoding="utf-8").strip()
        else:
            commit = head
        if re.fullmatch(r"[0-9a-fA-F]{7,40}", commit):
            return commit[:7] if short else commit
    except OSError:
        pass
    return "unknown"


def run_check_runtime() -> dict[str, Any]:
    script = PROJECT_DIR / "check_runtime.bat"
    if not script.exists():
        return {"ok": False, "output": "check_runtime.bat not found"}
    result = subprocess.run(
        ["cmd", "/c", "echo.", "|", "check_runtime.bat"],
        cwd=PROJECT_DIR,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=30,
        check=False,
    )
    output = (result.stdout or "") + (result.stderr or "")
    return {
        "ok": result.returncode == 0 and "LOCAL_COMMIT=" in output and "RUNNING_COMMIT=" in output,
        "returncode": result.returncode,
        "output": output.strip(),
    }


def page_result(path: str, response: SimpleResponse | None, error: str = "") -> dict[str, Any]:
    if response is None:
        return {
            "path": path,
            "status": "ERROR",
            "ok_200": False,
            "redirect_303": False,
            "is_403": False,
            "is_404": False,
            "is_500": False,
            "json_white_page": False,
            "error_markers": "",
            "error": error,
            "length": 0,
        }
    text = response.text or ""
    content_type = response.headers.get("content-type", "")
    markers = []
    for marker in ERROR_MARKERS:
        if marker not in text:
            continue
        if marker == "Jinja" and not re.search(r"Jinja(?:2)?\s*(?:error|exception)|jinja2\.exceptions", text, re.IGNORECASE):
            continue
        markers.append(marker)
    return {
        "path": path,
        "status": response.status_code,
        "ok_200": response.status_code == 200,
        "redirect_303": response.status_code == 303,
        "is_403": response.status_code == 403,
        "is_404": response.status_code == 404,
        "is_500": response.status_code >= 500,
        "json_white_page": response.status_code == 200 and "application/json" in content_type and path not in {"/healthz", "/__version"},
        "error_markers": ", ".join(markers),
        "error": "",
        "length": len(text),
    }


def request_get(client: SimpleClient, base_url: str, path: str) -> tuple[SimpleResponse | None, str]:
    try:
        return client.get(urljoin(base_url, path)), ""
    except Exception as exc:  # noqa: BLE001 - report transport failures as acceptance evidence
        return None, str(exc)


def request_post(client: SimpleClient, base_url: str, path: str, data: dict[str, str]) -> tuple[SimpleResponse | None, str]:
    try:
        return client.post(urljoin(base_url, path), data=data), ""
    except Exception as exc:  # noqa: BLE001
        return None, str(exc)


def normalize_internal_url(raw_value: str, source_path: str) -> str | None:
    value = html.unescape((raw_value or "").strip())
    if not value or value.startswith(("#", "javascript:", "mailto:", "tel:")):
        return None
    if "{{" in value or "{%" in value:
        return None
    parsed = urlsplit(value)
    if parsed.scheme and parsed.scheme not in {"http", "https"}:
        return None
    if parsed.scheme and parsed.netloc and parsed.netloc not in {"127.0.0.1:8701", "localhost:8701"}:
        return None
    if value.startswith("http://") or value.startswith("https://"):
        return parsed.path + (f"?{parsed.query}" if parsed.query else "")
    resolved = urljoin(source_path if source_path.endswith("/") else source_path.rsplit("/", 1)[0] + "/", value)
    parsed_resolved = urlsplit(resolved)
    return parsed_resolved.path + (f"?{parsed_resolved.query}" if parsed_resolved.query else "")


def parse_registered_routes() -> list[tuple[str, str]]:
    text = MAIN_FILE.read_text(encoding="utf-8")
    return [(method.upper(), path) for method, path in re.findall(r'@app\.(get|post)\("([^"]+)"', text)]


def route_pattern(path: str) -> re.Pattern[str]:
    escaped = re.escape(path)
    escaped = re.sub(r"\\\{[^{}:]+:path\\\}", ".+", escaped)
    escaped = re.sub(r"\\\{[^{}]+\\\}", "[^/]+", escaped)
    return re.compile(f"^{escaped}$")


def route_exists(routes: list[tuple[str, str]], method: str, path: str) -> bool:
    clean_path = urlsplit(path).path
    for route_method, route_path in routes:
        if route_method == method.upper() and route_pattern(route_path).match(clean_path):
            return True
    return False


def sqlite_readonly_summary() -> dict[str, Any]:
    if not REAL_DB_PATH.exists():
        return {"ok": False, "error": f"{REAL_DB_PATH} not found"}
    uri = f"file:{REAL_DB_PATH.as_posix()}?mode=ro"
    try:
        with sqlite3.connect(uri, uri=True) as connection:
            connection.row_factory = sqlite3.Row
            tables = [
                row["name"]
                for row in connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
                ).fetchall()
            ]
            counts: dict[str, int] = {}
            for table in tables:
                if table.startswith("sqlite_"):
                    continue
                counts[table] = int(connection.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0])
            role_names = []
            if "admin_roles" in counts:
                role_names = [
                    row["role_name"]
                    for row in connection.execute("SELECT role_name FROM admin_roles ORDER BY id").fetchall()
                ]
            return {"ok": True, "table_count": len(tables), "counts": counts, "role_names": role_names}
    except sqlite3.Error as exc:
        return {"ok": False, "error": str(exc)}


def file_fingerprint(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"exists": False}
    stat = path.stat()
    return {"exists": True, "size": stat.st_size, "mtime_ns": stat.st_mtime_ns}


def table(headers: list[str], rows: list[list[Any]]) -> str:
    if not rows:
        return "_无记录_\n"
    output = ["| " + " | ".join(headers) + " |", "| " + " | ".join("---" for _ in headers) + " |"]
    for row in rows:
        output.append("| " + " | ".join(str(value).replace("\n", "<br>") for value in row) + " |")
    return "\n".join(output) + "\n"


def write_report(mode: str, data: dict[str, Any], report_path: Path | None = None) -> Path:
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    path = report_path or DOCS_DIR / f"non_browser_acceptance_report_{now_stamp()}.md"
    issues = data.get("issues", [])
    lines = [
        f"# 非浏览器全功能验收报告 - {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
        f"- 模式：`{mode}`",
        f"- 项目目录：`{PROJECT_DIR}`",
        f"- 本地 commit：`{data.get('local_commit', 'unknown')}`",
        f"- 基准 URL：`{data.get('base_url', DEFAULT_BASE_URL)}`",
        f"- 报告说明：本报告不使用浏览器、不做视觉判断；read-only 模式不会执行真实登录 POST，避免写入登录审计日志。",
        "",
        "## 摘要",
        "",
        f"- 问题/限制数量：{len(issues)}",
        f"- 真实 DB 指纹是否保持不变：{data.get('db_fingerprint_unchanged', '未检查')}",
        f"- 真实 DB 表行数是否保持不变：{data.get('db_row_counts_unchanged', '未检查')}",
        "",
    ]
    if issues:
        lines.extend(["## 问题清单", ""])
        for index, issue in enumerate(issues, start=1):
            lines.append(f"{index}. **{issue.get('severity', 'info')}** `{issue.get('area', '')}`：{issue.get('message', '')}")
        lines.append("")
    else:
        lines.extend(
            [
                "## 问题清单",
                "",
                "未发现自动化阻塞问题；仍保留人工验收项：视觉样式、真实浏览器交互、真实业务写入链路未在 read-only 模式执行。",
                "",
            ]
        )

    if "runtime" in data:
        runtime = data["runtime"]
        lines.extend(
            [
                "## 运行态检查",
                "",
                table(
                    ["项目", "结果"],
                    [
                        ["healthz", runtime.get("healthz", "")],
                        ["__version commit", runtime.get("version_commit", "")],
                        ["commit 是否一致", runtime.get("commit_match", "")],
                        ["check_runtime.bat", "OK" if runtime.get("check_runtime_ok") else "FAIL"],
                    ],
                ),
                "```text",
                str(runtime.get("check_runtime_output", ""))[:3000],
                "```",
                "",
            ]
        )

    for section_key, title in [
        ("public_pages", "公开页面检查"),
        ("unauth_admin_pages", "后台未登录保护检查"),
        ("auth_pages", "后台会话页面检查"),
    ]:
        if section_key in data:
            lines.extend([f"## {title}", ""])
            rows = [
                [
                    item["path"],
                    item["status"],
                    "是" if item["ok_200"] else "",
                    "是" if item["redirect_303"] else "",
                    "是" if item["is_403"] else "",
                    "是" if item["is_404"] else "",
                    "是" if item["is_500"] else "",
                    "是" if item["json_white_page"] else "",
                    item["error_markers"] or item["error"],
                ]
                for item in data[section_key]
            ]
            lines.append(table(["路径", "状态", "200", "303", "403", "404", "500", "JSON 白页", "错误关键词"], rows))
            lines.append("")

    if "text_checks" in data:
        lines.extend(["## 页面文本功能点检查", ""])
        rows = [
            [item["group"], item["path"], item["label"], "通过" if item["passed"] else "失败", item.get("missing", "")]
            for item in data["text_checks"]
        ]
        lines.append(table(["模块", "路径", "检查点", "结果", "缺失"], rows))
        lines.append("")

    if "link_scan" in data:
        scan = data["link_scan"]
        lines.extend(
            [
                "## 链接和表单扫描",
                "",
                f"- 扫描页面数：{scan.get('page_count', 0)}",
                f"- GET 链接数：{scan.get('get_link_count', 0)}",
                f"- 只做静态匹配的 GET 数：{scan.get('static_get_count', 0)}",
                f"- 表单 action 数：{scan.get('form_action_count', 0)}",
                f"- 问题数：{len(scan.get('issues', []))}",
                "",
            ]
        )
        if scan.get("issues"):
            lines.append(table(["来源", "类型", "目标", "状态/原因"], [[i["source"], i["kind"], i["target"], i["detail"]] for i in scan["issues"]]))
            lines.append("")

    if "sqlite" in data:
        sqlite_info = data["sqlite"]
        lines.extend(["## SQLite 只读检查", ""])
        if sqlite_info.get("ok"):
            counts = sqlite_info.get("counts", {})
            important = [
                "tickets",
                "admin_users",
                "admin_roles",
                "admin_role_permissions",
                "employees",
                "shift_types",
                "store_schedules",
                "ticket_assignment_rules",
                "ticket_sla_rules",
                "request_type_templates",
                "admin_operation_logs",
                "admin_login_logs",
            ]
            rows = [[name, counts.get(name, "缺失")] for name in important]
            lines.append(table(["表", "行数"], rows))
            lines.append(f"- admin_roles：{', '.join(sqlite_info.get('role_names', []))}")
        else:
            lines.append(f"- SQLite 只读检查失败：{sqlite_info.get('error')}")
        lines.append("")

    if "isolated_write" in data:
        isolated = data["isolated_write"]
        lines.extend(["## isolated-write 验收", ""])
        lines.append(table(["检查项", "结果"], [[key, value] for key, value in isolated.items()]))
        lines.append("")

    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def collect_text_checks(client: SimpleClient, base_url: str, cache: dict[str, str], issues: list[dict[str, str]]) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    for group, checks in TEXT_CHECKS.items():
        for path, label, terms in checks:
            if path not in cache:
                response, error = request_get(client, base_url, path)
                cache[path] = response.text if response is not None else ""
                if error:
                    issues.append({"severity": "warning", "area": "text-check", "message": f"{path} 请求失败：{error}"})
            text = cache.get(path, "")
            missing = [term for term in terms if term not in text]
            passed = not missing
            if not passed:
                issues.append(
                    {
                        "severity": "warning",
                        "area": "page-text",
                        "message": f"{path} 缺少功能点文案：{label}；缺失 {', '.join(missing)}",
                    }
                )
            results.append({"group": group, "path": path, "label": label, "passed": passed, "missing": ", ".join(missing)})
    submit_text = cache.get("/submit", "")
    if "商品信息必填" in submit_text or "商品信息为必填" in submit_text:
        issues.append({"severity": "risk", "area": "ticket-submit", "message": "/submit 出现商品信息必填类文案，与建议填写目标冲突。"})
    return results


def scan_links_and_forms(client: SimpleClient, base_url: str, page_cache: dict[str, str], issues: list[dict[str, str]]) -> dict[str, Any]:
    registered_routes = parse_registered_routes()
    scan_issues: list[dict[str, str]] = []
    get_links: set[tuple[str, str]] = set()
    form_actions: list[tuple[str, str, str]] = []
    static_get_count = 0
    for source in SCAN_PAGES:
        text = page_cache.get(source)
        if text is None:
            response, error = request_get(client, base_url, source)
            if response is None:
                scan_issues.append({"source": source, "kind": "page", "target": source, "detail": error})
                continue
            text = response.text
            page_cache[source] = text
        parser = LinkFormParser()
        parser.feed(text)
        for href in parser.hrefs:
            target = normalize_internal_url(href, source)
            if target:
                get_links.add((source, target))
        for action in parser.form_actions:
            target = normalize_internal_url(action["action"], source)
            if target:
                form_actions.append((source, action["method"].upper(), target))

    for source, target in sorted(get_links):
        target_path = urlsplit(target).path
        if target_path.startswith(READ_ONLY_STATIC_GET_PREFIXES):
            static_get_count += 1
            if not route_exists(registered_routes, "GET", target):
                scan_issues.append({"source": source, "kind": "GET static", "target": target, "detail": "未匹配 main.py 注册路由"})
            continue
        response, error = request_get(client, base_url, target)
        if response is None:
            scan_issues.append({"source": source, "kind": "GET", "target": target, "detail": error})
            continue
        if response.status_code in {404} or response.status_code >= 500:
            scan_issues.append({"source": source, "kind": "GET", "target": target, "detail": str(response.status_code)})

    for source, method, target in form_actions:
        if not route_exists(registered_routes, method, target):
            scan_issues.append({"source": source, "kind": f"{method} form", "target": target, "detail": "未匹配 main.py 注册路由"})

    for issue in scan_issues:
        issues.append({"severity": "warning", "area": "link-form", "message": f"{issue['source']} {issue['kind']} {issue['target']}：{issue['detail']}"})

    return {
        "page_count": len(SCAN_PAGES),
        "get_link_count": len(get_links),
        "static_get_count": static_get_count,
        "form_action_count": len(form_actions),
        "issues": scan_issues,
    }


def run_read_only(base_url: str, report_path: Path | None) -> Path:
    issues: list[dict[str, str]] = []
    before_db = file_fingerprint(REAL_DB_PATH)
    before_sqlite = sqlite_readonly_summary()
    env_values = parse_env_file()
    credential = first_admin_credential(env_values)
    local_commit = git_commit(short=False)
    local_short = git_commit(short=True)
    data: dict[str, Any] = {"base_url": base_url, "local_commit": local_commit}
    check_runtime = run_check_runtime()

    with SimpleClient(timeout=15.0, headers={"User-Agent": "store-request-tool-acceptance-smoke/1.0"}) as anonymous:
        health_response, health_error = request_get(anonymous, base_url, "/healthz")
        version_response, version_error = request_get(anonymous, base_url, "/__version")
        version_commit = ""
        if version_response is not None:
            try:
                version_commit = str(version_response.json().get("git_commit") or "")
            except ValueError:
                issues.append({"severity": "blocker", "area": "__version", "message": "__version 返回非 JSON。"})
        if health_response is None:
            issues.append({"severity": "blocker", "area": "runtime", "message": f"/healthz 请求失败：{health_error}"})
        elif health_response.status_code != 200:
            issues.append({"severity": "blocker", "area": "runtime", "message": f"/healthz 状态异常：{health_response.status_code}"})
        if version_response is None:
            issues.append({"severity": "blocker", "area": "runtime", "message": f"/__version 请求失败：{version_error}"})
        elif version_commit != local_short:
            issues.append({"severity": "blocker", "area": "runtime", "message": f"运行 commit {version_commit} 与本地 {local_short} 不一致。"})
        data["runtime"] = {
            "healthz": health_response.status_code if health_response is not None else health_error,
            "version_commit": version_commit,
            "commit_match": version_commit == local_short,
            "check_runtime_ok": check_runtime.get("ok"),
            "check_runtime_output": check_runtime.get("output", ""),
        }

        public_results = []
        for path in PUBLIC_PAGES:
            response, error = request_get(anonymous, base_url, path)
            result = page_result(path, response, error)
            if path == "/" and result["redirect_303"]:
                pass
            elif not result["ok_200"]:
                issues.append({"severity": "warning", "area": "public-page", "message": f"{path} 未返回 200：{result['status']}"})
            if result["is_500"] or result["error_markers"]:
                issues.append({"severity": "blocker", "area": "public-page", "message": f"{path} 出现服务端错误或模板错误。"})
            public_results.append(result)
        data["public_pages"] = public_results

        unauth_results = []
        for path in UNAUTH_ADMIN_PAGES:
            response, error = request_get(anonymous, base_url, path)
            result = page_result(path, response, error)
            if not (result["redirect_303"] or result["is_403"] or result["status"] == 401):
                issues.append({"severity": "warning", "area": "auth-guard", "message": f"{path} 未登录状态返回 {result['status']}，不是 303/401/403。"})
            unauth_results.append(result)
        data["unauth_admin_pages"] = unauth_results

    page_cache: dict[str, str] = {}
    with SimpleClient(timeout=15.0, headers={"User-Agent": "store-request-tool-acceptance-smoke/1.0"}) as authed:
        if credential:
            username, _password, source = credential
            secret = env_values.get("SESSION_SECRET", DEFAULT_SESSION_SECRET)
            authed.cookies[SESSION_COOKIE_NAME] = create_readonly_admin_session(username, secret)
            data["admin_username"] = username
            data["admin_credential_source"] = source
        else:
            issues.append({"severity": "blocker", "area": "auth", "message": "未在 .env 找到 ADMIN_USERS 或 ADMIN_USERNAME/ADMIN_PASSWORD。"})

        auth_results = []
        for path in AUTH_PAGES:
            response, error = request_get(authed, base_url, path)
            result = page_result(path, response, error)
            if response is not None and response.status_code == 200:
                page_cache[path] = response.text
            if result["is_404"] or result["is_500"] or result["json_white_page"] or result["error_markers"]:
                issues.append({"severity": "blocker", "area": "admin-page", "message": f"{path} 状态/内容异常：{result}"})
            elif result["status"] not in {200, 303, 403}:
                issues.append({"severity": "warning", "area": "admin-page", "message": f"{path} 返回非预期状态：{result['status']}"})
            auth_results.append(result)
        data["auth_pages"] = auth_results
        data["text_checks"] = collect_text_checks(authed, base_url, page_cache, issues)
        data["link_scan"] = scan_links_and_forms(authed, base_url, page_cache, issues)

    permission_page = page_cache.get("/admin/permission-overview", "")
    match = re.search(r"高风险 POST 未接入[:：]\s*(\d+)", permission_page)
    if not match:
        issues.append({"severity": "warning", "area": "permission-overview", "message": "未能解析高风险 POST 未接入计数。"})
    elif match.group(1) != "0":
        issues.append({"severity": "blocker", "area": "permission-overview", "message": f"高风险 POST 未接入为 {match.group(1)}，不是 0。"})

    data["sqlite"] = sqlite_readonly_summary()
    if data["sqlite"].get("ok"):
        role_names = data["sqlite"].get("role_names", [])
        missing_roles = [role_name for role_name in EXPECTED_ORG_ROLES if role_name not in role_names]
        legacy_roles = [role_name for role_name in LEGACY_COMPAT_ROLE_NAMES if role_name in role_names]
        if missing_roles:
            issues.append({"severity": "blocker", "area": "roles", "message": f"缺少默认组织角色：{', '.join(missing_roles)}"})
        if legacy_roles:
            issues.append({"severity": "warning", "area": "roles", "message": f"真实 DB 仍保留历史兼容角色：{', '.join(legacy_roles)}；需确认后台是否应隐藏或归档展示。"})
    after_db = file_fingerprint(REAL_DB_PATH)
    after_sqlite = sqlite_readonly_summary()
    before_counts = before_sqlite.get("counts", {}) if before_sqlite.get("ok") else {}
    after_counts = after_sqlite.get("counts", {}) if after_sqlite.get("ok") else {}
    data["db_fingerprint_unchanged"] = before_db == after_db
    data["db_row_counts_unchanged"] = before_counts == after_counts
    if before_counts != after_counts:
        issues.append({"severity": "risk", "area": "read-only", "message": "read-only 验收前后真实 DB 行数发生变化。"})
    elif before_db != after_db:
        issues.append({"severity": "info", "area": "read-only", "message": "真实 DB 文件 mtime/指纹变化，但只读表行数未变化；需结合 SQLite/WAL 状态人工确认。"})

    data["issues"] = issues
    return write_report("read-only", data, report_path)


def write_minimal_config(config_dir: Path) -> None:
    if (PROJECT_DIR / "config").is_dir():
        shutil.copytree(PROJECT_DIR / "config", config_dir, dirs_exist_ok=True)
        return
    config_dir.mkdir(parents=True, exist_ok=True)
    defaults: dict[str, Any] = {
        "stores.json": ["隔离验收门店"],
        "request_types.json": ["其他"],
        "urgency_levels.json": ["普通"],
        "statuses.json": ["待处理", "处理中", "已完成"],
        "brands.json": [],
        "handlers.json": ["acceptance-admin"],
        "holidays.json": {},
        "request_type_rules.json": {},
        "system.json": {"app_name": "止痒 ERP", "port": 8701, "page_size": 50, "max_bulk_schedule_count": 200},
    }
    for filename, value in defaults.items():
        (config_dir / filename).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def testclient_csrf_token(client: Any, path: str) -> str:
    response = client.get(path)
    if response.status_code != 200:
        raise RuntimeError(f"{path} did not return 200 for CSRF token: {response.status_code}")
    match = re.search(r'name="csrf_token" value="([^"]+)"', response.text)
    if not match:
        raise RuntimeError(f"{path} did not render csrf_token")
    return match.group(1)


def db_row(db_path: Path, table: str, row_id: int) -> dict[str, Any] | None:
    with sqlite3.connect(db_path) as connection:
        connection.row_factory = sqlite3.Row
        row = connection.execute(f'SELECT * FROM "{table}" WHERE id = ?', (int(row_id),)).fetchone()
        return dict(row) if row else None


def db_scalar(db_path: Path, sql: str, params: tuple[Any, ...] = ()) -> Any:
    with sqlite3.connect(db_path) as connection:
        row = connection.execute(sql, params).fetchone()
        return row[0] if row else None


def role_id_by_name(db_path: Path, role_name: str) -> int:
    role_id = db_scalar(db_path, "SELECT id FROM admin_roles WHERE role_name = ?", (role_name,))
    if role_id is None:
        raise RuntimeError(f"role not found: {role_name}")
    return int(role_id)


def user_id_by_username(db_path: Path, username: str) -> int:
    user_id = db_scalar(db_path, "SELECT id FROM admin_users WHERE username = ?", (username,))
    if user_id is None:
        raise RuntimeError(f"user not found: {username}")
    return int(user_id)


def create_acceptance_admin_user(
    main_module: Any,
    db_path: Path,
    username: str,
    display_name: str,
    role_name: str = "店长",
    allow_login: int = 1,
    is_active: int = 1,
    is_assignable: int = 0,
) -> int:
    role_id = role_id_by_name(db_path, role_name)
    timestamp = "2026-07-08 10:00:00"
    with sqlite3.connect(db_path) as connection:
        cursor = connection.execute(
            """
            INSERT INTO admin_users (
                username, display_name, password_hash, role_id, employee_id,
                allow_login, participate_schedule, is_active, is_assignable,
                data_scope, store_names, created_at, updated_at, created_by, updated_by
            )
            VALUES (?, ?, ?, ?, NULL, ?, 0, ?, ?, 'all', '', ?, ?, 'acceptance', 'acceptance')
            """,
            (
                username,
                display_name,
                main_module.hash_password("acceptance-secret"),
                role_id,
                int(allow_login),
                int(is_active),
                int(is_assignable),
                timestamp,
                timestamp,
            ),
        )
        return int(cursor.lastrowid)


def run_isolated_write(report_path: Path | None) -> Path:
    before_db = file_fingerprint(REAL_DB_PATH)
    issues: list[dict[str, str]] = []
    local_commit = git_commit(short=False)
    isolated_result: dict[str, Any] = {}

    def note_check(key: str, ok: bool, message: str) -> None:
        isolated_result[key] = "通过" if ok else "失败"
        if not ok:
            issues.append({"severity": "blocker", "area": "isolated-write", "message": message})

    with tempfile.TemporaryDirectory(prefix="store_request_acceptance_", ignore_cleanup_errors=True) as tmp:
        tmp_dir = Path(tmp)
        db_path = tmp_dir / "tickets.db"
        upload_dir = tmp_dir / "uploads"
        config_dir = tmp_dir / "config"
        write_minimal_config(config_dir)
        os.environ["STORE_REQUEST_DB_PATH"] = str(db_path)
        os.environ["STORE_REQUEST_UPLOAD_DIR"] = str(upload_dir)
        os.environ["STORE_REQUEST_CONFIG_DIR"] = str(config_dir)
        os.environ["ADMIN_USERS"] = "acceptance-admin:acceptance-secret"
        os.environ["ADMIN_USERNAME"] = "acceptance-admin"
        os.environ["ADMIN_PASSWORD"] = "acceptance-secret"
        os.environ["SESSION_SECRET"] = "acceptance-session-secret"
        os.environ["APP_ENV"] = "test"
        sys.path.insert(0, str(PROJECT_DIR))
        sys.modules.pop("main", None)
        main = importlib.import_module("main")

        from fastapi.testclient import TestClient

        with TestClient(main.app) as client:
            login = client.post(
                "/admin/login",
                data={"username": "acceptance-admin", "password": "acceptance-secret"},
                follow_redirects=False,
            )
            isolated_result["login_status"] = login.status_code
            if login.status_code != 303:
                issues.append({"severity": "blocker", "area": "isolated-login", "message": f"临时环境登录失败：{login.status_code}"})

            submit = client.post(
                "/submit",
                data={
                    "store_name": main.load_app_config().stores[0],
                    "submitter": "隔离验收",
                    "request_type": main.load_app_config().request_types[-1],
                    "urgency": main.load_app_config().urgency_levels[0],
                    "brand": "",
                    "product_name": "",
                    "sku_barcode": "",
                    "quantity": "",
                    "description": "isolated-write 临时工单，不写入真实数据库。",
                    "expected_finish_date": "",
                },
            )
            isolated_result["submit_without_product_status"] = submit.status_code
            if submit.status_code != 200:
                issues.append({"severity": "blocker", "area": "isolated-submit", "message": f"商品信息为空提交失败：{submit.status_code}"})

        employee_id = main.create_employee(
            "隔离验收员工",
            main.load_app_config().stores[0],
            "店员",
            "",
            "在职",
            main.load_app_config(),
            store_names=[main.load_app_config().stores[0]],
            primary_store_name=main.load_app_config().stores[0],
        )
        shifts = main.fetch_shift_types(active_only=True, store_names=[main.load_app_config().stores[0]], global_scope="all")
        if not shifts:
            raise RuntimeError("临时环境没有可用班次")
        schedule_result = main.bulk_upsert_schedules(
            main.load_app_config().stores[0],
            [employee_id],
            [(datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")],
            int(shifts[0]["id"]),
            "isolated acceptance",
            False,
            "acceptance-admin",
            200,
        )
        isolated_result["schedule_saved_count"] = schedule_result.get("saved_count")

        with TestClient(main.app) as client:
            login = client.post(
                "/admin/login",
                data={"username": "acceptance-admin", "password": "acceptance-secret"},
                follow_redirects=False,
            )
            note_check("delete_flow_login", login.status_code == 303, f"删除机制验证登录失败：{login.status_code}")

            employee_csrf = testclient_csrf_token(client, "/admin/employees")
            create_safe_employee = client.post(
                "/admin/employees",
                data={
                    "csrf_token": employee_csrf,
                    "employee_name": "隔离仅审计员工",
                    "primary_store_name": main.load_app_config().stores[0],
                    "store_names": [main.load_app_config().stores[0]],
                    "role": "店员",
                    "phone": "",
                    "status": "在职",
                    "show_in_employee_management": "1",
                    "participate_schedule": "1",
                },
                follow_redirects=False,
            )
            safe_employee_id = int(db_scalar(db_path, "SELECT id FROM employees WHERE employee_name = ?", ("隔离仅审计员工",)) or 0)
            audit_only_assessment = main.can_hard_delete_employee(safe_employee_id)
            note_check(
                "employee_safe_delete_allows_create_audit_only",
                create_safe_employee.status_code == 303
                and safe_employee_id > 0
                and audit_only_assessment.get("can_hard_delete") is True
                and "历史日志" not in str(audit_only_assessment.get("reason_text") or "")
                and "隔离仅审计员工" in client.get("/admin/employees?filter=safe_delete").text,
                "通过页面新建且只有创建审计日志的员工未被判定为可安全删除。",
            )
            safe_delete_employee = client.post(
                f"/admin/employees/{safe_employee_id}/safe-delete",
                data={"csrf_token": employee_csrf, "confirm_delete": "1", "source_scope": "active"},
                follow_redirects=False,
            )
            note_check(
                "employee_safe_delete_deleted",
                safe_delete_employee.status_code == 303 and db_row(db_path, "employees", safe_employee_id) is None,
                "无历史引用员工 safe-delete 未真实删除。",
            )

            mapped_employee_id = main.create_employee(
                "隔离仅门店映射员工",
                main.load_app_config().stores[0],
                "店员",
                "",
                "在职",
                main.load_app_config(),
                store_names=[main.load_app_config().stores[0]],
                primary_store_name=main.load_app_config().stores[0],
            )
            mapped_assessment = main.can_hard_delete_employee(mapped_employee_id)
            mapped_delete = client.post(
                f"/admin/employees/{mapped_employee_id}/safe-delete",
                data={"csrf_token": employee_csrf, "confirm_delete": "1", "source_scope": "active"},
                follow_redirects=False,
            )
            mapped_store_rows = int(
                db_scalar(db_path, "SELECT COUNT(*) FROM employee_store_map WHERE employee_id = ?", (mapped_employee_id,)) or 0
            )
            note_check(
                "employee_safe_delete_allows_store_map_only",
                mapped_assessment.get("can_hard_delete") is True
                and mapped_delete.status_code == 303
                and db_row(db_path, "employees", mapped_employee_id) is None
                and mapped_store_rows == 0,
                "只有 employee_store_map 的员工未能安全删除或映射未清理。",
            )

            account_from_employee_id = main.create_employee(
                "隔离开通账号员工",
                main.load_app_config().stores[0],
                "店员",
                "",
                "在职",
                main.load_app_config(),
                store_names=[main.load_app_config().stores[0]],
                primary_store_name=main.load_app_config().stores[0],
            )
            create_account_page = client.get("/admin/employees")
            create_account_csrf = testclient_csrf_token(client, "/admin/employees")
            create_account_from_employee = client.post(
                f"/admin/employees/{account_from_employee_id}/create-account",
                data={
                    "csrf_token": create_account_csrf,
                    "username": "accept-created-user",
                    "password": "acceptance-secret",
                    "password_confirm": "acceptance-secret",
                    "role_id": str(role_id_by_name(db_path, "店长")),
                    "data_scope": "stores",
                    "store_names": main.load_app_config().stores[0],
                    "allow_login": "1",
                    "is_active": "1",
                    "is_assignable": "1",
                },
                follow_redirects=False,
            )
            created_account_id = int(db_scalar(db_path, "SELECT id FROM admin_users WHERE username = ?", ("accept-created-user",)) or 0)
            created_account = db_row(db_path, "admin_users", created_account_id) or {}
            employee_after_create_account = db_row(db_path, "employees", account_from_employee_id) or {}
            note_check(
                "employee_create_account_from_employee_page",
                "开通后台账号" in create_account_page.text
                and create_account_from_employee.status_code == 303
                and created_account.get("employee_id") == account_from_employee_id
                and employee_after_create_account.get("user_id") == created_account_id
                and employee_after_create_account.get("participate_schedule") == 1
                and "accept-created-user" in client.get("/admin/account?filter=all").text,
                "人员管理开通后台账号未成功创建双向绑定账号或账号页不可见。",
            )

            scheduled_employee_id = main.create_employee(
                "隔离有排班员工",
                main.load_app_config().stores[0],
                "店员",
                "",
                "在职",
                main.load_app_config(),
                store_names=[main.load_app_config().stores[0]],
                primary_store_name=main.load_app_config().stores[0],
            )
            with sqlite3.connect(db_path) as connection:
                shift_id = connection.execute("SELECT id FROM shift_types ORDER BY id LIMIT 1").fetchone()[0]
                connection.execute(
                    """
                    INSERT INTO store_schedules (store_name, employee_id, schedule_date, shift_type_id, created_by, created_at, updated_at)
                    VALUES (?, ?, '2026-07-08', ?, 'acceptance', '2026-07-08 10:00:00', '2026-07-08 10:00:00')
                    """,
                    (main.load_app_config().stores[0], scheduled_employee_id, shift_id),
                )
            blocked_employee_delete = client.post(
                f"/admin/employees/{scheduled_employee_id}/safe-delete",
                data={"csrf_token": employee_csrf, "confirm_delete": "1", "source_scope": "active"},
                follow_redirects=True,
            )
            note_check(
                "employee_safe_delete_blocks_history",
                blocked_employee_delete.status_code == 200
                and db_row(db_path, "employees", scheduled_employee_id) is not None
                and "存在历史排班" in blocked_employee_delete.text,
                "有历史排班员工 safe-delete 未被页面原因拒绝。",
            )

            archive_employee = client.post(
                f"/admin/employees/{scheduled_employee_id}/archive",
                data={"csrf_token": employee_csrf, "archive_reason": "isolated-write"},
                follow_redirects=False,
            )
            note_check(
                "employee_archive_hides_default_and_shows_archived",
                archive_employee.status_code == 303
                and "隔离有排班员工" not in client.get("/admin/employees").text
                and "隔离有排班员工" in client.get("/admin/employees?filter=archived").text,
                "员工归档后默认页/归档筛选页状态不正确。",
            )
            employee_csrf = testclient_csrf_token(client, "/admin/employees?filter=archived")
            unarchive_employee = client.post(
                f"/admin/employees/{scheduled_employee_id}/unarchive",
                data={"csrf_token": employee_csrf},
                follow_redirects=False,
            )
            note_check(
                "employee_unarchive_restores_default",
                unarchive_employee.status_code == 303 and "隔离有排班员工" in client.get("/admin/employees").text,
                "员工恢复归档后默认页未显示。",
            )
            employee_csrf = testclient_csrf_token(client, "/admin/employees")
            hide_employee = client.post(
                f"/admin/employees/{scheduled_employee_id}/hide",
                data={"csrf_token": employee_csrf},
                follow_redirects=False,
            )
            note_check(
                "employee_hide_hides_default_and_shows_hidden",
                hide_employee.status_code == 303
                and "隔离有排班员工" not in client.get("/admin/employees").text
                and "隔离有排班员工" in client.get("/admin/employees?filter=hidden").text,
                "员工隐藏后默认页/隐藏筛选页状态不正确。",
            )
            employee_csrf = testclient_csrf_token(client, "/admin/employees?filter=hidden")
            show_employee = client.post(
                f"/admin/employees/{scheduled_employee_id}/show",
                data={"csrf_token": employee_csrf},
                follow_redirects=False,
            )
            note_check(
                "employee_show_restores_default",
                show_employee.status_code == 303 and "隔离有排班员工" in client.get("/admin/employees").text,
                "员工显示后默认页未恢复。",
            )

            account_csrf = testclient_csrf_token(client, "/admin/account?filter=all")
            account_employee_id = main.create_employee(
                "隔离账号员工",
                main.load_app_config().stores[0],
                "店员",
                "",
                "在职",
                main.load_app_config(),
                store_names=[main.load_app_config().stores[0]],
                primary_store_name=main.load_app_config().stores[0],
            )
            deletable_user_id = create_acceptance_admin_user(main, db_path, "accept-delete-user", "隔离可删账号")
            with sqlite3.connect(db_path) as connection:
                connection.execute("UPDATE admin_users SET employee_id = ? WHERE id = ?", (account_employee_id, deletable_user_id))
                connection.execute("UPDATE employees SET user_id = ?, allow_login = 1 WHERE id = ?", (deletable_user_id, account_employee_id))
            delete_account = client.post(
                f"/admin/accounts/{deletable_user_id}/safe-delete",
                data={"csrf_token": account_csrf, "confirm_delete": "1"},
                follow_redirects=False,
            )
            employee_after_account_delete = db_row(db_path, "employees", account_employee_id)
            note_check(
                "account_safe_delete_deletes_user_and_unlinks_employee",
                delete_account.status_code == 303
                and db_row(db_path, "admin_users", deletable_user_id) is None
                and employee_after_account_delete is not None
                and employee_after_account_delete.get("user_id") is None,
                "无历史账号 safe-delete 未删除账号或未清空 employee.user_id。",
            )

            history_user_id = create_acceptance_admin_user(main, db_path, "accept-history-user", "隔离历史账号")
            with sqlite3.connect(db_path) as connection:
                connection.execute(
                    """
                    INSERT INTO admin_login_logs (username, success, ip_address, user_agent, message, created_at)
                    VALUES ('accept-history-user', 1, '127.0.0.1', 'acceptance', '历史登录', '2026-07-08 10:00:00')
                    """
                )
            account_csrf = testclient_csrf_token(client, "/admin/account?filter=all")
            blocked_account_delete = client.post(
                f"/admin/accounts/{history_user_id}/safe-delete",
                data={"csrf_token": account_csrf, "confirm_delete": "1"},
                follow_redirects=True,
            )
            note_check(
                "account_safe_delete_blocks_login_history",
                blocked_account_delete.status_code == 400
                and db_row(db_path, "admin_users", history_user_id) is not None
                and ("历史登录" in blocked_account_delete.text or "不能永久删除" in blocked_account_delete.text),
                "有登录日志账号 safe-delete 未被页面原因拒绝。",
            )
            account_csrf = testclient_csrf_token(client, "/admin/account?filter=all")
            disabled_account = client.post(
                f"/admin/accounts/{history_user_id}/disable-login",
                data={"csrf_token": account_csrf},
                follow_redirects=False,
            )
            history_user = db_row(db_path, "admin_users", history_user_id) or {}
            note_check(
                "account_disable_login_flags",
                disabled_account.status_code == 303
                and history_user.get("allow_login") == 0
                and history_user.get("is_active") == 0
                and history_user.get("is_assignable") == 0,
                "关闭后台账号后 allow_login/is_active/is_assignable 未同步为 0。",
            )
            account_csrf = testclient_csrf_token(client, "/admin/account?filter=all")
            enabled_account = client.post(
                f"/admin/accounts/{history_user_id}/enable-login",
                data={"csrf_token": account_csrf},
                follow_redirects=False,
            )
            history_user = db_row(db_path, "admin_users", history_user_id) or {}
            note_check(
                "account_enable_login_flags",
                enabled_account.status_code == 303
                and history_user.get("allow_login") == 1
                and history_user.get("is_active") == 1,
                "启用后台账号后 allow_login/is_active 未恢复为 1。",
            )
            account_csrf = testclient_csrf_token(client, "/admin/account?filter=all")
            hidden_account = client.post(
                f"/admin/accounts/{history_user_id}/hide",
                data={"csrf_token": account_csrf},
                follow_redirects=False,
            )
            history_user = db_row(db_path, "admin_users", history_user_id) or {}
            note_check(
                "account_hide_hides_default_and_shows_hidden",
                hidden_account.status_code == 303
                and history_user.get("show_in_account_management") == 0
                and "隔离历史账号" not in client.get("/admin/account").text
                and "隔离历史账号" in client.get("/admin/account?filter=hidden").text,
                "隐藏账号后默认页/隐藏筛选页状态不正确。",
            )
            account_csrf = testclient_csrf_token(client, "/admin/account?filter=hidden")
            shown_account = client.post(
                f"/admin/accounts/{history_user_id}/show",
                data={"csrf_token": account_csrf},
                follow_redirects=False,
            )
            history_user = db_row(db_path, "admin_users", history_user_id) or {}
            note_check(
                "account_show_restores_default",
                shown_account.status_code == 303
                and history_user.get("show_in_account_management") == 1
                and "隔离历史账号" in client.get("/admin/account").text,
                "显示账号后默认页未恢复。",
            )

            admin_id = user_id_by_username(db_path, "acceptance-admin")
            account_csrf = testclient_csrf_token(client, "/admin/account?filter=all")
            last_admin_disabled = client.post(
                f"/admin/accounts/{admin_id}/disable-login",
                data={"csrf_token": account_csrf},
                follow_redirects=True,
            )
            admin_user = db_row(db_path, "admin_users", admin_id) or {}
            note_check(
                "last_active_system_admin_protected",
                last_admin_disabled.status_code == 400
                and admin_user.get("allow_login") == 1
                and admin_user.get("is_active") == 1
                and ("不能停用最后一个系统管理员" in last_admin_disabled.text or "不能停用当前登录账号" in last_admin_disabled.text),
                "最后一个 active 系统管理员关闭保护未生效。",
            )

            with sqlite3.connect(db_path) as connection:
                operation_actions = {
                    row[0]
                    for row in connection.execute("SELECT action FROM admin_operation_logs").fetchall()
                }
            required_actions = {
                "employee.hard_delete",
                "employee.archive",
                "employee.unarchive",
                "employee.hide",
                "employee.show",
                "employee.create_account",
                "account.hard_delete",
                "account.disable",
                "account.enable",
                "account.hide",
                "account.show",
            }
            note_check(
                "delete_archive_hide_show_audit_logs",
                required_actions.issubset(operation_actions),
                f"删除/归档/隐藏/显示审计日志缺失：{sorted(required_actions - operation_actions)}",
            )

        with sqlite3.connect(db_path) as connection:
            isolated_result["temp_ticket_count"] = connection.execute("SELECT COUNT(*) FROM tickets").fetchone()[0]
            isolated_result["temp_employee_count"] = connection.execute("SELECT COUNT(*) FROM employees").fetchone()[0]
            isolated_result["temp_schedule_count"] = connection.execute("SELECT COUNT(*) FROM store_schedules").fetchone()[0]
            isolated_result["temp_login_log_count"] = connection.execute("SELECT COUNT(*) FROM admin_login_logs").fetchone()[0]
        isolated_result["temp_db_path"] = str(db_path)
        isolated_result["temp_upload_dir"] = str(upload_dir)
        isolated_result["temp_config_dir"] = str(config_dir)

    after_db = file_fingerprint(REAL_DB_PATH)
    data = {
        "base_url": "TestClient isolated",
        "local_commit": local_commit,
        "db_fingerprint_unchanged": before_db == after_db,
        "isolated_write": isolated_result,
        "issues": issues,
    }
    if before_db != after_db:
        data["issues"].append({"severity": "risk", "area": "isolated-write", "message": "isolated-write 前后真实 DB 文件指纹发生变化。"})
    return write_report("isolated-write", data, report_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="Non-browser acceptance smoke checks for Store Request Tool.")
    parser.add_argument("--mode", choices=["read-only", "isolated-write"], default="read-only")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL)
    parser.add_argument("--report-path", default="")
    args = parser.parse_args()

    report_path = Path(args.report_path).resolve() if args.report_path else None
    if args.mode == "read-only":
        path = run_read_only(args.base_url.rstrip("/"), report_path)
    else:
        path = run_isolated_write(report_path)
    try:
        display_path = path.relative_to(PROJECT_DIR)
    except ValueError:
        display_path = path
    print(f"REPORT_PATH={display_path.as_posix()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
