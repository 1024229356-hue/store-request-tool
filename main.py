import base64
import hashlib
import hmac
import json
import os
import secrets
import sqlite3
import uuid
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
import re
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import quote, urlencode, urlsplit

from fastapi import Depends, FastAPI, File, Form, HTTPException, Query, Request, UploadFile, status as http_status
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, UnidentifiedImageError


BASE_DIR = Path(__file__).resolve().parent
CONFIG_DIR = BASE_DIR / "config"
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"
DEFAULT_DATA_DIR = BASE_DIR / "data"
DEFAULT_UPLOAD_DIR = BASE_DIR / "uploads"
ENV_FILE = BASE_DIR / ".env"
SESSION_COOKIE_NAME = "admin_session"
DEFAULT_SESSION_SECRET = "store-request-tool-local-dev-session-secret"
PLACEHOLDER_SESSION_SECRET = "change-this-to-a-random-long-secret"
DEFAULT_SESSION_MAX_AGE_HOURS = 12

DEFAULT_STORES = ["南京门东店", "南昌万寿宫店", "山城巷店", "东郊记忆店", "蟠龙天地店", "秀水街店", "湾里店", "下浩里店", "烟台山店"]
DEFAULT_REQUEST_TYPES = ["建单需求", "审单需求", "商品异常", "缺货需求", "新品需求", "系统问题", "其他"]
DEFAULT_URGENCY_LEVELS = ["普通", "加急", "当天必须处理"]
DEFAULT_STATUSES = ["待处理", "处理中", "待门店补充", "已完成", "已驳回"]
DEFAULT_BRANDS: List[str] = []
DEFAULT_HANDLERS = ["总部商品", "总部运营", "采购", "财务"]
DUE_STATUS_OPTIONS = ["已超时", "今日到期", "未到期", "未设置", "超时完成", "按时完成"]
DEFAULT_SYSTEM = {
    "app_name": "门店需求工单系统",
    "port": 8701,
    "max_image_mb": 10,
    "allowed_image_extensions": ["jpg", "jpeg", "png", "webp"],
    "default_status": "待处理",
    "excel_filename_prefix": "门店需求工单",
    "page_size": 50,
    "max_image_count": 5,
    "max_total_upload_mb": 30,
    "allowed_file_extensions": ["pdf", "doc", "docx", "xls", "xlsx", "csv", "txt", "zip", "rar"],
    "max_file_mb": 20,
    "max_file_count": 5,
    "max_total_file_upload_mb": 50,
    "store_query_default_days": 30,
    "store_query_page_size": 20,
    "supplement_status_after_store_update": "待处理",
}
COMPLETED_STATUS = "已完成"
BLOCKED_FILE_EXTENSIONS = {"exe", "bat", "cmd", "js", "py", "sh", "php", "jar", "msi"}
IMAGE_FORMAT_EXTENSIONS = {
    "JPEG": {"jpg", "jpeg"},
    "PNG": {"png"},
    "WEBP": {"webp"},
}


@dataclass(frozen=True)
class AppConfig:
    stores: List[str]
    request_types: List[str]
    urgency_levels: List[str]
    statuses: List[str]
    brands: List[str]
    handlers: List[str]
    app_name: str
    port: int
    max_image_mb: int
    allowed_image_extensions: List[str]
    default_status: str
    excel_filename_prefix: str
    page_size: int
    max_image_count: int
    max_total_upload_mb: int
    allowed_file_extensions: List[str]
    max_file_mb: int
    max_file_count: int
    max_total_file_upload_mb: int
    store_query_default_days: int
    store_query_page_size: int
    supplement_status_after_store_update: str

    @property
    def max_image_bytes(self) -> int:
        return self.max_image_mb * 1024 * 1024

    @property
    def max_total_upload_bytes(self) -> int:
        return self.max_total_upload_mb * 1024 * 1024

    @property
    def max_file_bytes(self) -> int:
        return self.max_file_mb * 1024 * 1024

    @property
    def max_total_file_upload_bytes(self) -> int:
        return self.max_total_file_upload_mb * 1024 * 1024


@dataclass(frozen=True)
class PreparedFile:
    original_filename: str
    file_ext: str
    file_size: int
    content: bytes

EXCEL_HEADERS = [
    "工单号",
    "提交时间",
    "门店",
    "提报人",
    "需求类型",
    "紧急程度",
    "品牌",
    "商品名称",
    "规格条码",
    "数量",
    "问题说明",
    "图片路径",
    "文件名称",
    "文件路径",
    "期望完成时间",
    "当前状态",
    "处理人",
    "处理备注",
    "完成时间",
    "最后更新时间",
    "处理时长小时",
    "是否超时",
    "时效状态",
]


def load_env_file() -> None:
    if not ENV_FILE.exists():
        return
    for raw_line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        if not key or key in os.environ:
            continue
        os.environ[key] = value.strip().strip("\"'")


def parse_admin_users(raw_users: str) -> List[Tuple[str, str]]:
    users: List[Tuple[str, str]] = []
    for raw_item in raw_users.split(","):
        username, separator, password = raw_item.partition(":")
        if not separator:
            continue
        username = username.strip()
        password = password.strip()
        if username and password:
            users.append((username, password))
    return users


def get_admin_credentials() -> List[Tuple[str, str]]:
    admin_users = os.environ.get("ADMIN_USERS")
    if admin_users is not None:
        users = parse_admin_users(admin_users)
        if users:
            return users
        raise HTTPException(status_code=503, detail="Admin credentials are not configured.")

    username = os.environ.get("ADMIN_USERNAME", "").strip()
    password = os.environ.get("ADMIN_PASSWORD", "")
    if not username or not password:
        raise HTTPException(status_code=503, detail="Admin credentials are not configured.")
    return [(username, password)]


def authenticate_admin(username: str, password: str) -> Optional[str]:
    input_username = username.strip()
    for expected_username, expected_password in get_admin_credentials():
        username_ok = secrets.compare_digest(input_username, expected_username)
        password_ok = secrets.compare_digest(password, expected_password)
        if username_ok and password_ok:
            return expected_username
    return None


def admin_username_exists(username: str) -> bool:
    input_username = username.strip()
    for expected_username, _expected_password in get_admin_credentials():
        if secrets.compare_digest(input_username, expected_username):
            return True
    return False


def get_session_secret() -> str:
    return os.environ.get("SESSION_SECRET", DEFAULT_SESSION_SECRET)


def get_app_env() -> str:
    return os.environ.get("APP_ENV", "development").strip().lower() or "development"


def get_session_max_age_seconds() -> int:
    raw_value = os.environ.get("SESSION_MAX_AGE_HOURS", str(DEFAULT_SESSION_MAX_AGE_HOURS)).strip()
    try:
        hours = float(raw_value)
    except ValueError:
        hours = DEFAULT_SESSION_MAX_AGE_HOURS
    if hours <= 0:
        hours = DEFAULT_SESSION_MAX_AGE_HOURS
    return int(hours * 3600)


def session_cookie_secure() -> bool:
    return os.environ.get("SESSION_COOKIE_SECURE", "false").strip().lower() in {"1", "true", "yes", "on"}


def warn_if_insecure_production_session() -> None:
    if get_app_env() == "production" and get_session_secret() in {DEFAULT_SESSION_SECRET, PLACEHOLDER_SESSION_SECRET}:
        print(
            "WARNING: APP_ENV=production but SESSION_SECRET is still the default or example value. "
            "Set a random long SESSION_SECRET before exposing the admin system.",
            flush=True,
        )


def base64url_encode(raw: bytes) -> str:
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def base64url_decode(raw: str) -> bytes:
    padding = "=" * (-len(raw) % 4)
    return base64.urlsafe_b64decode((raw + padding).encode("ascii"))


def sign_session_payload(payload: str) -> str:
    signature = hmac.new(get_session_secret().encode("utf-8"), payload.encode("ascii"), hashlib.sha256).digest()
    return base64url_encode(signature)


def create_admin_session(
    username: str,
    issued_at: Optional[int] = None,
    max_age_seconds: Optional[int] = None,
) -> str:
    issued_at_value = int(issued_at if issued_at is not None else datetime.now().timestamp())
    max_age_value = int(max_age_seconds if max_age_seconds is not None else get_session_max_age_seconds())
    payload = base64url_encode(
        json.dumps(
            {
                "username": username,
                "issued_at": issued_at_value,
                "expires_at": issued_at_value + max_age_value,
                "csrf_token": secrets.token_urlsafe(32),
            },
            ensure_ascii=False,
            separators=(",", ":"),
        ).encode("utf-8")
    )
    return f"{payload}.{sign_session_payload(payload)}"


def read_admin_session_data(token: str) -> Optional[Dict[str, Any]]:
    if not token or "." not in token:
        return None
    payload, signature = token.rsplit(".", 1)
    expected_signature = sign_session_payload(payload)
    if not secrets.compare_digest(signature, expected_signature):
        return None
    try:
        data = json.loads(base64url_decode(payload).decode("utf-8"))
    except (ValueError, json.JSONDecodeError):
        return None
    username = str(data.get("username") or "").strip()
    expires_at = data.get("expires_at")
    if not isinstance(expires_at, (int, float)) or datetime.now().timestamp() >= float(expires_at):
        return None
    csrf_token = str(data.get("csrf_token") or "").strip()
    if username and csrf_token and admin_username_exists(username):
        return data
    return None


def read_admin_session(token: str) -> Optional[str]:
    data = read_admin_session_data(token)
    if not data:
        return None
    return str(data.get("username") or "").strip() or None


def current_admin_username(request: Request) -> Optional[str]:
    return read_admin_session(request.cookies.get(SESSION_COOKIE_NAME, ""))


def current_csrf_token(request: Request) -> str:
    data = read_admin_session_data(request.cookies.get(SESSION_COOKIE_NAME, ""))
    if not data:
        return ""
    return str(data.get("csrf_token") or "")


def require_admin_csrf(request: Request, csrf_token: str) -> None:
    expected_token = current_csrf_token(request)
    if not expected_token or not secrets.compare_digest(csrf_token or "", expected_token):
        raise HTTPException(status_code=http_status.HTTP_403_FORBIDDEN, detail="CSRF token invalid.")


def safe_admin_return_url(value: str) -> str:
    value = (value or "").strip()
    if not value:
        return "/admin"
    parsed = urlsplit(value)
    if parsed.scheme or parsed.netloc or not value.startswith("/admin") or value.startswith("//"):
        return "/admin"
    return value


def safe_query_return_url(value: str) -> str:
    value = (value or "").strip()
    if not value:
        return "/query"
    parsed = urlsplit(value)
    if parsed.scheme or parsed.netloc or value.startswith("//"):
        return "/query"
    if parsed.path != "/query" and not parsed.path.startswith("/query/"):
        return "/query"
    return value


def request_path_with_query(request: Request) -> str:
    return request.url.path + (f"?{request.url.query}" if request.url.query else "")


def build_ticket_detail_url(ticket_id: int, return_url: str = "", **flags: str) -> str:
    params = {key: value for key, value in flags.items() if value}
    safe_return_url = safe_admin_return_url(return_url)
    if safe_return_url != "/admin":
        params["return_url"] = safe_return_url
    query = urlencode(params)
    return f"/admin/ticket/{ticket_id}" + (f"?{query}" if query else "")


def build_store_query_url(store_name: str) -> str:
    clean_store_name = store_name.strip()
    if not clean_store_name:
        return "/query"
    return "/query?" + urlencode({"store_name": clean_store_name})


def store_query_list_return_url(store_name: str, return_url: str = "") -> str:
    if not return_url.strip():
        return build_store_query_url(store_name)
    safe_return_url = safe_query_return_url(return_url)
    if urlsplit(safe_return_url).path == "/query":
        return safe_return_url
    return build_store_query_url(store_name)


def build_store_ticket_detail_url(ticket_id: int, store_name: str, return_url: str = "") -> str:
    params = {"store_name": store_name.strip()}
    safe_return_url = safe_query_return_url(return_url)
    if safe_return_url != "/query":
        params["return_url"] = safe_return_url
    return f"/query/ticket/{ticket_id}?" + urlencode(params)


def build_store_ticket_supplement_url(ticket_id: int, store_name: str, return_url: str = "") -> str:
    params = {"store_name": store_name.strip()}
    safe_return_url = safe_query_return_url(return_url)
    if safe_return_url != "/query":
        params["return_url"] = safe_return_url
    return f"/query/ticket/{ticket_id}/supplement?" + urlencode(params)


def login_redirect_location(request: Request) -> str:
    next_url = safe_admin_return_url(request_path_with_query(request))
    return "/admin/login?" + urlencode({"next": next_url})


def should_redirect_to_login(request: Request) -> bool:
    path = request.url.path
    if path.startswith("/admin/uploads") or path.startswith("/admin/files") or path.startswith("/admin/export"):
        return False
    return (
        path == "/admin"
        or path == "/admin/dashboard"
        or path == "/admin/settings"
        or path == "/admin/account"
        or path == "/admin/system"
        or path.startswith("/admin/ticket")
    )


def require_admin(request: Request) -> str:
    username = current_admin_username(request)
    if username:
        return username
    if should_redirect_to_login(request):
        raise HTTPException(status_code=303, detail="Login required.", headers={"Location": login_redirect_location(request)})
    raise HTTPException(
        status_code=http_status.HTTP_401_UNAUTHORIZED,
        detail="Login required.",
    )


def get_db_path() -> Path:
    return Path(os.environ.get("STORE_REQUEST_DB_PATH", DEFAULT_DATA_DIR / "tickets.db"))


def get_upload_dir() -> Path:
    return Path(os.environ.get("STORE_REQUEST_UPLOAD_DIR", DEFAULT_UPLOAD_DIR))


def get_config_dir() -> Path:
    return Path(os.environ.get("STORE_REQUEST_CONFIG_DIR", CONFIG_DIR))


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def ensure_directories() -> None:
    get_db_path().parent.mkdir(parents=True, exist_ok=True)
    get_upload_dir().mkdir(parents=True, exist_ok=True)
    get_config_dir().mkdir(parents=True, exist_ok=True)
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    STATIC_DIR.mkdir(parents=True, exist_ok=True)


def get_connection() -> sqlite3.Connection:
    connection = sqlite3.connect(get_db_path())
    connection.row_factory = sqlite3.Row
    return connection


def table_exists(connection: sqlite3.Connection, table_name: str) -> bool:
    row = connection.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
        (table_name,),
    ).fetchone()
    return row is not None


def column_exists(connection: sqlite3.Connection, table_name: str, column_name: str) -> bool:
    if not table_exists(connection, table_name):
        return False
    rows = connection.execute(f"PRAGMA table_info({table_name})").fetchall()
    return any(str(row["name"]) == column_name for row in rows)


def add_column_if_missing(
    connection: sqlite3.Connection,
    table_name: str,
    column_name: str,
    column_definition: str,
) -> None:
    if not column_exists(connection, table_name, column_name):
        connection.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_definition}")


MULTI_VALUE_SPLIT_RE = re.compile(r"[,\uFF0C\u3001;\uFF1B]+")


def unique_clean_values(values: Iterable[object]) -> List[str]:
    cleaned: List[str] = []
    seen: set[str] = set()
    for value in values:
        text = str(value or "").strip()
        if not text or text in seen:
            continue
        cleaned.append(text)
        seen.add(text)
    return cleaned


def split_multi_value_text(value: object) -> List[str]:
    return unique_clean_values(MULTI_VALUE_SPLIT_RE.split(str(value or "")))


def join_display_values(values: Iterable[object]) -> str:
    return "、".join(unique_clean_values(values))


def normalize_store_names(raw_store_names: Optional[List[str]], legacy_store_name: str = "") -> List[str]:
    candidates: List[object] = []
    for item in raw_store_names or []:
        candidates.extend(split_multi_value_text(item))
    if not candidates and legacy_store_name.strip():
        candidates.extend(split_multi_value_text(legacy_store_name))
    return unique_clean_values(candidates)


def normalize_brand_names(
    raw_brands: Optional[List[str]],
    brand_extra: str = "",
    legacy_brand: str = "",
) -> List[str]:
    candidates: List[object] = []
    for item in raw_brands or []:
        candidates.extend(split_multi_value_text(item))
    candidates.extend(split_multi_value_text(brand_extra))
    if not candidates and legacy_brand.strip():
        candidates.extend(split_multi_value_text(legacy_brand))
    return unique_clean_values(candidates)


def backfill_ticket_relations(connection: sqlite3.Connection) -> None:
    timestamp = now_text()
    rows = connection.execute("SELECT id, store_name, brand, created_at FROM tickets").fetchall()
    for row in rows:
        ticket_id = int(row["id"])
        created_at = str(row["created_at"] or timestamp)
        for store_name in split_multi_value_text(row["store_name"]):
            connection.execute(
                """
                INSERT OR IGNORE INTO ticket_stores (ticket_id, store_name, created_at)
                VALUES (?, ?, ?)
                """,
                (ticket_id, store_name, created_at),
            )
        for brand in split_multi_value_text(row["brand"]):
            connection.execute(
                """
                INSERT OR IGNORE INTO ticket_brands (ticket_id, brand, created_at)
                VALUES (?, ?, ?)
                """,
                (ticket_id, brand, created_at),
            )


def init_db() -> None:
    ensure_directories()
    with get_connection() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS tickets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_no TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                store_name TEXT NOT NULL,
                submitter TEXT NOT NULL,
                request_type TEXT NOT NULL,
                urgency TEXT NOT NULL,
                brand TEXT,
                product_name TEXT,
                sku_barcode TEXT,
                quantity INTEGER,
                description TEXT NOT NULL,
                expected_finish_date TEXT,
                status TEXT NOT NULL,
                assigned_to TEXT,
                handler_note TEXT,
                closed_at TEXT
            )
            """
        )
        add_column_if_missing(connection, "tickets", "assigned_to", "TEXT")
        add_column_if_missing(connection, "tickets", "closed_at", "TEXT")
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS ticket_stores (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_id INTEGER NOT NULL,
                store_name TEXT NOT NULL,
                created_at TEXT NOT NULL,
                UNIQUE(ticket_id, store_name),
                FOREIGN KEY (ticket_id) REFERENCES tickets(id) ON DELETE CASCADE
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS ticket_brands (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_id INTEGER NOT NULL,
                brand TEXT NOT NULL,
                created_at TEXT NOT NULL,
                UNIQUE(ticket_id, brand),
                FOREIGN KEY (ticket_id) REFERENCES tickets(id) ON DELETE CASCADE
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS ticket_images (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_id INTEGER NOT NULL,
                image_path TEXT NOT NULL,
                uploaded_at TEXT NOT NULL,
                source TEXT,
                uploaded_by TEXT,
                supplement_id INTEGER,
                FOREIGN KEY (ticket_id) REFERENCES tickets(id) ON DELETE CASCADE
            )
            """
        )
        add_column_if_missing(connection, "ticket_images", "source", "TEXT")
        add_column_if_missing(connection, "ticket_images", "uploaded_by", "TEXT")
        add_column_if_missing(connection, "ticket_images", "supplement_id", "INTEGER")
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS ticket_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_id INTEGER NOT NULL,
                original_filename TEXT NOT NULL,
                stored_filename TEXT NOT NULL,
                file_path TEXT NOT NULL,
                file_ext TEXT NOT NULL,
                file_size INTEGER NOT NULL,
                uploaded_at TEXT NOT NULL,
                source TEXT,
                uploaded_by TEXT,
                supplement_id INTEGER,
                FOREIGN KEY (ticket_id) REFERENCES tickets(id) ON DELETE CASCADE
            )
            """
        )
        add_column_if_missing(connection, "ticket_files", "source", "TEXT")
        add_column_if_missing(connection, "ticket_files", "uploaded_by", "TEXT")
        add_column_if_missing(connection, "ticket_files", "supplement_id", "INTEGER")
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS ticket_supplements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_id INTEGER NOT NULL,
                store_name TEXT NOT NULL,
                submitter TEXT NOT NULL,
                note TEXT,
                image_count INTEGER DEFAULT 0,
                file_count INTEGER DEFAULT 0,
                created_at TEXT NOT NULL,
                FOREIGN KEY (ticket_id) REFERENCES tickets(id) ON DELETE CASCADE
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS ticket_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_id INTEGER NOT NULL,
                action TEXT NOT NULL,
                old_status TEXT,
                new_status TEXT,
                old_assigned_to TEXT,
                new_assigned_to TEXT,
                note TEXT,
                operator TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY (ticket_id) REFERENCES tickets(id) ON DELETE CASCADE
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS notification_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                event_type TEXT NOT NULL,
                ticket_id INTEGER,
                ticket_no TEXT,
                store_name TEXT,
                title TEXT NOT NULL,
                content TEXT,
                severity TEXT NOT NULL DEFAULT 'info',
                created_by TEXT,
                created_at TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS notification_reads (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                event_id INTEGER NOT NULL,
                username TEXT NOT NULL,
                read_at TEXT NOT NULL,
                UNIQUE(event_id, username),
                FOREIGN KEY (event_id) REFERENCES notification_events(id) ON DELETE CASCADE
            )
            """
        )
        connection.execute("CREATE INDEX IF NOT EXISTS idx_tickets_created_at ON tickets(created_at)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_tickets_status ON tickets(status)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_tickets_assigned_to ON tickets(assigned_to)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_ticket_stores_store_name ON ticket_stores(store_name)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_ticket_stores_ticket_id ON ticket_stores(ticket_id)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_ticket_brands_brand ON ticket_brands(brand)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_ticket_brands_ticket_id ON ticket_brands(ticket_id)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_ticket_images_ticket_id ON ticket_images(ticket_id)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_ticket_files_ticket_id ON ticket_files(ticket_id)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_ticket_logs_ticket_id ON ticket_logs(ticket_id)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_ticket_supplements_ticket_id ON ticket_supplements(ticket_id)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_notification_events_created_at ON notification_events(created_at)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_notification_events_ticket_id ON notification_events(ticket_id)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_notification_reads_username ON notification_reads(username)")
        backfill_ticket_relations(connection)


def clean_string_list(value: object, default: List[str], allow_empty: bool = False) -> List[str]:
    if not isinstance(value, list):
        return list(default)
    clean_values = [str(item).strip() for item in value if str(item).strip()]
    if clean_values or allow_empty:
        return clean_values
    return list(default)


def load_json_file(filename: str, default: object) -> object:
    try:
        with (get_config_dir() / filename).open("r", encoding="utf-8") as file:
            return json.load(file)
    except (OSError, json.JSONDecodeError):
        return default


def load_list_config(filename: str, default: List[str], allow_empty: bool = False) -> List[str]:
    return clean_string_list(load_json_file(filename, default), default, allow_empty=allow_empty)


def positive_int_config(system: Dict[str, object], key: str) -> int:
    value = system.get(key)
    default_value = int(DEFAULT_SYSTEM[key])
    if isinstance(value, bool):
        return default_value
    if isinstance(value, int) and value > 0:
        return value
    return default_value


def normalized_extensions(value: object, default: List[str], blocked: Optional[set[str]] = None) -> List[str]:
    extensions = clean_string_list(value, default)
    blocked = blocked or set()
    normalized: List[str] = []
    for extension in extensions:
        clean_extension = extension.lower().lstrip(".")
        if clean_extension and clean_extension not in blocked and clean_extension not in normalized:
            normalized.append(clean_extension)
    if normalized:
        return normalized
    return [extension for extension in default if extension not in blocked]


def load_system_config(statuses: List[str]) -> Dict[str, object]:
    raw_system = load_json_file("system.json", DEFAULT_SYSTEM)
    system = dict(DEFAULT_SYSTEM)
    if isinstance(raw_system, dict):
        system.update(raw_system)

    for key in (
        "max_image_mb",
        "page_size",
        "max_image_count",
        "max_total_upload_mb",
        "max_file_mb",
        "max_file_count",
        "max_total_file_upload_mb",
        "store_query_default_days",
        "store_query_page_size",
    ):
        system[key] = positive_int_config(system, key)

    system["allowed_image_extensions"] = normalized_extensions(
        system.get("allowed_image_extensions"),
        list(DEFAULT_SYSTEM["allowed_image_extensions"]),
    )
    system["allowed_file_extensions"] = normalized_extensions(
        system.get("allowed_file_extensions"),
        list(DEFAULT_SYSTEM["allowed_file_extensions"]),
        BLOCKED_FILE_EXTENSIONS,
    )

    default_status = str(system.get("default_status", "")).strip()
    system["default_status"] = default_status if default_status in statuses else statuses[0]

    supplement_status = str(system.get("supplement_status_after_store_update", "")).strip()
    system["supplement_status_after_store_update"] = supplement_status if supplement_status in statuses else "待处理"

    for key in ("app_name", "excel_filename_prefix"):
        value = str(system.get(key, "")).strip()
        if value:
            system[key] = value
        else:
            system[key] = DEFAULT_SYSTEM[key]

    port = system.get("port")
    if not isinstance(port, int) or port <= 0:
        system["port"] = DEFAULT_SYSTEM["port"]

    return system


def load_app_config() -> AppConfig:
    statuses = load_list_config("statuses.json", DEFAULT_STATUSES)
    system = load_system_config(statuses)
    return AppConfig(
        stores=load_list_config("stores.json", DEFAULT_STORES),
        request_types=load_list_config("request_types.json", DEFAULT_REQUEST_TYPES),
        urgency_levels=load_list_config("urgency_levels.json", DEFAULT_URGENCY_LEVELS),
        statuses=statuses,
        brands=load_list_config("brands.json", DEFAULT_BRANDS, allow_empty=True),
        handlers=load_list_config("handlers.json", DEFAULT_HANDLERS, allow_empty=True),
        app_name=str(system["app_name"]),
        port=int(system["port"]),
        max_image_mb=int(system["max_image_mb"]),
        allowed_image_extensions=list(system["allowed_image_extensions"]),
        default_status=str(system["default_status"]),
        excel_filename_prefix=str(system["excel_filename_prefix"]),
        page_size=int(system["page_size"]),
        max_image_count=int(system["max_image_count"]),
        max_total_upload_mb=int(system["max_total_upload_mb"]),
        allowed_file_extensions=list(system["allowed_file_extensions"]),
        max_file_mb=int(system["max_file_mb"]),
        max_file_count=int(system["max_file_count"]),
        max_total_file_upload_mb=int(system["max_total_file_upload_mb"]),
        store_query_default_days=int(system["store_query_default_days"]),
        store_query_page_size=int(system["store_query_page_size"]),
        supplement_status_after_store_update=str(system["supplement_status_after_store_update"]),
    )


def load_stores() -> List[str]:
    return load_app_config().stores


REQUEST_RULE_FIELD_LABELS = {
    "brand": "品牌",
    "product_name": "商品名称",
    "sku_barcode": "规格条码",
    "quantity": "数量",
    "description": "问题说明",
    "expected_finish_date": "期望完成时间",
}


def load_request_type_rules() -> Dict[str, Dict[str, object]]:
    raw_rules = load_json_file("request_type_rules.json", {})
    if not isinstance(raw_rules, dict):
        return {}
    rules: Dict[str, Dict[str, object]] = {}
    for request_type, raw_rule in raw_rules.items():
        if not isinstance(raw_rule, dict):
            continue
        required_fields = [
            str(field).strip()
            for field in raw_rule.get("required_fields", [])
            if str(field).strip() in REQUEST_RULE_FIELD_LABELS
        ]
        rules[str(request_type)] = {
            "required_fields": required_fields,
            "require_image": bool(raw_rule.get("require_image", False)),
            "require_file": bool(raw_rule.get("require_file", False)),
            "require_any_attachment": bool(raw_rule.get("require_any_attachment", False)),
            "description_hint": str(raw_rule.get("description_hint") or "").strip(),
        }
    return rules


def compact_text(value: Optional[str], max_len: int = 36) -> str:
    text = (value or "").strip()
    if len(text) <= max_len:
        return text
    return text[:max_len] + "..."


def image_filename(image_path: str) -> str:
    return Path(str(image_path).replace("\\", "/")).name


def protected_upload_url(image_path: str) -> str:
    return f"/admin/uploads/{quote(image_filename(image_path))}"


def protected_file_url(file_id: object) -> str:
    return f"/admin/files/{file_id}"


def file_size_label(size: object) -> str:
    try:
        size_value = int(size)
    except (TypeError, ValueError):
        return ""
    if size_value >= 1024 * 1024:
        return f"{size_value / (1024 * 1024):.1f}MB"
    if size_value >= 1024:
        return f"{size_value / 1024:.1f}KB"
    return f"{size_value}B"


def safe_uploaded_name(filename: str) -> str:
    return Path(str(filename or "").replace("\\", "/")).name.strip()


def resolve_upload_path(filename: str) -> Optional[Path]:
    normalized = filename.replace("\\", "/")
    if not normalized or "/" in normalized or normalized in {".", ".."}:
        return None
    upload_root = get_upload_dir().resolve()
    target = (upload_root / normalized).resolve()
    try:
        target.relative_to(upload_root)
    except ValueError:
        return None
    return target


def resolve_upload_file(filename: str) -> Optional[Path]:
    target = resolve_upload_path(filename)
    if target is None:
        return None
    if not target.is_file():
        return None
    return target


def generate_ticket_no(connection: sqlite3.Connection) -> str:
    today = datetime.now().strftime("%Y%m%d")
    prefix = f"REQ-{today}-"
    row = connection.execute(
        "SELECT MAX(CAST(SUBSTR(ticket_no, 14) AS INTEGER)) AS max_no FROM tickets WHERE ticket_no LIKE ?",
        (prefix + "%",),
    ).fetchone()
    next_no = (row["max_no"] or 0) + 1
    return f"{prefix}{next_no:04d}"


def parse_datetime_text(value: object) -> Optional[datetime]:
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def processing_hours(ticket: Dict[str, object]) -> object:
    created_at = parse_datetime_text(ticket.get("created_at"))
    if not created_at:
        return ""
    finished_at = parse_datetime_text(ticket.get("closed_at")) or datetime.now()
    seconds = max((finished_at - created_at).total_seconds(), 0)
    return round(seconds / 3600, 2)


def overdue_text(ticket: Dict[str, object]) -> str:
    expected = parse_datetime_text(ticket.get("expected_finish_date"))
    if not expected:
        return ""
    closed_at = parse_datetime_text(ticket.get("closed_at"))
    compare_date = closed_at.date() if closed_at else datetime.now().date()
    return "是" if compare_date > expected.date() else "否"


def due_status_label(ticket: Dict[str, object]) -> str:
    expected = parse_datetime_text(ticket.get("expected_finish_date"))
    if not expected:
        return "未设置"
    expected_date = expected.date()
    status = str(ticket.get("status") or "")
    closed_at = parse_datetime_text(ticket.get("closed_at"))
    if status == COMPLETED_STATUS:
        closed_date = (closed_at or parse_datetime_text(ticket.get("updated_at")) or datetime.now()).date()
        return "超时完成" if closed_date > expected_date else "按时完成"
    today = datetime.now().date()
    if today > expected_date:
        return "已超时"
    if today == expected_date:
        return "今日到期"
    return "未到期"


def due_status_class(label: str) -> str:
    mapping = {
        "已超时": "overdue",
        "今日到期": "due-today",
        "未到期": "not-due",
        "未设置": "unset",
        "超时完成": "completed-late",
        "按时完成": "completed-on-time",
    }
    return mapping.get(label, "unset")


def annotate_ticket_runtime(ticket: Dict[str, object]) -> Dict[str, object]:
    ticket.setdefault("store_names", split_multi_value_text(ticket.get("store_name")))
    ticket.setdefault("brand_names", split_multi_value_text(ticket.get("brand")))
    label = due_status_label(ticket)
    ticket["due_status"] = label
    ticket["due_status_class"] = due_status_class(label)
    ticket["description_summary"] = compact_text(str(ticket.get("description") or ""))
    return ticket


def relation_map_for_tickets(ticket_ids: List[int], table_name: str, value_column: str) -> Dict[int, List[str]]:
    if not ticket_ids:
        return {}
    placeholders = ",".join("?" for _ in ticket_ids)
    mapping: Dict[int, List[str]] = {ticket_id: [] for ticket_id in ticket_ids}
    with get_connection() as connection:
        rows = connection.execute(
            f"""
            SELECT ticket_id, {value_column} AS value
            FROM {table_name}
            WHERE ticket_id IN ({placeholders})
            ORDER BY id
            """,
            ticket_ids,
        ).fetchall()
    for row in rows:
        ticket_id = int(row["ticket_id"])
        value = str(row["value"] or "").strip()
        if value:
            mapping.setdefault(ticket_id, []).append(value)
    return mapping


def attach_ticket_relations(tickets: List[Dict[str, object]]) -> List[Dict[str, object]]:
    ticket_ids = [int(ticket["id"]) for ticket in tickets if ticket.get("id") is not None]
    store_map = relation_map_for_tickets(ticket_ids, "ticket_stores", "store_name")
    brand_map = relation_map_for_tickets(ticket_ids, "ticket_brands", "brand")
    for ticket in tickets:
        ticket_id = int(ticket["id"])
        ticket["store_names"] = unique_clean_values(store_map.get(ticket_id, [])) or split_multi_value_text(ticket.get("store_name"))
        ticket["brand_names"] = unique_clean_values(brand_map.get(ticket_id, [])) or split_multi_value_text(ticket.get("brand"))
        annotate_ticket_runtime(ticket)
    return tickets


def build_query_params(filters: Dict[str, str], sort: str, page: Optional[int] = None) -> str:
    params = {key: value for key, value in filters.items() if str(value or "").strip()}
    if sort:
        params["sort"] = sort
    if page is not None:
        params["page"] = str(page)
    return urlencode(params)


def validate_submission(
    store_names: List[str],
    submitter: str,
    request_type: str,
    urgency: str,
    quantity: str,
    description: str,
    stores: Iterable[str],
    request_types: Iterable[str],
    urgency_levels: Iterable[str],
) -> Optional[str]:
    valid_stores = set(stores)
    if not store_names:
        return "请至少选择一个门店。"
    if any(store_name not in valid_stores for store_name in store_names):
        return "请选择有效门店。"
    if not submitter.strip():
        return "请填写提报人。"
    if request_type not in request_types:
        return "请选择有效需求类型。"
    if urgency not in urgency_levels:
        return "请选择有效紧急程度。"
    if not description.strip():
        return "请填写问题说明。"
    if quantity.strip() and not quantity.strip().isdigit():
        return "数量只能填写数字。"
    return None


def validate_request_type_rule(
    request_type: str,
    values: Dict[str, object],
    prepared_images: List[Tuple[str, bytes]],
    prepared_files: List[PreparedFile],
) -> Optional[str]:
    rule = load_request_type_rules().get(request_type)
    if not rule:
        return None
    for field in rule.get("required_fields", []):
        if not str(values.get(str(field), "")).strip():
            if str(field) == "brand":
                return f"{request_type}必须至少选择或填写一个品牌。"
            return f"{request_type}必须填写{REQUEST_RULE_FIELD_LABELS.get(str(field), str(field))}。"
    if rule.get("require_image") and not prepared_images:
        return f"{request_type}必须上传至少一张图片。"
    if rule.get("require_file") and not prepared_files:
        return f"{request_type}必须上传至少一个文件附件。"
    if rule.get("require_any_attachment") and not prepared_images and not prepared_files:
        return f"{request_type}必须上传图片或文件附件。"
    return None


async def prepare_images(images: Optional[List[UploadFile]], config: AppConfig) -> List[Tuple[str, bytes]]:
    raw_images: List[Tuple[str, bytes]] = []
    for image in images or []:
        if not image or not image.filename:
            continue
        original_name = safe_uploaded_name(image.filename)
        content = await image.read()
        if not content:
            continue
        raw_images.append((original_name, content))

    if len(raw_images) > config.max_image_count:
        raise ValueError(f"最多上传 {config.max_image_count} 张图片。")

    total_bytes = sum(len(content) for _, content in raw_images)
    if total_bytes > config.max_total_upload_bytes:
        raise ValueError(f"图片总大小不能超过 {config.max_total_upload_mb}MB。")

    prepared_images: List[Tuple[str, bytes]] = []
    allowed_extensions = set(config.allowed_image_extensions)
    for original_name, content in raw_images:
        extension = Path(original_name).suffix.lower().lstrip(".")
        if extension not in allowed_extensions:
            allowed_text = "、".join(config.allowed_image_extensions)
            raise ValueError(f"图片仅支持 {allowed_text} 格式。")
        if len(content) > config.max_image_bytes:
            raise ValueError(f"单张图片不能超过 {config.max_image_mb}MB。")

        try:
            with Image.open(BytesIO(content)) as image_file:
                image_format = str(image_file.format or "").upper()
                image_file.verify()
        except (UnidentifiedImageError, OSError, ValueError):
            raise ValueError("图片文件无法识别，请重新上传。") from None

        valid_extensions = IMAGE_FORMAT_EXTENSIONS.get(image_format, {image_format.lower()})
        if extension not in valid_extensions:
            raise ValueError("图片后缀与实际格式不匹配，请重新上传。")

        prepared_images.append((extension, content))
    return prepared_images


async def prepare_files(files: Optional[List[UploadFile]], config: AppConfig) -> List[PreparedFile]:
    raw_files: List[PreparedFile] = []
    for upload in files or []:
        if not upload or not upload.filename:
            continue
        original_name = safe_uploaded_name(upload.filename)
        content = await upload.read()
        if not original_name or not content:
            continue
        extension = Path(original_name).suffix.lower().lstrip(".")
        raw_files.append(
            PreparedFile(
                original_filename=original_name,
                file_ext=extension,
                file_size=len(content),
                content=content,
            )
        )

    if len(raw_files) > config.max_file_count:
        raise ValueError(f"最多上传 {config.max_file_count} 个文件。")

    total_bytes = sum(file.file_size for file in raw_files)
    if total_bytes > config.max_total_file_upload_bytes:
        raise ValueError(f"文件总大小不能超过 {config.max_total_file_upload_mb}MB。")

    allowed_extensions = set(config.allowed_file_extensions)
    for file in raw_files:
        if not file.file_ext or file.file_ext in BLOCKED_FILE_EXTENSIONS or file.file_ext not in allowed_extensions:
            allowed_text = "、".join(config.allowed_file_extensions)
            raise ValueError(f"文件仅支持 {allowed_text} 格式。")
        if file.file_size > config.max_file_bytes:
            raise ValueError(f"单个文件不能超过 {config.max_file_mb}MB。")
    return raw_files


def save_images(ticket_no: str, prepared_images: List[Tuple[str, bytes]]) -> Tuple[List[str], List[Path]]:
    image_paths: List[str] = []
    saved_files: List[Path] = []
    upload_dir = get_upload_dir()
    upload_dir.mkdir(parents=True, exist_ok=True)
    try:
        for index, (extension, content) in enumerate(prepared_images, start=1):
            filename = f"{ticket_no}_{index}_{uuid.uuid4().hex[:8]}.{extension}"
            target_path = upload_dir / filename
            target_path.write_bytes(content)
            saved_files.append(target_path)
            image_paths.append(f"uploads/{filename}")
        return image_paths, saved_files
    except Exception:
        cleanup_saved_files(saved_files)
        raise


def save_files(ticket_no: str, prepared_files: List[PreparedFile]) -> Tuple[List[Dict[str, object]], List[Path]]:
    file_records: List[Dict[str, object]] = []
    saved_files: List[Path] = []
    upload_dir = get_upload_dir()
    upload_dir.mkdir(parents=True, exist_ok=True)
    try:
        for prepared_file in prepared_files:
            stored_filename = f"FILE-{ticket_no}-{uuid.uuid4().hex}.{prepared_file.file_ext}"
            target_path = upload_dir / stored_filename
            target_path.write_bytes(prepared_file.content)
            saved_files.append(target_path)
            file_records.append(
                {
                    "original_filename": prepared_file.original_filename,
                    "stored_filename": stored_filename,
                    "file_path": f"uploads/{stored_filename}",
                    "file_ext": prepared_file.file_ext,
                    "file_size": prepared_file.file_size,
                }
            )
        return file_records, saved_files
    except Exception:
        cleanup_saved_files(saved_files)
        raise


def cleanup_saved_files(paths: List[Path]) -> None:
    for path in paths:
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass


def is_ticket_no_conflict(error: sqlite3.IntegrityError) -> bool:
    return "ticket_no" in str(error).lower()


def create_ticket_with_images(
    ticket_data: Dict[str, object],
    prepared_images: List[Tuple[str, bytes]],
    config: AppConfig,
    prepared_files: Optional[List[PreparedFile]] = None,
) -> Tuple[int, str]:
    prepared_files = prepared_files or []
    for _attempt in range(3):
        saved_files: List[Path] = []
        timestamp = now_text()
        try:
            with get_connection() as connection:
                ticket_no = generate_ticket_no(connection)
                try:
                    cursor = connection.execute(
                        """
                        INSERT INTO tickets (
                            ticket_no, created_at, updated_at, store_name, submitter,
                            request_type, urgency, brand, product_name, sku_barcode,
                            quantity, description, expected_finish_date, status,
                            assigned_to, handler_note, closed_at
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            ticket_no,
                            timestamp,
                            timestamp,
                            ticket_data["store_name"],
                            ticket_data["submitter"],
                            ticket_data["request_type"],
                            ticket_data["urgency"],
                            ticket_data["brand"],
                            ticket_data["product_name"],
                            ticket_data["sku_barcode"],
                            ticket_data["quantity"],
                            ticket_data["description"],
                            ticket_data["expected_finish_date"],
                            config.default_status,
                            "",
                            "",
                            None,
                        ),
                    )
                except sqlite3.IntegrityError as exc:
                    if is_ticket_no_conflict(exc):
                        continue
                    raise

                ticket_id = int(cursor.lastrowid)
                for store_name in ticket_data.get("store_names", []):
                    connection.execute(
                        """
                        INSERT OR IGNORE INTO ticket_stores (ticket_id, store_name, created_at)
                        VALUES (?, ?, ?)
                        """,
                        (ticket_id, store_name, timestamp),
                    )
                for brand in ticket_data.get("brands", []):
                    connection.execute(
                        """
                        INSERT OR IGNORE INTO ticket_brands (ticket_id, brand, created_at)
                        VALUES (?, ?, ?)
                        """,
                        (ticket_id, brand, timestamp),
                    )
                image_paths, saved_image_files = save_images(ticket_no, prepared_images)
                saved_files.extend(saved_image_files)
                for image_path in image_paths:
                    connection.execute(
                        """
                        INSERT INTO ticket_images (ticket_id, image_path, uploaded_at, source, uploaded_by, supplement_id)
                        VALUES (?, ?, ?, ?, ?, ?)
                        """,
                        (ticket_id, image_path, timestamp, "store_initial", ticket_data["submitter"], None),
                    )
                file_records, saved_attachment_files = save_files(ticket_no, prepared_files)
                saved_files.extend(saved_attachment_files)
                for file_record in file_records:
                    connection.execute(
                        """
                        INSERT INTO ticket_files (
                            ticket_id, original_filename, stored_filename, file_path,
                            file_ext, file_size, uploaded_at, source, uploaded_by, supplement_id
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            ticket_id,
                            file_record["original_filename"],
                            file_record["stored_filename"],
                            file_record["file_path"],
                            file_record["file_ext"],
                            file_record["file_size"],
                            timestamp,
                            "store_initial",
                            ticket_data["submitter"],
                            None,
                        ),
                    )
            return ticket_id, ticket_no
        except sqlite3.IntegrityError as exc:
            cleanup_saved_files(saved_files)
            if is_ticket_no_conflict(exc):
                continue
            raise
        except Exception:
            cleanup_saved_files(saved_files)
            raise
    raise RuntimeError("工单号生成冲突，请稍后重试。")


def build_ticket_where(filters: Dict[str, str]) -> Tuple[str, List[str]]:
    clauses: List[str] = []
    params: List[str] = []

    exact_fields = {
        "ticket_no": "ticket_no",
        "submitter": "submitter",
        "request_type": "request_type",
        "urgency": "urgency",
        "status": "status",
        "assigned_to": "assigned_to",
    }
    for filter_key, column_name in exact_fields.items():
        value = filters.get(filter_key, "").strip()
        if value:
            clauses.append(f"{column_name} = ?")
            params.append(value)

    store_name = filters.get("store_name", "").strip()
    if store_name:
        clauses.append(
            """
            EXISTS (
                SELECT 1
                FROM ticket_stores
                WHERE ticket_stores.ticket_id = tickets.id
                  AND ticket_stores.store_name = ?
            )
            """
        )
        params.append(store_name)

    date_start = filters.get("date_start", "").strip()
    if date_start:
        clauses.append("DATE(created_at) >= DATE(?)")
        params.append(date_start)

    date_end = filters.get("date_end", "").strip()
    if date_end:
        clauses.append("DATE(created_at) <= DATE(?)")
        params.append(date_end)

    keyword = filters.get("keyword", "").strip()
    if keyword:
        like_value = f"%{keyword}%"
        clauses.append(
            """
            (
                ticket_no LIKE ?
                OR store_name LIKE ?
                OR submitter LIKE ?
                OR brand LIKE ?
                OR product_name LIKE ?
                OR sku_barcode LIKE ?
                OR description LIKE ?
                OR assigned_to LIKE ?
                OR handler_note LIKE ?
            )
            """
        )
        params.extend([like_value] * 9)

    where_sql = " WHERE " + " AND ".join(clauses) if clauses else ""
    return where_sql, params


def build_case_order(values: List[str]) -> str:
    cases = []
    for index, value in enumerate(values):
        escaped_value = value.replace("'", "''")
        cases.append(f"WHEN '{escaped_value}' THEN {index}")
    return " ".join(cases)


def build_order_sql(sort: str, config: AppConfig) -> str:
    if sort == "urgency":
        urgency_cases = build_case_order(config.urgency_levels)
        return """
        ORDER BY
            CASE urgency
                {urgency_cases}
                ELSE {fallback_index}
            END,
            created_at DESC,
            tickets.id DESC
        """.format(urgency_cases=urgency_cases, fallback_index=len(config.urgency_levels))
    if sort == "status":
        status_cases = build_case_order(config.statuses)
        return """
        ORDER BY
            CASE status
                {status_cases}
                ELSE {fallback_index}
            END,
            created_at DESC,
            tickets.id DESC
        """.format(status_cases=status_cases, fallback_index=len(config.statuses))
    return "ORDER BY created_at DESC, tickets.id DESC"


def count_tickets(filters: Dict[str, str]) -> int:
    if filters.get("due_status", "").strip():
        return len(fetch_tickets(filters, "newest", load_app_config()))
    where_sql, params = build_ticket_where(filters)
    with get_connection() as connection:
        row = connection.execute(f"SELECT COUNT(*) AS total FROM tickets {where_sql}", params).fetchone()
    return int(row["total"] or 0)


def fetch_ticket_summary(filters: Dict[str, str]) -> Dict[str, int]:
    if filters.get("due_status", "").strip():
        tickets = fetch_tickets(filters, "newest", load_app_config())
        return {
            "total_count": len(tickets),
            "pending_count": sum(1 for ticket in tickets if ticket.get("status") == "待处理"),
            "processing_count": sum(1 for ticket in tickets if ticket.get("status") == "处理中"),
            "today_urgent_count": sum(1 for ticket in tickets if ticket.get("urgency") == "当天必须处理"),
            "completed_count": sum(1 for ticket in tickets if ticket.get("status") == COMPLETED_STATUS),
        }
    where_sql, params = build_ticket_where(filters)
    with get_connection() as connection:
        row = connection.execute(
            f"""
            SELECT
                COUNT(*) AS total_count,
                SUM(CASE WHEN status = '待处理' THEN 1 ELSE 0 END) AS pending_count,
                SUM(CASE WHEN status = '处理中' THEN 1 ELSE 0 END) AS processing_count,
                SUM(CASE WHEN urgency = '当天必须处理' THEN 1 ELSE 0 END) AS today_urgent_count,
                SUM(CASE WHEN status = '已完成' THEN 1 ELSE 0 END) AS completed_count
            FROM tickets
            {where_sql}
            """,
            params,
        ).fetchone()
    return {
        "total_count": int(row["total_count"] or 0),
        "pending_count": int(row["pending_count"] or 0),
        "processing_count": int(row["processing_count"] or 0),
        "today_urgent_count": int(row["today_urgent_count"] or 0),
        "completed_count": int(row["completed_count"] or 0),
    }


def fetch_tickets(
    filters: Dict[str, str],
    sort: str,
    config: AppConfig,
    limit: Optional[int] = None,
    offset: int = 0,
) -> List[Dict[str, object]]:
    where_sql, params = build_ticket_where(filters)
    order_sql = build_order_sql(sort, config)
    limit_sql = ""
    query_params: List[object] = list(params)
    if limit is not None:
        limit_sql = "LIMIT ? OFFSET ?"
        query_params.extend([limit, offset])
    with get_connection() as connection:
        rows = connection.execute(
            f"""
            SELECT
                tickets.id, ticket_no, created_at, updated_at, store_name, submitter,
                request_type, urgency, brand, product_name, sku_barcode,
                quantity, description, expected_finish_date, status, assigned_to,
                handler_note, closed_at,
                COALESCE(image_counts.image_count, 0) AS image_count,
                COALESCE(file_counts.file_count, 0) AS file_count
            FROM tickets
            LEFT JOIN (
                SELECT ticket_id, COUNT(*) AS image_count
                FROM ticket_images
                GROUP BY ticket_id
            ) AS image_counts ON image_counts.ticket_id = tickets.id
            LEFT JOIN (
                SELECT ticket_id, COUNT(*) AS file_count
                FROM ticket_files
                GROUP BY ticket_id
            ) AS file_counts ON file_counts.ticket_id = tickets.id
            {where_sql}
            {order_sql}
            {limit_sql}
            """,
            query_params,
        ).fetchall()
    tickets = attach_ticket_relations([dict(row) for row in rows])
    due_filter = filters.get("due_status", "").strip()
    if due_filter:
        tickets = [ticket for ticket in tickets if ticket.get("due_status") == due_filter]
    return tickets


def fetch_ticket_page(filters: Dict[str, str], sort: str, config: AppConfig, page: int) -> Tuple[List[Dict[str, object]], Dict[str, int]]:
    if filters.get("due_status", "").strip():
        all_tickets = fetch_tickets(filters, sort, config)
        total_count = len(all_tickets)
        total_pages = max((total_count + config.page_size - 1) // config.page_size, 1)
        current_page = min(max(page, 1), total_pages)
        offset = (current_page - 1) * config.page_size
        return all_tickets[offset : offset + config.page_size], {
            "current_page": current_page,
            "total_pages": total_pages,
            "total_count": total_count,
            "page_size": config.page_size,
            "has_prev": 1 if current_page > 1 else 0,
            "has_next": 1 if current_page < total_pages else 0,
            "prev_page": max(current_page - 1, 1),
            "next_page": min(current_page + 1, total_pages),
        }
    total_count = count_tickets(filters)
    total_pages = max((total_count + config.page_size - 1) // config.page_size, 1)
    current_page = min(max(page, 1), total_pages)
    offset = (current_page - 1) * config.page_size
    tickets = fetch_tickets(filters, sort, config, limit=config.page_size, offset=offset)
    return tickets, {
        "current_page": current_page,
        "total_pages": total_pages,
        "total_count": total_count,
        "page_size": config.page_size,
        "has_prev": 1 if current_page > 1 else 0,
        "has_next": 1 if current_page < total_pages else 0,
        "prev_page": max(current_page - 1, 1),
        "next_page": min(current_page + 1, total_pages),
    }


def fetch_ticket(ticket_id: int) -> Optional[Dict[str, object]]:
    with get_connection() as connection:
        row = connection.execute("SELECT * FROM tickets WHERE id = ?", (ticket_id,)).fetchone()
    if not row:
        return None
    tickets = attach_ticket_relations([dict(row)])
    return tickets[0] if tickets else None


def fetch_store_ticket(ticket_id: int, store_name: str) -> Optional[Dict[str, object]]:
    clean_store_name = store_name.strip()
    if not clean_store_name:
        return None
    with get_connection() as connection:
        row = connection.execute(
            """
            SELECT tickets.*
            FROM tickets
            WHERE tickets.id = ?
              AND EXISTS (
                  SELECT 1
                  FROM ticket_stores
                  WHERE ticket_stores.ticket_id = tickets.id
                    AND ticket_stores.store_name = ?
              )
            """,
            (ticket_id, clean_store_name),
        ).fetchone()
    if not row:
        return None
    tickets = attach_ticket_relations([dict(row)])
    ticket = tickets[0]
    ticket["query_store_name"] = clean_store_name
    return ticket


def fetch_store_ticket_supplements(ticket_id: int) -> List[Dict[str, object]]:
    with get_connection() as connection:
        rows = connection.execute(
            """
            SELECT id, ticket_id, store_name, submitter, note, image_count, file_count, created_at
            FROM ticket_supplements
            WHERE ticket_id = ?
            ORDER BY created_at DESC, id DESC
            """,
            (ticket_id,),
        ).fetchall()
    return [dict(row) for row in rows]


def fetch_store_visible_logs(ticket_id: int) -> List[Dict[str, object]]:
    with get_connection() as connection:
        rows = connection.execute(
            """
            SELECT id, ticket_id, action, old_status, new_status, note, created_at
            FROM ticket_logs
            WHERE ticket_id = ?
              AND action IN ('update', '门店补充资料')
            ORDER BY created_at DESC, id DESC
            """,
            (ticket_id,),
        ).fetchall()
    return [dict(row) for row in rows]


def fetch_ticket_attachment_counts(ticket_id: int) -> Dict[str, object]:
    with get_connection() as connection:
        image_row = connection.execute(
            "SELECT COUNT(*) AS total FROM ticket_images WHERE ticket_id = ?",
            (ticket_id,),
        ).fetchone()
        file_row = connection.execute(
            "SELECT COUNT(*) AS total FROM ticket_files WHERE ticket_id = ?",
            (ticket_id,),
        ).fetchone()
        supplement_row = connection.execute(
            """
            SELECT COUNT(*) AS total, MAX(created_at) AS latest_created_at
            FROM ticket_supplements
            WHERE ticket_id = ?
            """,
            (ticket_id,),
        ).fetchone()
    return {
        "image_count": int(image_row["total"] or 0),
        "file_count": int(file_row["total"] or 0),
        "supplement_count": int(supplement_row["total"] or 0),
        "latest_supplement_at": supplement_row["latest_created_at"] or "",
    }


def fetch_ticket_images(ticket_id: int) -> List[Dict[str, object]]:
    with get_connection() as connection:
        rows = connection.execute(
            """
            SELECT id, ticket_id, image_path, uploaded_at, source, uploaded_by, supplement_id
            FROM ticket_images
            WHERE ticket_id = ?
            ORDER BY id
            """,
            (ticket_id,),
        ).fetchall()
    images = [dict(row) for row in rows]
    for image in images:
        image["filename"] = image_filename(str(image.get("image_path") or ""))
        image["protected_url"] = protected_upload_url(str(image.get("image_path") or ""))
    return images


def fetch_ticket_files(ticket_id: int) -> List[Dict[str, object]]:
    with get_connection() as connection:
        rows = connection.execute(
            """
            SELECT id, ticket_id, original_filename, stored_filename, file_path,
                   file_ext, file_size, uploaded_at, source, uploaded_by, supplement_id
            FROM ticket_files
            WHERE ticket_id = ?
            ORDER BY id
            """,
            (ticket_id,),
        ).fetchall()
    files = [dict(row) for row in rows]
    for file in files:
        file["download_url"] = protected_file_url(file["id"])
        file["size_label"] = file_size_label(file.get("file_size"))
    return files


def fetch_ticket_file(file_id: int, ticket_id: Optional[int] = None) -> Optional[Dict[str, object]]:
    sql = """
        SELECT id, ticket_id, original_filename, stored_filename, file_path,
               file_ext, file_size, uploaded_at, source, uploaded_by, supplement_id
        FROM ticket_files
        WHERE id = ?
    """
    params: List[object] = [file_id]
    if ticket_id is not None:
        sql += " AND ticket_id = ?"
        params.append(ticket_id)
    with get_connection() as connection:
        row = connection.execute(sql, params).fetchone()
    return dict(row) if row else None


def fetch_ticket_image(image_id: int, ticket_id: int) -> Optional[Dict[str, object]]:
    with get_connection() as connection:
        row = connection.execute(
            """
            SELECT id, ticket_id, image_path, uploaded_at, source, uploaded_by, supplement_id
            FROM ticket_images
            WHERE id = ? AND ticket_id = ?
            """,
            (image_id, ticket_id),
        ).fetchone()
    return dict(row) if row else None


def fetch_ticket_logs(ticket_id: int) -> List[Dict[str, object]]:
    with get_connection() as connection:
        rows = connection.execute(
            """
            SELECT id, ticket_id, action, old_status, new_status, old_assigned_to,
                   new_assigned_to, note, operator, created_at
            FROM ticket_logs
            WHERE ticket_id = ?
            ORDER BY created_at DESC, id DESC
            """,
            (ticket_id,),
        ).fetchall()
    return [dict(row) for row in rows]


def add_ticket_attachments(
    ticket_id: int,
    ticket_no: str,
    prepared_images: List[Tuple[str, bytes]],
    prepared_files: List[PreparedFile],
    operator: str,
) -> Tuple[int, int]:
    saved_files: List[Path] = []
    timestamp = now_text()
    image_count = 0
    file_count = 0
    try:
        image_paths, saved_image_files = save_images(ticket_no, prepared_images)
        saved_files.extend(saved_image_files)
        file_records, saved_attachment_files = save_files(ticket_no, prepared_files)
        saved_files.extend(saved_attachment_files)
        image_count = len(image_paths)
        file_count = len(file_records)
        note = f"新增图片 {image_count} 张，文件 {file_count} 个"
        with get_connection() as connection:
            ticket_exists = connection.execute("SELECT 1 FROM tickets WHERE id = ?", (ticket_id,)).fetchone()
            if not ticket_exists:
                raise HTTPException(status_code=404, detail="工单不存在")
            for image_path in image_paths:
                connection.execute(
                    """
                    INSERT INTO ticket_images (ticket_id, image_path, uploaded_at, source, uploaded_by, supplement_id)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (ticket_id, image_path, timestamp, "admin_supplement", operator, None),
                )
            for file_record in file_records:
                connection.execute(
                    """
                    INSERT INTO ticket_files (
                        ticket_id, original_filename, stored_filename, file_path,
                        file_ext, file_size, uploaded_at, source, uploaded_by, supplement_id
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        ticket_id,
                        file_record["original_filename"],
                        file_record["stored_filename"],
                        file_record["file_path"],
                        file_record["file_ext"],
                        file_record["file_size"],
                        timestamp,
                        "admin_supplement",
                        operator,
                        None,
                    ),
                )
            connection.execute("UPDATE tickets SET updated_at = ? WHERE id = ?", (timestamp, ticket_id))
            connection.execute(
                """
                INSERT INTO ticket_logs (
                    ticket_id, action, old_status, new_status, old_assigned_to,
                    new_assigned_to, note, operator, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (ticket_id, "补充附件", None, None, None, None, note, operator, timestamp),
            )
    except Exception:
        cleanup_saved_files(saved_files)
        raise
    return image_count, file_count


def notification_severity_for_urgency(urgency: str) -> str:
    if urgency == "当天必须处理":
        return "urgent"
    if urgency == "加急":
        return "warning"
    return "info"


def create_notification_event(
    event_type: str,
    ticket_id: Optional[int],
    ticket_no: str,
    store_name: str,
    title: str,
    content: str,
    severity: str = "info",
    created_by: str = "",
) -> int:
    severity_value = severity if severity in {"info", "warning", "urgent"} else "info"
    with get_connection() as connection:
        cursor = connection.execute(
            """
            INSERT INTO notification_events (
                event_type, ticket_id, ticket_no, store_name, title, content,
                severity, created_by, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                event_type,
                ticket_id,
                ticket_no,
                store_name,
                title,
                content,
                severity_value,
                created_by,
                now_text(),
            ),
        )
        return int(cursor.lastrowid)


def create_new_ticket_notification(ticket: Dict[str, object]) -> None:
    create_notification_event(
        event_type="new_ticket",
        ticket_id=int(ticket["id"]),
        ticket_no=str(ticket.get("ticket_no") or ""),
        store_name=str(ticket.get("store_name") or ""),
        title="新工单",
        content=f"{ticket.get('store_name')} 提交了 {ticket.get('request_type')} 工单",
        severity=notification_severity_for_urgency(str(ticket.get("urgency") or "")),
        created_by=f"门店:{ticket.get('submitter') or ''}",
    )


def create_store_supplement_notification(ticket: Dict[str, object], submitter: str) -> None:
    create_notification_event(
        event_type="store_supplement",
        ticket_id=int(ticket["id"]),
        ticket_no=str(ticket.get("ticket_no") or ""),
        store_name=str(ticket.get("store_name") or ""),
        title="门店补充资料",
        content=f"{ticket.get('store_name')} 为工单 {ticket.get('ticket_no')} 补充了资料",
        severity="warning",
        created_by=f"门店:{submitter}",
    )


def create_need_store_supplement_notification(ticket: Dict[str, object], operator: str) -> None:
    create_notification_event(
        event_type="need_store_supplement",
        ticket_id=int(ticket["id"]),
        ticket_no=str(ticket.get("ticket_no") or ""),
        store_name=str(ticket.get("store_name") or ""),
        title="待门店补充",
        content=f"工单 {ticket.get('ticket_no')} 已标记为待门店补充",
        severity="warning",
        created_by=operator,
    )


def fetch_notifications(
    username: str,
    after_id: int = 0,
    limit: int = 20,
    unread_only: bool = False,
) -> List[Dict[str, object]]:
    clean_username = username.strip()
    clean_limit = min(max(int(limit or 20), 1), 100)
    clauses: List[str] = []
    params: List[object] = [clean_username]
    if after_id > 0:
        clauses.append("events.id > ?")
        params.append(after_id)
    if unread_only:
        clauses.append("reads.id IS NULL")
    where_sql = "WHERE " + " AND ".join(clauses) if clauses else ""
    params.append(clean_limit)
    with get_connection() as connection:
        rows = connection.execute(
            f"""
            SELECT
                events.id, events.event_type, events.ticket_id, events.ticket_no,
                events.store_name, events.title, events.content, events.severity,
                events.created_by, events.created_at,
                CASE WHEN reads.id IS NULL THEN 0 ELSE 1 END AS is_read
            FROM notification_events AS events
            LEFT JOIN notification_reads AS reads
                ON reads.event_id = events.id AND reads.username = ?
            {where_sql}
            ORDER BY events.id DESC
            LIMIT ?
            """,
            params,
        ).fetchall()
    notifications = [dict(row) for row in rows]
    for notification in notifications:
        ticket_id = notification.get("ticket_id")
        notification["is_read"] = bool(notification.get("is_read"))
        notification["detail_url"] = f"/admin/ticket/{ticket_id}" if ticket_id else ""
    return notifications


def count_unread_notifications(username: str) -> int:
    with get_connection() as connection:
        row = connection.execute(
            """
            SELECT COUNT(*) AS unread_count
            FROM notification_events AS events
            LEFT JOIN notification_reads AS reads
                ON reads.event_id = events.id AND reads.username = ?
            WHERE reads.id IS NULL
            """,
            (username.strip(),),
        ).fetchone()
    return int(row["unread_count"] or 0)


def latest_notification_id() -> int:
    with get_connection() as connection:
        row = connection.execute("SELECT MAX(id) AS latest_id FROM notification_events").fetchone()
    return int(row["latest_id"] or 0)


def mark_notification_read(username: str, event_id: int) -> bool:
    timestamp = now_text()
    with get_connection() as connection:
        event = connection.execute("SELECT id FROM notification_events WHERE id = ?", (event_id,)).fetchone()
        if not event:
            return False
        connection.execute(
            """
            INSERT OR IGNORE INTO notification_reads (event_id, username, read_at)
            VALUES (?, ?, ?)
            """,
            (event_id, username.strip(), timestamp),
        )
    return True


def mark_all_notifications_read(username: str) -> None:
    timestamp = now_text()
    with get_connection() as connection:
        connection.execute(
            """
            INSERT OR IGNORE INTO notification_reads (event_id, username, read_at)
            SELECT id, ?, ? FROM notification_events
            """,
            (username.strip(), timestamp),
        )


def create_store_supplement(
    ticket: Dict[str, object],
    submitter: str,
    note: str,
    prepared_images: List[Tuple[str, bytes]],
    prepared_files: List[PreparedFile],
    config: AppConfig,
) -> int:
    ticket_id = int(ticket["id"])
    ticket_no = str(ticket["ticket_no"])
    store_name = str(ticket.get("query_store_name") or ticket.get("store_name") or "").strip()
    old_status = str(ticket.get("status") or "")
    old_assigned_to = str(ticket.get("assigned_to") or "")
    timestamp = now_text()
    saved_files: List[Path] = []
    try:
        image_paths, saved_image_files = save_images(ticket_no, prepared_images)
        saved_files.extend(saved_image_files)
        file_records, saved_attachment_files = save_files(ticket_no, prepared_files)
        saved_files.extend(saved_attachment_files)
        image_count = len(image_paths)
        file_count = len(file_records)
        new_status = (
            config.supplement_status_after_store_update
            if old_status == "待门店补充"
            else old_status
        )
        closed_at = ticket.get("closed_at")
        if new_status == COMPLETED_STATUS and not closed_at:
            closed_at = timestamp
        elif old_status == COMPLETED_STATUS and new_status != COMPLETED_STATUS:
            closed_at = None
        note_parts = []
        clean_note = note.strip()
        if clean_note:
            note_parts.append(clean_note)
        note_parts.append(f"新增图片 {image_count} 张，文件 {file_count} 个")
        log_note = "；".join(note_parts)

        with get_connection() as connection:
            current = connection.execute("SELECT * FROM tickets WHERE id = ?", (ticket_id,)).fetchone()
            if not current:
                raise HTTPException(status_code=404, detail="工单不存在")
            store_match = connection.execute(
                """
                SELECT 1
                FROM ticket_stores
                WHERE ticket_id = ? AND store_name = ?
                """,
                (ticket_id, store_name),
            ).fetchone()
            if not store_match:
                raise HTTPException(status_code=403, detail="门店不匹配")
            cursor = connection.execute(
                """
                INSERT INTO ticket_supplements (
                    ticket_id, store_name, submitter, note, image_count, file_count, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (ticket_id, store_name, submitter, clean_note, image_count, file_count, timestamp),
            )
            supplement_id = int(cursor.lastrowid)
            for image_path in image_paths:
                connection.execute(
                    """
                    INSERT INTO ticket_images (ticket_id, image_path, uploaded_at, source, uploaded_by, supplement_id)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (ticket_id, image_path, timestamp, "store_supplement", submitter, supplement_id),
                )
            for file_record in file_records:
                connection.execute(
                    """
                    INSERT INTO ticket_files (
                        ticket_id, original_filename, stored_filename, file_path,
                        file_ext, file_size, uploaded_at, source, uploaded_by, supplement_id
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        ticket_id,
                        file_record["original_filename"],
                        file_record["stored_filename"],
                        file_record["file_path"],
                        file_record["file_ext"],
                        file_record["file_size"],
                        timestamp,
                        "store_supplement",
                        submitter,
                        supplement_id,
                    ),
                )
            connection.execute(
                "UPDATE tickets SET status = ?, closed_at = ?, updated_at = ? WHERE id = ?",
                (new_status, closed_at, timestamp, ticket_id),
            )
            connection.execute(
                """
                INSERT INTO ticket_logs (
                    ticket_id, action, old_status, new_status, old_assigned_to,
                    new_assigned_to, note, operator, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    ticket_id,
                    "门店补充资料",
                    old_status,
                    new_status,
                    old_assigned_to,
                    old_assigned_to,
                    log_note,
                    f"门店:{submitter}",
                    timestamp,
                ),
            )
        return supplement_id
    except Exception:
        cleanup_saved_files(saved_files)
        raise


def build_public_query_params(filters: Dict[str, str], page: Optional[int] = None) -> str:
    params = {key: value for key, value in filters.items() if str(value or "").strip()}
    if page is not None:
        params["page"] = str(page)
    return urlencode(params)


def normalize_store_query_filters(raw_filters: Dict[str, str], config: AppConfig) -> Dict[str, str]:
    filters = {key: str(value or "").strip() for key, value in raw_filters.items()}
    store_name = filters.get("store_name", "")
    if store_name and store_name in config.stores:
        optional_keys = ["ticket_no", "submitter", "keyword", "status", "date_start", "date_end"]
        if not any(filters.get(key) for key in optional_keys):
            filters["date_start"] = (datetime.now().date() - timedelta(days=config.store_query_default_days)).isoformat()
    return filters


def fetch_store_query_page(
    filters: Dict[str, str],
    config: AppConfig,
    page: int,
) -> Tuple[List[Dict[str, object]], Dict[str, int]]:
    page_size = config.store_query_page_size
    total_count = count_tickets(filters)
    total_pages = max((total_count + page_size - 1) // page_size, 1)
    current_page = min(max(page, 1), total_pages)
    offset = (current_page - 1) * page_size
    tickets = fetch_tickets(filters, "newest", config, limit=page_size, offset=offset)
    return tickets, {
        "current_page": current_page,
        "total_pages": total_pages,
        "total_count": total_count,
        "page_size": page_size,
        "has_prev": 1 if current_page > 1 else 0,
        "has_next": 1 if current_page < total_pages else 0,
        "prev_page": max(current_page - 1, 1),
        "next_page": min(current_page + 1, total_pages),
    }


def percent_value(count: int, total: int) -> int:
    if total <= 0:
        return 0
    return round(count * 100 / total)


def counter_rows(counter: Counter[str], total: int, limit: Optional[int] = None) -> List[Dict[str, object]]:
    rows = [
        {"label": label or "未填写", "count": count, "percent": percent_value(count, total)}
        for label, count in counter.most_common(limit)
    ]
    return rows


def fetch_dashboard_stats(filters: Dict[str, str]) -> Dict[str, object]:
    config = load_app_config()
    tickets = fetch_tickets(filters, "newest", config)
    total = len(tickets)
    today_prefix = datetime.now().date().isoformat()
    status_counter: Counter[str] = Counter(str(ticket.get("status") or "未填写") for ticket in tickets)
    type_counter: Counter[str] = Counter(str(ticket.get("request_type") or "未填写") for ticket in tickets)
    store_counter: Counter[str] = Counter()
    for ticket in tickets:
        store_names = ticket.get("store_names") or split_multi_value_text(ticket.get("store_name"))
        for store_name in store_names or ["未填写"]:
            store_counter[str(store_name or "未填写")] += 1
    handler_counter: Counter[str] = Counter(str(ticket.get("assigned_to") or "未指定") or "未指定" for ticket in tickets)
    urgency_counter: Counter[str] = Counter(str(ticket.get("urgency") or "未填写") for ticket in tickets)
    today_new_count = sum(1 for ticket in tickets if str(ticket.get("created_at") or "").startswith(today_prefix))
    overdue_count = sum(1 for ticket in tickets if ticket.get("due_status") == "已超时")
    due_today_count = sum(1 for ticket in tickets if ticket.get("due_status") == "今日到期")
    image_ticket_count = sum(1 for ticket in tickets if int(ticket.get("image_count") or 0) > 0)
    file_ticket_count = sum(1 for ticket in tickets if int(ticket.get("file_count") or 0) > 0)
    attachment_ticket_count = sum(
        1
        for ticket in tickets
        if int(ticket.get("image_count") or 0) > 0 or int(ticket.get("file_count") or 0) > 0
    )
    no_attachment_count = sum(
        1
        for ticket in tickets
        if int(ticket.get("image_count") or 0) == 0 and int(ticket.get("file_count") or 0) == 0
    )
    supplement_count = 0
    if tickets:
        ticket_ids = [int(ticket["id"]) for ticket in tickets]
        placeholders = ",".join("?" for _ in ticket_ids)
        with get_connection() as connection:
            row = connection.execute(
                f"SELECT COUNT(*) AS total FROM ticket_supplements WHERE ticket_id IN ({placeholders})",
                ticket_ids,
            ).fetchone()
        supplement_count = int(row["total"] or 0)
    cards = [
        {"label": "总工单数", "count": total, "percent": 100 if total else 0},
        {"label": "待处理数", "count": status_counter.get("待处理", 0), "percent": percent_value(status_counter.get("待处理", 0), total)},
        {"label": "处理中数", "count": status_counter.get("处理中", 0), "percent": percent_value(status_counter.get("处理中", 0), total)},
        {"label": "待门店补充数", "count": status_counter.get("待门店补充", 0), "percent": percent_value(status_counter.get("待门店补充", 0), total)},
        {"label": "已完成数", "count": status_counter.get("已完成", 0), "percent": percent_value(status_counter.get("已完成", 0), total)},
        {"label": "已驳回数", "count": status_counter.get("已驳回", 0), "percent": percent_value(status_counter.get("已驳回", 0), total)},
        {"label": "超时工单数", "count": overdue_count, "percent": percent_value(overdue_count, total)},
        {"label": "今日到期数", "count": due_today_count, "percent": percent_value(due_today_count, total)},
    ]
    attachment_rows = [
        {"label": "有图片工单数", "count": image_ticket_count, "percent": percent_value(image_ticket_count, total)},
        {"label": "有文件工单数", "count": file_ticket_count, "percent": percent_value(file_ticket_count, total)},
        {"label": "无附件工单数", "count": no_attachment_count, "percent": percent_value(no_attachment_count, total)},
    ]
    type_structure = [
        {"label": label, "count": type_counter.get(label, 0), "percent": percent_value(type_counter.get(label, 0), total)}
        for label in config.request_types
    ]
    return {
        "total": total,
        "today_new_count": today_new_count,
        "pending_count": status_counter.get("待处理", 0),
        "processing_count": status_counter.get("处理中", 0),
        "need_supplement_count": status_counter.get("待门店补充", 0),
        "completed_count": status_counter.get("已完成", 0),
        "overdue_count": overdue_count,
        "today_urgent_count": urgency_counter.get("当天必须处理", 0),
        "attachment_ticket_count": attachment_ticket_count,
        "store_supplement_count": supplement_count,
        "recent_tickets": tickets[:5],
        "cards": cards,
        "by_request_type": type_structure,
        "by_store": counter_rows(store_counter, total, limit=10),
        "by_handler": counter_rows(handler_counter, total),
        "by_status": counter_rows(status_counter, total),
        "by_urgency": counter_rows(urgency_counter, total),
        "attachments": attachment_rows,
    }


def build_excel(tickets: List[Dict[str, object]]) -> BytesIO:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "门店需求工单"

    sheet.append(EXCEL_HEADERS)
    header_fill = PatternFill(fill_type="solid", fgColor="E8EEF7")
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    image_map: Dict[int, List[str]] = {}
    file_name_map: Dict[int, List[str]] = {}
    file_url_map: Dict[int, List[str]] = {}
    if tickets:
        ids = [int(ticket["id"]) for ticket in tickets]
        placeholders = ",".join("?" for _ in ids)
        with get_connection() as connection:
            rows = connection.execute(
                f"SELECT ticket_id, image_path FROM ticket_images WHERE ticket_id IN ({placeholders}) ORDER BY id",
                ids,
            ).fetchall()
            file_rows = connection.execute(
                f"""
                SELECT id, ticket_id, original_filename
                FROM ticket_files
                WHERE ticket_id IN ({placeholders})
                ORDER BY id
                """,
                ids,
            ).fetchall()
        for row in rows:
            image_map.setdefault(int(row["ticket_id"]), []).append(protected_upload_url(str(row["image_path"])))
        for row in file_rows:
            ticket_id = int(row["ticket_id"])
            file_name_map.setdefault(ticket_id, []).append(str(row["original_filename"]))
            file_url_map.setdefault(ticket_id, []).append(protected_file_url(row["id"]))

    for ticket in tickets:
        sheet.append(
            [
                ticket.get("ticket_no"),
                ticket.get("created_at"),
                ticket.get("store_name"),
                ticket.get("submitter"),
                ticket.get("request_type"),
                ticket.get("urgency"),
                ticket.get("brand") or "",
                ticket.get("product_name") or "",
                ticket.get("sku_barcode") or "",
                ticket.get("quantity") if ticket.get("quantity") is not None else "",
                ticket.get("description"),
                "; ".join(image_map.get(int(ticket["id"]), [])),
                "; ".join(file_name_map.get(int(ticket["id"]), [])),
                "; ".join(file_url_map.get(int(ticket["id"]), [])),
                ticket.get("expected_finish_date") or "",
                ticket.get("status"),
                ticket.get("assigned_to") or "",
                ticket.get("handler_note") or "",
                ticket.get("closed_at") or "",
                ticket.get("updated_at"),
                processing_hours(ticket),
                overdue_text(ticket),
                ticket.get("due_status") or due_status_label(ticket),
            ]
        )

    widths = [22, 20, 16, 14, 14, 14, 16, 20, 20, 10, 38, 42, 26, 34, 18, 14, 14, 32, 20, 20, 16, 12, 14]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    sheet.freeze_panes = "A2"

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def create_app() -> FastAPI:
    load_env_file()
    warn_if_insecure_production_session()
    ensure_directories()
    init_db()

    app = FastAPI(title=load_app_config().app_name)
    templates = Jinja2Templates(directory=str(TEMPLATES_DIR))
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

    def render_submit_form(
        request: Request,
        status_code: int = 200,
        error: str = "",
        values: Optional[Dict[str, object]] = None,
    ) -> HTMLResponse:
        config = load_app_config()
        return templates.TemplateResponse(
            request,
            "submit.html",
            {
                "request": request,
                "stores": config.stores,
                "request_types": config.request_types,
                "urgency_levels": config.urgency_levels,
                "brands": config.brands,
                "image_accept": ",".join(f".{extension}" for extension in config.allowed_image_extensions),
                "file_accept": ",".join(f".{extension}" for extension in config.allowed_file_extensions),
                "max_image_count": config.max_image_count,
                "max_total_upload_mb": config.max_total_upload_mb,
                "max_image_mb": config.max_image_mb,
                "allowed_file_extensions": config.allowed_file_extensions,
                "max_file_count": config.max_file_count,
                "max_file_mb": config.max_file_mb,
                "max_total_file_upload_mb": config.max_total_file_upload_mb,
                "error": error,
                "values": values or {},
                "request_type_rules_json": json.dumps(load_request_type_rules(), ensure_ascii=False),
            },
            status_code=status_code,
        )

    def render_login_form(
        request: Request,
        status_code: int = 200,
        error: str = "",
        username: str = "",
        next_url: str = "/admin",
        logged_out: str = "",
    ) -> HTMLResponse:
        return templates.TemplateResponse(
            request,
            "login.html",
            {
                "request": request,
                "error": error,
                "username": username,
                "next_url": safe_admin_return_url(next_url),
                "logged_out": logged_out,
            },
            status_code=status_code,
        )

    def render_ticket_detail(
        request: Request,
        ticket_id: int,
        admin: str,
        saved: str = "",
        attachments_saved: str = "",
        upload_error: str = "",
        return_url: str = "",
        status_code: int = 200,
    ) -> HTMLResponse:
        ticket = fetch_ticket(ticket_id)
        if not ticket:
            raise HTTPException(status_code=404, detail="工单不存在")
        images = fetch_ticket_images(ticket_id)
        files = fetch_ticket_files(ticket_id)
        logs = fetch_ticket_logs(ticket_id)
        config = load_app_config()
        return templates.TemplateResponse(
            request,
            "ticket_detail.html",
            {
                "request": request,
                "ticket": ticket,
                "images": images,
                "files": files,
                "statuses": config.statuses,
                "handlers": config.handlers,
                "logs": logs,
                "saved": saved,
                "attachments_saved": attachments_saved,
                "upload_error": upload_error,
                "return_url": safe_admin_return_url(return_url),
                "admin_user": admin,
                "image_accept": ",".join(f".{extension}" for extension in config.allowed_image_extensions),
                "file_accept": ",".join(f".{extension}" for extension in config.allowed_file_extensions),
                "max_image_count": config.max_image_count,
                "max_image_mb": config.max_image_mb,
                "max_file_count": config.max_file_count,
                "max_file_mb": config.max_file_mb,
                "max_total_file_upload_mb": config.max_total_file_upload_mb,
                "csrf_token": current_csrf_token(request),
            },
            status_code=status_code,
        )

    def render_query_page(
        request: Request,
        filters: Optional[Dict[str, str]] = None,
        tickets: Optional[List[Dict[str, object]]] = None,
        pagination: Optional[Dict[str, int]] = None,
        error: str = "",
        searched: bool = False,
        status_code: int = 200,
    ) -> HTMLResponse:
        config = load_app_config()
        active_filters = filters or {}
        page_state = pagination or {
            "current_page": 1,
            "total_pages": 1,
            "total_count": 0,
            "page_size": config.store_query_page_size,
            "has_prev": 0,
            "has_next": 0,
            "prev_page": 1,
            "next_page": 1,
        }
        prev_query = build_public_query_params(active_filters, page_state["prev_page"])
        next_query = build_public_query_params(active_filters, page_state["next_page"])
        return templates.TemplateResponse(
            request,
            "query.html",
            {
                "request": request,
                "stores": config.stores,
                "statuses": config.statuses,
                "filters": active_filters,
                "tickets": tickets or [],
                "pagination": page_state,
                "prev_page_url": "/query" + (f"?{prev_query}" if prev_query else ""),
                "next_page_url": "/query" + (f"?{next_query}" if next_query else ""),
                "error": error,
                "searched": searched,
            },
            status_code=status_code,
        )

    def render_supplement_page(
        request: Request,
        ticket: Dict[str, object],
        store_name: str,
        return_url: str = "",
        error: str = "",
        success: bool = False,
        status_code: int = 200,
    ) -> HTMLResponse:
        config = load_app_config()
        query_url = store_query_list_return_url(store_name, return_url)
        detail_url = build_store_ticket_detail_url(int(ticket["id"]), store_name, query_url)
        supplement_url = build_store_ticket_supplement_url(int(ticket["id"]), store_name, query_url)
        return templates.TemplateResponse(
            request,
            "supplement.html",
            {
                "request": request,
                "ticket": ticket,
                "store_name": store_name,
                "error": error,
                "success": success,
                "query_url": query_url,
                "detail_url": detail_url,
                "supplement_url": supplement_url,
                "return_url": query_url,
                "image_accept": ",".join(f".{extension}" for extension in config.allowed_image_extensions),
                "file_accept": ",".join(f".{extension}" for extension in config.allowed_file_extensions),
                "max_image_count": config.max_image_count,
                "max_image_mb": config.max_image_mb,
                "max_file_count": config.max_file_count,
                "max_file_mb": config.max_file_mb,
                "max_total_file_upload_mb": config.max_total_file_upload_mb,
            },
            status_code=status_code,
        )

    def render_store_ticket_detail(
        request: Request,
        ticket: Dict[str, object],
        store_name: str,
        return_url: str = "",
        status_code: int = 200,
    ) -> HTMLResponse:
        query_url = store_query_list_return_url(store_name, return_url)
        detail_url = build_store_ticket_detail_url(int(ticket["id"]), store_name, query_url)
        supplement_url = build_store_ticket_supplement_url(int(ticket["id"]), store_name, query_url)
        supplements = fetch_store_ticket_supplements(int(ticket["id"]))
        return templates.TemplateResponse(
            request,
            "query_detail.html",
            {
                "request": request,
                "ticket": ticket,
                "store_name": store_name,
                "query_url": query_url,
                "detail_url": detail_url,
                "supplement_url": supplement_url,
                "attachment_counts": fetch_ticket_attachment_counts(int(ticket["id"])),
                "supplements": supplements,
                "logs": fetch_store_visible_logs(int(ticket["id"])),
                "needs_store_supplement": str(ticket.get("status") or "") == "待门店补充",
                "handler_note": str(ticket.get("handler_note") or "").strip(),
                "error": "",
            },
            status_code=status_code,
        )

    @app.get("/admin/login", response_class=HTMLResponse)
    def admin_login_page(request: Request, next: str = Query("/admin"), logged_out: str = Query("")) -> HTMLResponse:
        next_url = safe_admin_return_url(next)
        if current_admin_username(request):
            return RedirectResponse(url=next_url, status_code=303)
        return render_login_form(request, next_url=next_url, logged_out=logged_out)

    @app.post("/admin/login", response_class=HTMLResponse)
    def admin_login(
        request: Request,
        username: str = Form(""),
        password: str = Form(""),
        next: str = Form("/admin"),
    ) -> HTMLResponse:
        next_url = safe_admin_return_url(next)
        authenticated_username = authenticate_admin(username, password)
        if not authenticated_username:
            return render_login_form(
                request,
                status_code=400,
                error="用户名或密码不正确。",
                username=username.strip(),
                next_url=next_url,
            )
        response = RedirectResponse(url=next_url, status_code=303)
        max_age = get_session_max_age_seconds()
        response.set_cookie(
            SESSION_COOKIE_NAME,
            create_admin_session(authenticated_username, max_age_seconds=max_age),
            httponly=True,
            samesite="lax",
            max_age=max_age,
            secure=session_cookie_secure(),
        )
        return response

    @app.post("/admin/logout")
    def admin_logout(request: Request, csrf_token: str = Form("")) -> RedirectResponse:
        if current_admin_username(request):
            require_admin_csrf(request, csrf_token)
        response = RedirectResponse(url="/admin/login?logged_out=1", status_code=303)
        response.delete_cookie(SESSION_COOKIE_NAME)
        return response

    @app.get("/admin/logout")
    def admin_logout_get() -> RedirectResponse:
        response = RedirectResponse(url="/admin/login?logged_out=1", status_code=303)
        response.delete_cookie(SESSION_COOKIE_NAME)
        return response

    @app.get("/", include_in_schema=False)
    def root() -> RedirectResponse:
        return RedirectResponse(url="/submit", status_code=303)

    @app.get("/submit", response_class=HTMLResponse)
    def submit_page(request: Request) -> HTMLResponse:
        return render_submit_form(request)

    @app.post("/submit", response_class=HTMLResponse)
    async def submit_ticket(
        request: Request,
        store_names: Optional[List[str]] = Form(None),
        brands: Optional[List[str]] = Form(None),
        brand_extra: str = Form(""),
        store_name: str = Form(""),
        submitter: str = Form(""),
        request_type: str = Form(""),
        urgency: str = Form(""),
        brand: str = Form(""),
        product_name: str = Form(""),
        sku_barcode: str = Form(""),
        quantity: str = Form(""),
        description: str = Form(""),
        expected_finish_date: str = Form(""),
        images: Optional[List[UploadFile]] = File(None),
        files: Optional[List[UploadFile]] = File(None),
    ) -> HTMLResponse:
        config = load_app_config()
        normalized_stores = normalize_store_names(store_names, store_name)
        normalized_brands = normalize_brand_names(brands, brand_extra, brand)
        store_display = join_display_values(normalized_stores)
        brand_display = join_display_values(normalized_brands)
        form_values = {
            "store_name": store_name,
            "store_names": normalized_stores,
            "submitter": submitter,
            "request_type": request_type,
            "urgency": urgency,
            "brand": brand,
            "brands": normalized_brands,
            "brand_extra": brand_extra,
            "product_name": product_name,
            "sku_barcode": sku_barcode,
            "quantity": quantity,
            "description": description,
            "expected_finish_date": expected_finish_date,
        }
        error = validate_submission(
            normalized_stores,
            submitter,
            request_type,
            urgency,
            quantity,
            description,
            config.stores,
            config.request_types,
            config.urgency_levels,
        )
        if error:
            return render_submit_form(request, status_code=400, error=error, values=form_values)

        try:
            prepared_images = await prepare_images(images, config)
            prepared_files = await prepare_files(files, config)
        except ValueError as exc:
            return render_submit_form(request, status_code=400, error=str(exc), values=form_values)

        rule_values = dict(form_values)
        rule_values["brand"] = brand_display
        rule_error = validate_request_type_rule(request_type, rule_values, prepared_images, prepared_files)
        if rule_error:
            return render_submit_form(request, status_code=400, error=rule_error, values=form_values)

        timestamp = now_text()
        quantity_value = int(quantity.strip()) if quantity.strip() else None
        try:
            ticket_id, ticket_no = create_ticket_with_images(
                {
                    "created_at": timestamp,
                    "store_name": store_display,
                    "store_names": normalized_stores,
                    "submitter": submitter.strip(),
                    "request_type": request_type,
                    "urgency": urgency,
                    "brand": brand_display,
                    "brands": normalized_brands,
                    "product_name": product_name.strip(),
                    "sku_barcode": sku_barcode.strip(),
                    "quantity": quantity_value,
                    "description": description.strip(),
                    "expected_finish_date": expected_finish_date.strip(),
                },
                prepared_images,
                config,
                prepared_files,
            )
            try:
                created_ticket = fetch_ticket(ticket_id)
                if created_ticket:
                    create_new_ticket_notification(created_ticket)
            except Exception:
                pass
        except RuntimeError as exc:
            return render_submit_form(request, status_code=500, error=str(exc), values=form_values)
        except OSError:
            return render_submit_form(request, status_code=500, error="附件保存失败，请稍后重试。", values=form_values)

        return templates.TemplateResponse(
            request,
            "submit_success.html",
            {"request": request, "ticket_no": ticket_no},
        )

    @app.get("/query", response_class=HTMLResponse)
    def query_page(
        request: Request,
        store_name: str = Query(""),
        ticket_no: str = Query(""),
        submitter: str = Query(""),
        keyword: str = Query(""),
        status: str = Query(""),
        date_start: str = Query(""),
        date_end: str = Query(""),
        page: int = Query(1),
    ) -> HTMLResponse:
        config = load_app_config()
        raw_filters = {
            "store_name": store_name,
            "ticket_no": ticket_no,
            "submitter": submitter,
            "keyword": keyword,
            "status": status,
            "date_start": date_start,
            "date_end": date_end,
        }
        submitted = any(str(value or "").strip() for value in raw_filters.values())
        filters = normalize_store_query_filters(raw_filters, config)
        if not submitted:
            return render_query_page(request, filters=filters)
        if not filters.get("store_name") or filters.get("store_name") not in config.stores:
            return render_query_page(
                request,
                filters=filters,
                error="请选择门店后再查询。",
                searched=True,
                status_code=400,
            )
        if filters.get("status") and filters.get("status") not in config.statuses:
            return render_query_page(
                request,
                filters=filters,
                error="请选择有效状态。",
                searched=True,
                status_code=400,
            )
        tickets, pagination = fetch_store_query_page(filters, config, page)
        return_url = safe_query_return_url(request_path_with_query(request))
        for ticket in tickets:
            ticket["detail_url"] = build_store_ticket_detail_url(int(ticket["id"]), filters["store_name"], return_url)
            ticket["supplement_url"] = build_store_ticket_supplement_url(int(ticket["id"]), filters["store_name"], return_url)
        return render_query_page(request, filters=filters, tickets=tickets, pagination=pagination, searched=True)

    @app.post("/query")
    def query_submit(
        store_name: str = Form(""),
        ticket_no: str = Form(""),
        submitter: str = Form(""),
        keyword: str = Form(""),
        status: str = Form(""),
        date_start: str = Form(""),
        date_end: str = Form(""),
    ) -> RedirectResponse:
        params = {
            "store_name": store_name.strip(),
            "ticket_no": ticket_no.strip(),
            "submitter": submitter.strip(),
            "keyword": keyword.strip(),
            "status": status.strip(),
            "date_start": date_start.strip(),
            "date_end": date_end.strip(),
        }
        query = build_public_query_params(params)
        return RedirectResponse(url="/query" + (f"?{query}" if query else ""), status_code=303)

    @app.get("/query/ticket/{ticket_id}", response_class=HTMLResponse)
    def store_ticket_detail(
        request: Request,
        ticket_id: int,
        store_name: str = Query(""),
        return_url: str = Query(""),
    ) -> HTMLResponse:
        clean_store_name = store_name.strip()
        if not clean_store_name:
            return HTMLResponse("请选择门店后查看工单详情。", status_code=400)
        ticket = fetch_store_ticket(ticket_id, clean_store_name)
        if not ticket:
            raise HTTPException(status_code=404, detail="未找到该门店对应工单")
        return render_store_ticket_detail(request, ticket, clean_store_name, return_url=return_url)

    @app.get("/query/ticket/{ticket_id}/supplement", response_class=HTMLResponse)
    def supplement_page(
        request: Request,
        ticket_id: int,
        store_name: str = Query(""),
        return_url: str = Query(""),
    ) -> HTMLResponse:
        clean_store_name = store_name.strip()
        ticket = fetch_store_ticket(ticket_id, clean_store_name)
        if not ticket:
            raise HTTPException(status_code=403, detail="门店不匹配")
        return render_supplement_page(request, ticket, clean_store_name, return_url=return_url)

    @app.post("/query/ticket/{ticket_id}/supplement", response_class=HTMLResponse)
    async def submit_supplement(
        request: Request,
        ticket_id: int,
        store_name: str = Form(""),
        submitter: str = Form(""),
        note: str = Form(""),
        return_url: str = Form(""),
        images: Optional[List[UploadFile]] = File(None),
        files: Optional[List[UploadFile]] = File(None),
    ) -> HTMLResponse:
        clean_store_name = store_name.strip()
        ticket = fetch_store_ticket(ticket_id, clean_store_name)
        if not ticket:
            raise HTTPException(status_code=403, detail="门店不匹配")
        clean_submitter = submitter.strip()
        if not clean_submitter:
            return render_supplement_page(
                request,
                ticket,
                clean_store_name,
                return_url=return_url,
                error="请填写补充人。",
                status_code=400,
            )
        config = load_app_config()
        try:
            prepared_images = await prepare_images(images, config)
            prepared_files = await prepare_files(files, config)
        except ValueError as exc:
            return render_supplement_page(
                request,
                ticket,
                clean_store_name,
                return_url=return_url,
                error=str(exc),
                status_code=400,
            )
        if not note.strip() and not prepared_images and not prepared_files:
            return render_supplement_page(
                request,
                ticket,
                clean_store_name,
                return_url=return_url,
                error="请填写补充说明或上传附件。",
                status_code=400,
            )
        try:
            create_store_supplement(ticket, clean_submitter, note, prepared_images, prepared_files, config)
            try:
                updated_for_notification = fetch_ticket(ticket_id) or ticket
                create_store_supplement_notification(updated_for_notification, clean_submitter)
            except Exception:
                pass
        except OSError:
            return render_supplement_page(
                request,
                ticket,
                clean_store_name,
                return_url=return_url,
                error="附件保存失败，请稍后重试。",
                status_code=500,
            )
        updated_ticket = fetch_store_ticket(ticket_id, clean_store_name) or ticket
        return render_supplement_page(request, updated_ticket, clean_store_name, return_url=return_url, success=True)

    @app.get("/admin", response_class=HTMLResponse)
    def admin_page(
        request: Request,
        admin: str = Depends(require_admin),
        store_name: str = Query(""),
        request_type: str = Query(""),
        urgency: str = Query(""),
        status: str = Query(""),
        assigned_to: str = Query(""),
        date_start: str = Query(""),
        date_end: str = Query(""),
        keyword: str = Query(""),
        due_status: str = Query(""),
        sort: str = Query("newest"),
        page: int = Query(1),
    ) -> HTMLResponse:
        filters = {
            "store_name": store_name,
            "request_type": request_type,
            "urgency": urgency,
            "status": status,
            "assigned_to": assigned_to,
            "date_start": date_start,
            "date_end": date_end,
            "keyword": keyword,
            "due_status": due_status,
        }
        config = load_app_config()
        tickets, pagination = fetch_ticket_page(filters, sort, config, page)
        summary = fetch_ticket_summary(filters)
        return_url = safe_admin_return_url(request_path_with_query(request))
        for ticket in tickets:
            ticket["detail_url"] = f"/admin/ticket/{ticket['id']}?{urlencode({'return_url': return_url})}"
        export_query = build_query_params(filters, sort)
        export_url = "/admin/export" + (f"?{export_query}" if export_query else "")
        prev_query = build_query_params(filters, sort, pagination["prev_page"])
        next_query = build_query_params(filters, sort, pagination["next_page"])
        return templates.TemplateResponse(
            request,
            "admin.html",
            {
                "request": request,
                "tickets": tickets,
                "stores": config.stores,
                "request_types": config.request_types,
                "urgency_levels": config.urgency_levels,
                "statuses": config.statuses,
                "handlers": config.handlers,
                "due_status_options": DUE_STATUS_OPTIONS,
                "filters": filters,
                "sort": sort,
                "summary": summary,
                "export_url": export_url,
                "pagination": pagination,
                "prev_page_url": "/admin" + (f"?{prev_query}" if prev_query else ""),
                "next_page_url": "/admin" + (f"?{next_query}" if next_query else ""),
                "admin_user": admin,
                "csrf_token": current_csrf_token(request),
            },
        )

    @app.get("/admin/dashboard", response_class=HTMLResponse)
    def admin_dashboard(
        request: Request,
        admin: str = Depends(require_admin),
        store_name: str = Query(""),
        request_type: str = Query(""),
        status: str = Query(""),
        assigned_to: str = Query(""),
        date_start: str = Query(""),
        date_end: str = Query(""),
    ) -> HTMLResponse:
        config = load_app_config()
        filters = {
            "store_name": store_name,
            "request_type": request_type,
            "status": status,
            "assigned_to": assigned_to,
            "date_start": date_start,
            "date_end": date_end,
        }
        stats = fetch_dashboard_stats(filters)
        for ticket in stats["recent_tickets"]:
            ticket["detail_url"] = build_ticket_detail_url(int(ticket["id"]), "/admin/dashboard")
        return templates.TemplateResponse(
            request,
            "dashboard.html",
            {
                "request": request,
                "stores": config.stores,
                "request_types": config.request_types,
                "statuses": config.statuses,
                "handlers": config.handlers,
                "filters": filters,
                "stats": stats,
                "today_label": datetime.now().strftime("%Y-%m-%d"),
                "admin_user": admin,
                "csrf_token": current_csrf_token(request),
            },
        )

    @app.get("/admin/settings", response_class=HTMLResponse)
    def admin_settings(request: Request, admin: str = Depends(require_admin)) -> HTMLResponse:
        return templates.TemplateResponse(
            request,
            "settings.html",
            {
                "request": request,
                "admin_user": admin,
                "csrf_token": current_csrf_token(request),
                "config_files": [
                    "stores.json",
                    "request_types.json",
                    "urgency_levels.json",
                    "statuses.json",
                    "brands.json",
                    "handlers.json",
                    "system.json",
                    "request_type_rules.json",
                ],
            },
        )

    @app.get("/admin/account", response_class=HTMLResponse)
    def admin_account(request: Request, admin: str = Depends(require_admin)) -> HTMLResponse:
        return templates.TemplateResponse(
            request,
            "account.html",
            {
                "request": request,
                "admin_user": admin,
                "csrf_token": current_csrf_token(request),
            },
        )

    @app.get("/admin/system", response_class=HTMLResponse)
    def admin_system(request: Request, admin: str = Depends(require_admin)) -> HTMLResponse:
        return templates.TemplateResponse(
            request,
            "system.html",
            {
                "request": request,
                "admin_user": admin,
                "csrf_token": current_csrf_token(request),
                "database_path": str(get_db_path()),
                "upload_dir": str(get_upload_dir()),
                "config_dir": str(get_config_dir()),
            },
        )

    @app.get("/admin/api/notifications")
    def notifications_api(
        admin: str = Depends(require_admin),
        after_id: int = Query(0),
        limit: int = Query(20),
        unread_only: bool = Query(False),
    ) -> Dict[str, object]:
        notifications = fetch_notifications(
            admin,
            after_id=max(after_id, 0),
            limit=limit,
            unread_only=unread_only,
        )
        return {
            "unread_count": count_unread_notifications(admin),
            "latest_id": latest_notification_id(),
            "notifications": notifications,
        }

    @app.post("/admin/api/notifications/{event_id}/read")
    def mark_notification_read_api(
        request: Request,
        event_id: int,
        admin: str = Depends(require_admin),
        csrf_token: str = Form(""),
    ) -> Dict[str, object]:
        require_admin_csrf(request, csrf_token)
        if not mark_notification_read(admin, event_id):
            raise HTTPException(status_code=404, detail="消息不存在")
        return {
            "ok": True,
            "unread_count": count_unread_notifications(admin),
        }

    @app.post("/admin/api/notifications/read-all")
    def mark_all_notifications_read_api(
        request: Request,
        admin: str = Depends(require_admin),
        csrf_token: str = Form(""),
    ) -> Dict[str, object]:
        require_admin_csrf(request, csrf_token)
        mark_all_notifications_read(admin)
        return {
            "ok": True,
            "unread_count": count_unread_notifications(admin),
        }

    @app.get("/admin/ticket/{ticket_id}", response_class=HTMLResponse)
    def ticket_detail(
        request: Request,
        ticket_id: int,
        saved: str = Query(""),
        attachments_saved: str = Query(""),
        upload_error: str = Query(""),
        return_url: str = Query(""),
        admin: str = Depends(require_admin),
    ) -> HTMLResponse:
        return render_ticket_detail(
            request,
            ticket_id,
            admin,
            saved=saved,
            attachments_saved=attachments_saved,
            upload_error=upload_error,
            return_url=return_url,
        )

    @app.post("/admin/ticket/{ticket_id}")
    def update_ticket(
        request: Request,
        ticket_id: int,
        admin: str = Depends(require_admin),
        status: str = Form(""),
        assigned_to: str = Form(""),
        handler_note: str = Form(""),
        return_url: str = Form(""),
        csrf_token: str = Form(""),
    ) -> RedirectResponse:
        require_admin_csrf(request, csrf_token)
        config = load_app_config()
        if status not in config.statuses:
            raise HTTPException(status_code=400, detail="状态不正确")
        assigned_to = assigned_to.strip()
        if assigned_to and config.handlers and assigned_to not in config.handlers:
            raise HTTPException(status_code=400, detail="处理人不正确")
        handler_note = handler_note.strip()
        timestamp = now_text()
        should_notify_need_supplement = False
        notification_ticket: Optional[Dict[str, object]] = None
        with get_connection() as connection:
            old_row = connection.execute("SELECT * FROM tickets WHERE id = ?", (ticket_id,)).fetchone()
            if not old_row:
                raise HTTPException(status_code=404, detail="工单不存在")
            old_ticket = dict(old_row)
            old_status = str(old_ticket.get("status") or "")
            old_assigned_to = str(old_ticket.get("assigned_to") or "")
            old_note = str(old_ticket.get("handler_note") or "")
            closed_at = old_ticket.get("closed_at")
            if status == COMPLETED_STATUS and not closed_at:
                closed_at = timestamp
            elif old_status == COMPLETED_STATUS and status != COMPLETED_STATUS:
                closed_at = None

            changed = status != old_status or assigned_to != old_assigned_to or handler_note != old_note
            if changed:
                connection.execute(
                    """
                    UPDATE tickets
                    SET status = ?, assigned_to = ?, handler_note = ?, closed_at = ?, updated_at = ?
                    WHERE id = ?
                    """,
                    (status, assigned_to, handler_note, closed_at, timestamp, ticket_id),
                )
                connection.execute(
                    """
                    INSERT INTO ticket_logs (
                        ticket_id, action, old_status, new_status, old_assigned_to,
                        new_assigned_to, note, operator, created_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        ticket_id,
                        "update",
                        old_status,
                        status,
                        old_assigned_to,
                        assigned_to,
                        handler_note,
                        admin,
                        timestamp,
                    ),
                )
                should_notify_need_supplement = old_status != status and status == "待门店补充"
                if should_notify_need_supplement:
                    notification_ticket = dict(old_ticket)
                    notification_ticket["status"] = status
                    notification_ticket["updated_at"] = timestamp
                    notification_ticket["assigned_to"] = assigned_to
        if should_notify_need_supplement and notification_ticket:
            try:
                create_need_store_supplement_notification(notification_ticket, admin)
            except Exception:
                pass
        return RedirectResponse(url=build_ticket_detail_url(ticket_id, return_url, saved="1"), status_code=303)

    @app.post("/admin/ticket/{ticket_id}/attachments", response_class=HTMLResponse)
    async def add_attachments(
        request: Request,
        ticket_id: int,
        admin: str = Depends(require_admin),
        new_images: Optional[List[UploadFile]] = File(None),
        new_files: Optional[List[UploadFile]] = File(None),
        return_url: str = Form(""),
        csrf_token: str = Form(""),
    ) -> HTMLResponse:
        require_admin_csrf(request, csrf_token)
        ticket = fetch_ticket(ticket_id)
        if not ticket:
            raise HTTPException(status_code=404, detail="工单不存在")
        config = load_app_config()
        try:
            prepared_images = await prepare_images(new_images, config)
            prepared_files = await prepare_files(new_files, config)
        except ValueError as exc:
            return render_ticket_detail(
                request,
                ticket_id,
                admin,
                upload_error=str(exc),
                return_url=return_url,
                status_code=400,
            )
        if not prepared_images and not prepared_files:
            return render_ticket_detail(
                request,
                ticket_id,
                admin,
                upload_error="请选择要上传的附件。",
                return_url=return_url,
                status_code=400,
            )
        try:
            add_ticket_attachments(ticket_id, str(ticket["ticket_no"]), prepared_images, prepared_files, admin)
        except OSError:
            return render_ticket_detail(
                request,
                ticket_id,
                admin,
                upload_error="附件保存失败，请稍后重试。",
                return_url=return_url,
                status_code=500,
            )
        return RedirectResponse(url=build_ticket_detail_url(ticket_id, return_url, attachments_saved="1"), status_code=303)

    @app.post("/admin/ticket/{ticket_id}/image/{image_id}/delete")
    def delete_ticket_image(
        request: Request,
        ticket_id: int,
        image_id: int,
        admin: str = Depends(require_admin),
        return_url: str = Form(""),
        csrf_token: str = Form(""),
    ) -> RedirectResponse:
        require_admin_csrf(request, csrf_token)
        image = fetch_ticket_image(image_id, ticket_id)
        if not image:
            raise HTTPException(status_code=404, detail="图片不存在")
        image_path = str(image.get("image_path") or "")
        physical_path = resolve_upload_path(image_filename(image_path))
        timestamp = now_text()
        with get_connection() as connection:
            connection.execute("DELETE FROM ticket_images WHERE id = ? AND ticket_id = ?", (image_id, ticket_id))
            connection.execute("UPDATE tickets SET updated_at = ? WHERE id = ?", (timestamp, ticket_id))
            connection.execute(
                """
                INSERT INTO ticket_logs (
                    ticket_id, action, old_status, new_status, old_assigned_to,
                    new_assigned_to, note, operator, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (ticket_id, "删除图片", None, None, None, None, image_path, admin, timestamp),
            )
        if physical_path:
            try:
                physical_path.unlink(missing_ok=True)
            except OSError:
                pass
        return RedirectResponse(url=build_ticket_detail_url(ticket_id, return_url, saved="1"), status_code=303)

    @app.post("/admin/ticket/{ticket_id}/file/{file_id}/delete")
    def delete_ticket_file(
        request: Request,
        ticket_id: int,
        file_id: int,
        admin: str = Depends(require_admin),
        return_url: str = Form(""),
        csrf_token: str = Form(""),
    ) -> RedirectResponse:
        require_admin_csrf(request, csrf_token)
        ticket_file = fetch_ticket_file(file_id, ticket_id)
        if not ticket_file:
            raise HTTPException(status_code=404, detail="文件不存在")
        stored_filename = safe_uploaded_name(str(ticket_file.get("stored_filename") or ""))
        physical_path = resolve_upload_path(stored_filename)
        original_filename = str(ticket_file.get("original_filename") or "")
        timestamp = now_text()
        with get_connection() as connection:
            connection.execute("DELETE FROM ticket_files WHERE id = ? AND ticket_id = ?", (file_id, ticket_id))
            connection.execute("UPDATE tickets SET updated_at = ? WHERE id = ?", (timestamp, ticket_id))
            connection.execute(
                """
                INSERT INTO ticket_logs (
                    ticket_id, action, old_status, new_status, old_assigned_to,
                    new_assigned_to, note, operator, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (ticket_id, "删除文件", None, None, None, None, original_filename, admin, timestamp),
            )
        if physical_path:
            try:
                physical_path.unlink(missing_ok=True)
            except OSError:
                pass
        return RedirectResponse(url=build_ticket_detail_url(ticket_id, return_url, saved="1"), status_code=303)

    @app.get("/admin/uploads/{filename:path}")
    def protected_upload(filename: str, _admin: str = Depends(require_admin)) -> FileResponse:
        target = resolve_upload_file(filename)
        if not target:
            raise HTTPException(status_code=404, detail="图片不存在")
        return FileResponse(target)

    @app.get("/admin/files/{file_id}")
    def protected_file(file_id: int, _admin: str = Depends(require_admin)) -> FileResponse:
        ticket_file = fetch_ticket_file(file_id)
        if not ticket_file:
            raise HTTPException(status_code=404, detail="文件不存在")
        target = resolve_upload_file(safe_uploaded_name(str(ticket_file.get("stored_filename") or "")))
        if not target:
            raise HTTPException(status_code=404, detail="文件不存在")
        return FileResponse(
            target,
            media_type="application/octet-stream",
            filename=str(ticket_file.get("original_filename") or target.name),
            content_disposition_type="attachment",
        )

    @app.get("/admin/export")
    def export_tickets(
        _admin: str = Depends(require_admin),
        store_name: str = Query(""),
        request_type: str = Query(""),
        urgency: str = Query(""),
        status: str = Query(""),
        assigned_to: str = Query(""),
        date_start: str = Query(""),
        date_end: str = Query(""),
        keyword: str = Query(""),
        due_status: str = Query(""),
        sort: str = Query("newest"),
    ) -> StreamingResponse:
        filters = {
            "store_name": store_name,
            "request_type": request_type,
            "urgency": urgency,
            "status": status,
            "assigned_to": assigned_to,
            "date_start": date_start,
            "date_end": date_end,
            "keyword": keyword,
            "due_status": due_status,
        }
        config = load_app_config()
        output = build_excel(fetch_tickets(filters, sort, config))
        filename = f"{config.excel_filename_prefix}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
        )

    return app


app = create_app()
