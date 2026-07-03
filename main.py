import json
import os
import secrets
import sqlite3
import uuid
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
from urllib.parse import quote, urlencode

from fastapi import Depends, FastAPI, File, Form, HTTPException, Query, Request, UploadFile, status as http_status
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
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
admin_security = HTTPBasic()

DEFAULT_STORES = ["南京门东店", "南昌万寿宫店", "山城巷店", "东郊记忆店", "蟠龙天地店", "秀水街店", "湾里店", "下浩里店", "烟台山店"]
DEFAULT_REQUEST_TYPES = ["建单需求", "审单需求", "商品异常", "缺货需求", "新品需求", "系统问题", "其他"]
DEFAULT_URGENCY_LEVELS = ["普通", "加急", "当天必须处理"]
DEFAULT_STATUSES = ["待处理", "处理中", "待门店补充", "已完成", "已驳回"]
DEFAULT_BRANDS: List[str] = []
DEFAULT_HANDLERS = ["总部商品", "总部运营", "采购", "财务"]
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
}
COMPLETED_STATUS = "已完成"
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

    @property
    def max_image_bytes(self) -> int:
        return self.max_image_mb * 1024 * 1024

    @property
    def max_total_upload_bytes(self) -> int:
        return self.max_total_upload_mb * 1024 * 1024

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
    "期望完成时间",
    "当前状态",
    "处理人",
    "处理备注",
    "完成时间",
    "最后更新时间",
    "处理时长小时",
    "是否超时",
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


def get_admin_credentials() -> Tuple[str, str]:
    username = os.environ.get("ADMIN_USERNAME", "").strip()
    password = os.environ.get("ADMIN_PASSWORD", "")
    if not username or not password:
        raise HTTPException(status_code=503, detail="Admin credentials are not configured.")
    return username, password


def require_admin(credentials: HTTPBasicCredentials = Depends(admin_security)) -> str:
    expected_username, expected_password = get_admin_credentials()
    username_ok = secrets.compare_digest(credentials.username, expected_username)
    password_ok = secrets.compare_digest(credentials.password, expected_password)
    if not (username_ok and password_ok):
        raise HTTPException(
            status_code=http_status.HTTP_401_UNAUTHORIZED,
            detail="Invalid admin credentials.",
            headers={"WWW-Authenticate": "Basic"},
        )
    return credentials.username


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
            CREATE TABLE IF NOT EXISTS ticket_images (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_id INTEGER NOT NULL,
                image_path TEXT NOT NULL,
                uploaded_at TEXT NOT NULL,
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
        connection.execute("CREATE INDEX IF NOT EXISTS idx_tickets_created_at ON tickets(created_at)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_tickets_status ON tickets(status)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_tickets_assigned_to ON tickets(assigned_to)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_ticket_images_ticket_id ON ticket_images(ticket_id)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_ticket_logs_ticket_id ON ticket_logs(ticket_id)")


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


def load_system_config(statuses: List[str]) -> Dict[str, object]:
    raw_system = load_json_file("system.json", DEFAULT_SYSTEM)
    system = dict(DEFAULT_SYSTEM)
    if isinstance(raw_system, dict):
        system.update(raw_system)

    for key in ("max_image_mb", "page_size", "max_image_count", "max_total_upload_mb"):
        system[key] = positive_int_config(system, key)

    allowed_extensions = clean_string_list(
        system.get("allowed_image_extensions"),
        list(DEFAULT_SYSTEM["allowed_image_extensions"]),
    )
    system["allowed_image_extensions"] = [extension.lower().lstrip(".") for extension in allowed_extensions]

    default_status = str(system.get("default_status", "")).strip()
    system["default_status"] = default_status if default_status in statuses else statuses[0]

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
    )


def load_stores() -> List[str]:
    return load_app_config().stores


def compact_text(value: Optional[str], max_len: int = 36) -> str:
    text = (value or "").strip()
    if len(text) <= max_len:
        return text
    return text[:max_len] + "..."


def image_filename(image_path: str) -> str:
    return Path(str(image_path).replace("\\", "/")).name


def protected_upload_url(image_path: str) -> str:
    return f"/admin/uploads/{quote(image_filename(image_path))}"


def resolve_upload_file(filename: str) -> Optional[Path]:
    normalized = filename.replace("\\", "/")
    if not normalized or "/" in normalized or normalized in {".", ".."}:
        return None
    upload_root = get_upload_dir().resolve()
    target = (upload_root / normalized).resolve()
    try:
        target.relative_to(upload_root)
    except ValueError:
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


def build_query_params(filters: Dict[str, str], sort: str, page: Optional[int] = None) -> str:
    params = {key: value for key, value in filters.items() if str(value or "").strip()}
    if sort:
        params["sort"] = sort
    if page is not None:
        params["page"] = str(page)
    return urlencode(params)


def validate_submission(
    store_name: str,
    submitter: str,
    request_type: str,
    urgency: str,
    quantity: str,
    description: str,
    stores: Iterable[str],
    request_types: Iterable[str],
    urgency_levels: Iterable[str],
) -> Optional[str]:
    if not store_name or store_name not in stores:
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


async def prepare_images(images: Optional[List[UploadFile]], config: AppConfig) -> List[Tuple[str, bytes]]:
    raw_images: List[Tuple[str, bytes]] = []
    for image in images or []:
        if not image or not image.filename:
            continue
        original_name = Path(image.filename).name
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


def save_images(ticket_no: str, prepared_images: List[Tuple[str, bytes]]) -> Tuple[List[str], List[Path]]:
    image_paths: List[str] = []
    saved_files: List[Path] = []
    upload_dir = get_upload_dir()
    upload_dir.mkdir(parents=True, exist_ok=True)
    for index, (extension, content) in enumerate(prepared_images, start=1):
        filename = f"{ticket_no}_{index}_{uuid.uuid4().hex[:8]}.{extension}"
        target_path = upload_dir / filename
        target_path.write_bytes(content)
        saved_files.append(target_path)
        image_paths.append(f"uploads/{filename}")
    return image_paths, saved_files


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
) -> Tuple[int, str]:
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
                image_paths, saved_files = save_images(ticket_no, prepared_images)
                for image_path in image_paths:
                    connection.execute(
                        "INSERT INTO ticket_images (ticket_id, image_path, uploaded_at) VALUES (?, ?, ?)",
                        (ticket_id, image_path, timestamp),
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
        "store_name": "store_name",
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
    where_sql, params = build_ticket_where(filters)
    with get_connection() as connection:
        row = connection.execute(f"SELECT COUNT(*) AS total FROM tickets {where_sql}", params).fetchone()
    return int(row["total"] or 0)


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
                handler_note, closed_at, COUNT(ticket_images.id) AS image_count
            FROM tickets
            LEFT JOIN ticket_images ON ticket_images.ticket_id = tickets.id
            {where_sql}
            GROUP BY tickets.id
            {order_sql}
            {limit_sql}
            """,
            query_params,
        ).fetchall()
    tickets = [dict(row) for row in rows]
    for ticket in tickets:
        ticket["description_summary"] = compact_text(str(ticket.get("description") or ""))
    return tickets


def fetch_ticket_page(filters: Dict[str, str], sort: str, config: AppConfig, page: int) -> Tuple[List[Dict[str, object]], Dict[str, int]]:
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
    return dict(row) if row else None


def fetch_ticket_images(ticket_id: int) -> List[Dict[str, object]]:
    with get_connection() as connection:
        rows = connection.execute(
            "SELECT id, ticket_id, image_path, uploaded_at FROM ticket_images WHERE ticket_id = ? ORDER BY id",
            (ticket_id,),
        ).fetchall()
    images = [dict(row) for row in rows]
    for image in images:
        image["filename"] = image_filename(str(image.get("image_path") or ""))
        image["protected_url"] = protected_upload_url(str(image.get("image_path") or ""))
    return images


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
    if tickets:
        ids = [int(ticket["id"]) for ticket in tickets]
        placeholders = ",".join("?" for _ in ids)
        with get_connection() as connection:
            rows = connection.execute(
                f"SELECT ticket_id, image_path FROM ticket_images WHERE ticket_id IN ({placeholders}) ORDER BY id",
                ids,
            ).fetchall()
        for row in rows:
            image_map.setdefault(int(row["ticket_id"]), []).append(protected_upload_url(str(row["image_path"])))

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
                ticket.get("expected_finish_date") or "",
                ticket.get("status"),
                ticket.get("assigned_to") or "",
                ticket.get("handler_note") or "",
                ticket.get("closed_at") or "",
                ticket.get("updated_at"),
                processing_hours(ticket),
                overdue_text(ticket),
            ]
        )

    widths = [22, 20, 16, 14, 14, 14, 16, 20, 20, 10, 38, 42, 18, 14, 14, 32, 20, 20, 16, 12]
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
    ensure_directories()
    init_db()

    app = FastAPI(title=load_app_config().app_name)
    templates = Jinja2Templates(directory=str(TEMPLATES_DIR))
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

    def render_submit_form(
        request: Request,
        status_code: int = 200,
        error: str = "",
        values: Optional[Dict[str, str]] = None,
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
                "max_image_count": config.max_image_count,
                "max_total_upload_mb": config.max_total_upload_mb,
                "max_image_mb": config.max_image_mb,
                "error": error,
                "values": values or {},
            },
            status_code=status_code,
        )

    @app.get("/", include_in_schema=False)
    def root() -> RedirectResponse:
        return RedirectResponse(url="/submit", status_code=303)

    @app.get("/submit", response_class=HTMLResponse)
    def submit_page(request: Request) -> HTMLResponse:
        return render_submit_form(request)

    @app.post("/submit", response_class=HTMLResponse)
    async def submit_ticket(
        request: Request,
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
    ) -> HTMLResponse:
        config = load_app_config()
        stores = config.stores
        form_values = {
            "store_name": store_name,
            "submitter": submitter,
            "request_type": request_type,
            "urgency": urgency,
            "brand": brand,
            "product_name": product_name,
            "sku_barcode": sku_barcode,
            "quantity": quantity,
            "description": description,
            "expected_finish_date": expected_finish_date,
        }
        error = validate_submission(
            store_name,
            submitter,
            request_type,
            urgency,
            quantity,
            description,
            stores,
            config.request_types,
            config.urgency_levels,
        )
        if error:
            return render_submit_form(request, status_code=400, error=error, values=form_values)

        try:
            prepared_images = await prepare_images(images, config)
        except ValueError as exc:
            return render_submit_form(request, status_code=400, error=str(exc), values=form_values)

        timestamp = now_text()
        quantity_value = int(quantity.strip()) if quantity.strip() else None
        try:
            _ticket_id, ticket_no = create_ticket_with_images(
                {
                    "created_at": timestamp,
                    "store_name": store_name.strip(),
                    "submitter": submitter.strip(),
                    "request_type": request_type,
                    "urgency": urgency,
                    "brand": brand.strip(),
                    "product_name": product_name.strip(),
                    "sku_barcode": sku_barcode.strip(),
                    "quantity": quantity_value,
                    "description": description.strip(),
                    "expected_finish_date": expected_finish_date.strip(),
                },
                prepared_images,
                config,
            )
        except RuntimeError as exc:
            return render_submit_form(request, status_code=500, error=str(exc), values=form_values)
        except OSError:
            return render_submit_form(request, status_code=500, error="图片保存失败，请稍后重试。", values=form_values)

        return templates.TemplateResponse(
            request,
            "submit_success.html",
            {"request": request, "ticket_no": ticket_no},
        )

    @app.get("/admin", response_class=HTMLResponse)
    def admin_page(
        request: Request,
        _admin: str = Depends(require_admin),
        store_name: str = Query(""),
        request_type: str = Query(""),
        urgency: str = Query(""),
        status: str = Query(""),
        assigned_to: str = Query(""),
        date_start: str = Query(""),
        date_end: str = Query(""),
        keyword: str = Query(""),
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
        }
        config = load_app_config()
        tickets, pagination = fetch_ticket_page(filters, sort, config, page)
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
                "filters": filters,
                "sort": sort,
                "export_url": export_url,
                "pagination": pagination,
                "prev_page_url": "/admin" + (f"?{prev_query}" if prev_query else ""),
                "next_page_url": "/admin" + (f"?{next_query}" if next_query else ""),
            },
        )

    @app.get("/admin/ticket/{ticket_id}", response_class=HTMLResponse)
    def ticket_detail(
        request: Request,
        ticket_id: int,
        saved: str = Query(""),
        _admin: str = Depends(require_admin),
    ) -> HTMLResponse:
        ticket = fetch_ticket(ticket_id)
        if not ticket:
            raise HTTPException(status_code=404, detail="工单不存在")
        images = fetch_ticket_images(ticket_id)
        logs = fetch_ticket_logs(ticket_id)
        config = load_app_config()
        return templates.TemplateResponse(
            request,
            "ticket_detail.html",
            {
                "request": request,
                "ticket": ticket,
                "images": images,
                "statuses": config.statuses,
                "handlers": config.handlers,
                "logs": logs,
                "saved": saved,
            },
        )

    @app.post("/admin/ticket/{ticket_id}")
    def update_ticket(
        ticket_id: int,
        admin: str = Depends(require_admin),
        status: str = Form(""),
        assigned_to: str = Form(""),
        handler_note: str = Form(""),
    ) -> RedirectResponse:
        config = load_app_config()
        if status not in config.statuses:
            raise HTTPException(status_code=400, detail="状态不正确")
        assigned_to = assigned_to.strip()
        if assigned_to and config.handlers and assigned_to not in config.handlers:
            raise HTTPException(status_code=400, detail="处理人不正确")
        handler_note = handler_note.strip()
        timestamp = now_text()
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
        return RedirectResponse(url=f"/admin/ticket/{ticket_id}?saved=1", status_code=303)

    @app.get("/admin/uploads/{filename:path}")
    def protected_upload(filename: str, _admin: str = Depends(require_admin)) -> FileResponse:
        target = resolve_upload_file(filename)
        if not target:
            raise HTTPException(status_code=404, detail="图片不存在")
        return FileResponse(target)

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
