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
from urllib.parse import quote

from fastapi import Depends, FastAPI, File, Form, HTTPException, Query, Request, UploadFile, status as http_status
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


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

    @property
    def max_image_bytes(self) -> int:
        return self.max_image_mb * 1024 * 1024

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
    "状态",
    "处理备注",
    "最后更新时间",
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
                handler_note TEXT
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
                FOREIGN KEY (ticket_id) REFERENCES tickets(id) ON DELETE CASCADE
            )
            """
        )
        connection.execute("CREATE INDEX IF NOT EXISTS idx_tickets_created_at ON tickets(created_at)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_tickets_status ON tickets(status)")
        connection.execute("CREATE INDEX IF NOT EXISTS idx_ticket_images_ticket_id ON ticket_images(ticket_id)")


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


def load_system_config(statuses: List[str]) -> Dict[str, object]:
    raw_system = load_json_file("system.json", DEFAULT_SYSTEM)
    system = dict(DEFAULT_SYSTEM)
    if isinstance(raw_system, dict):
        system.update(raw_system)

    max_image_mb = system.get("max_image_mb")
    if not isinstance(max_image_mb, int) or max_image_mb <= 0:
        system["max_image_mb"] = DEFAULT_SYSTEM["max_image_mb"]

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
    )


def load_stores() -> List[str]:
    return load_app_config().stores


def compact_text(value: Optional[str], max_len: int = 36) -> str:
    text = (value or "").strip()
    if len(text) <= max_len:
        return text
    return text[:max_len] + "..."


def generate_ticket_no(connection: sqlite3.Connection) -> str:
    today = datetime.now().strftime("%Y%m%d")
    prefix = f"REQ-{today}-"
    row = connection.execute(
        "SELECT MAX(CAST(SUBSTR(ticket_no, 14) AS INTEGER)) AS max_no FROM tickets WHERE ticket_no LIKE ?",
        (prefix + "%",),
    ).fetchone()
    next_no = (row["max_no"] or 0) + 1
    return f"{prefix}{next_no:04d}"


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
    prepared_images: List[Tuple[str, bytes]] = []
    allowed_extensions = set(config.allowed_image_extensions)
    for image in images or []:
        if not image or not image.filename:
            continue
        original_name = Path(image.filename).name
        extension = Path(original_name).suffix.lower().lstrip(".")
        if extension not in allowed_extensions:
            allowed_text = "、".join(config.allowed_image_extensions)
            raise ValueError(f"图片仅支持 {allowed_text} 格式。")
        content = await image.read()
        if len(content) > config.max_image_bytes:
            raise ValueError(f"单张图片不能超过 {config.max_image_mb}MB。")
        if not content:
            continue
        prepared_images.append((extension, content))
    return prepared_images


def save_images(ticket_id: int, ticket_no: str, prepared_images: List[Tuple[str, bytes]]) -> List[str]:
    image_paths: List[str] = []
    upload_dir = get_upload_dir()
    upload_dir.mkdir(parents=True, exist_ok=True)
    for index, (extension, content) in enumerate(prepared_images, start=1):
        filename = f"{ticket_no}_{index}_{uuid.uuid4().hex[:8]}.{extension}"
        target_path = upload_dir / filename
        target_path.write_bytes(content)
        image_paths.append(f"uploads/{filename}")
    return image_paths


def build_ticket_where(filters: Dict[str, str]) -> Tuple[str, List[str]]:
    clauses: List[str] = []
    params: List[str] = []

    exact_fields = {
        "store_name": "store_name",
        "request_type": "request_type",
        "urgency": "urgency",
        "status": "status",
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
                OR handler_note LIKE ?
            )
            """
        )
        params.extend([like_value] * 8)

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
            id DESC
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
            id DESC
        """.format(status_cases=status_cases, fallback_index=len(config.statuses))
    return "ORDER BY created_at DESC, id DESC"


def fetch_tickets(filters: Dict[str, str], sort: str, config: AppConfig) -> List[Dict[str, object]]:
    where_sql, params = build_ticket_where(filters)
    order_sql = build_order_sql(sort, config)
    with get_connection() as connection:
        rows = connection.execute(
            f"""
            SELECT
                id, ticket_no, created_at, updated_at, store_name, submitter,
                request_type, urgency, brand, product_name, sku_barcode,
                quantity, description, expected_finish_date, status, handler_note
            FROM tickets
            {where_sql}
            {order_sql}
            """,
            params,
        ).fetchall()
    tickets = [dict(row) for row in rows]
    for ticket in tickets:
        ticket["description_summary"] = compact_text(str(ticket.get("description") or ""))
    return tickets


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
            image_map.setdefault(int(row["ticket_id"]), []).append(str(row["image_path"]))

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
                ticket.get("status"),
                ticket.get("handler_note") or "",
                ticket.get("updated_at"),
            ]
        )

    widths = [22, 20, 16, 14, 14, 14, 16, 20, 20, 10, 38, 42, 14, 32, 20]
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
    app.mount("/uploads", StaticFiles(directory=str(get_upload_dir())), name="uploads")

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
        with get_connection() as connection:
            ticket_no = generate_ticket_no(connection)
            cursor = connection.execute(
                """
                INSERT INTO tickets (
                    ticket_no, created_at, updated_at, store_name, submitter,
                    request_type, urgency, brand, product_name, sku_barcode,
                    quantity, description, expected_finish_date, status, handler_note
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    ticket_no,
                    timestamp,
                    timestamp,
                    store_name.strip(),
                    submitter.strip(),
                    request_type,
                    urgency,
                    brand.strip(),
                    product_name.strip(),
                    sku_barcode.strip(),
                    quantity_value,
                    description.strip(),
                    expected_finish_date.strip(),
                    config.default_status,
                    "",
                ),
            )
            ticket_id = int(cursor.lastrowid)
            image_paths = save_images(ticket_id, ticket_no, prepared_images)
            for image_path in image_paths:
                connection.execute(
                    "INSERT INTO ticket_images (ticket_id, image_path, uploaded_at) VALUES (?, ?, ?)",
                    (ticket_id, image_path, timestamp),
                )

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
        date_start: str = Query(""),
        date_end: str = Query(""),
        keyword: str = Query(""),
        sort: str = Query("newest"),
    ) -> HTMLResponse:
        filters = {
            "store_name": store_name,
            "request_type": request_type,
            "urgency": urgency,
            "status": status,
            "date_start": date_start,
            "date_end": date_end,
            "keyword": keyword,
        }
        config = load_app_config()
        tickets = fetch_tickets(filters, sort, config)
        export_url = "/admin/export"
        if request.url.query:
            export_url += f"?{request.url.query}"
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
                "filters": filters,
                "sort": sort,
                "export_url": export_url,
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
        config = load_app_config()
        return templates.TemplateResponse(
            request,
            "ticket_detail.html",
            {
                "request": request,
                "ticket": ticket,
                "images": images,
                "statuses": config.statuses,
                "saved": saved,
            },
        )

    @app.post("/admin/ticket/{ticket_id}")
    def update_ticket(
        ticket_id: int,
        _admin: str = Depends(require_admin),
        status: str = Form(""),
        handler_note: str = Form(""),
    ) -> RedirectResponse:
        config = load_app_config()
        if status not in config.statuses:
            raise HTTPException(status_code=400, detail="状态不正确")
        with get_connection() as connection:
            cursor = connection.execute(
                "UPDATE tickets SET status = ?, handler_note = ?, updated_at = ? WHERE id = ?",
                (status, handler_note.strip(), now_text(), ticket_id),
            )
            if cursor.rowcount == 0:
                raise HTTPException(status_code=404, detail="工单不存在")
        return RedirectResponse(url=f"/admin/ticket/{ticket_id}?saved=1", status_code=303)

    @app.get("/admin/export")
    def export_tickets(
        _admin: str = Depends(require_admin),
        store_name: str = Query(""),
        request_type: str = Query(""),
        urgency: str = Query(""),
        status: str = Query(""),
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
