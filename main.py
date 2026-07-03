import json
import os
import secrets
import sqlite3
import uuid
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
STORES_FILE = CONFIG_DIR / "stores.json"

REQUEST_TYPES = ["建单需求", "审单需求", "商品异常", "缺货需求", "新品需求", "系统问题", "其他"]
URGENCY_LEVELS = ["普通", "加急", "当天必须处理"]
STATUSES = ["待处理", "处理中", "待门店补充", "已完成", "已驳回"]
ALLOWED_IMAGE_EXTENSIONS = {"jpg", "jpeg", "png", "webp"}
MAX_IMAGE_BYTES = 10 * 1024 * 1024
ENV_FILE = BASE_DIR / ".env"
admin_security = HTTPBasic()

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


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def ensure_directories() -> None:
    get_db_path().parent.mkdir(parents=True, exist_ok=True)
    get_upload_dir().mkdir(parents=True, exist_ok=True)
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
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


def load_stores() -> List[str]:
    try:
        with STORES_FILE.open("r", encoding="utf-8") as file:
            stores = json.load(file)
        clean_stores = [str(store).strip() for store in stores if str(store).strip()]
        if clean_stores:
            return clean_stores
    except (OSError, json.JSONDecodeError):
        pass
    return ["南京门东店", "南昌万寿宫店", "山城巷店", "东郊记忆店", "蟠龙天地店", "秀水街店", "湾里店", "下浩里店", "烟台山店"]


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
) -> Optional[str]:
    if not store_name or store_name not in stores:
        return "请选择有效门店。"
    if not submitter.strip():
        return "请填写提报人。"
    if request_type not in REQUEST_TYPES:
        return "请选择有效需求类型。"
    if urgency not in URGENCY_LEVELS:
        return "请选择有效紧急程度。"
    if not description.strip():
        return "请填写问题说明。"
    if quantity.strip() and not quantity.strip().isdigit():
        return "数量只能填写数字。"
    return None


async def prepare_images(images: Optional[List[UploadFile]]) -> List[Tuple[str, bytes]]:
    prepared_images: List[Tuple[str, bytes]] = []
    for image in images or []:
        if not image or not image.filename:
            continue
        original_name = Path(image.filename).name
        extension = Path(original_name).suffix.lower().lstrip(".")
        if extension not in ALLOWED_IMAGE_EXTENSIONS:
            raise ValueError("图片仅支持 jpg、jpeg、png、webp 格式。")
        content = await image.read()
        if len(content) > MAX_IMAGE_BYTES:
            raise ValueError("单张图片不能超过 10MB。")
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


def build_order_sql(sort: str) -> str:
    if sort == "urgency":
        return """
        ORDER BY
            CASE urgency
                WHEN '当天必须处理' THEN 0
                WHEN '加急' THEN 1
                ELSE 2
            END,
            created_at DESC,
            id DESC
        """
    if sort == "status":
        return """
        ORDER BY
            CASE status
                WHEN '待处理' THEN 0
                WHEN '处理中' THEN 1
                WHEN '待门店补充' THEN 2
                WHEN '已驳回' THEN 3
                WHEN '已完成' THEN 4
                ELSE 9
            END,
            created_at DESC,
            id DESC
        """
    return "ORDER BY created_at DESC, id DESC"


def fetch_tickets(filters: Dict[str, str], sort: str) -> List[Dict[str, object]]:
    where_sql, params = build_ticket_where(filters)
    order_sql = build_order_sql(sort)
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

    app = FastAPI(title="门店需求工单系统")
    templates = Jinja2Templates(directory=str(TEMPLATES_DIR))
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
    app.mount("/uploads", StaticFiles(directory=str(get_upload_dir())), name="uploads")

    def render_submit_form(
        request: Request,
        status_code: int = 200,
        error: str = "",
        values: Optional[Dict[str, str]] = None,
    ) -> HTMLResponse:
        return templates.TemplateResponse(
            request,
            "submit.html",
            {
                "request": request,
                "stores": load_stores(),
                "request_types": REQUEST_TYPES,
                "urgency_levels": URGENCY_LEVELS,
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
        stores = load_stores()
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
        error = validate_submission(store_name, submitter, request_type, urgency, quantity, description, stores)
        if error:
            return render_submit_form(request, status_code=400, error=error, values=form_values)

        try:
            prepared_images = await prepare_images(images)
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
                    "待处理",
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
        tickets = fetch_tickets(filters, sort)
        export_url = "/admin/export"
        if request.url.query:
            export_url += f"?{request.url.query}"
        return templates.TemplateResponse(
            request,
            "admin.html",
            {
                "request": request,
                "tickets": tickets,
                "stores": load_stores(),
                "request_types": REQUEST_TYPES,
                "urgency_levels": URGENCY_LEVELS,
                "statuses": STATUSES,
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
        return templates.TemplateResponse(
            request,
            "ticket_detail.html",
            {
                "request": request,
                "ticket": ticket,
                "images": images,
                "statuses": STATUSES,
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
        if status not in STATUSES:
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
        output = build_excel(fetch_tickets(filters, sort))
        filename = f"门店需求工单_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
        )

    return app


app = create_app()
