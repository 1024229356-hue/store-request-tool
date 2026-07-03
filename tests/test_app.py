import base64
import importlib
import json
import re
import sqlite3
import sys
from io import BytesIO
from pathlib import Path
from urllib.parse import quote

from fastapi.testclient import TestClient
from openpyxl import load_workbook


PROJECT_DIR = Path(__file__).resolve().parents[1]
ADMIN_AUTH = ("regional-admin", "very-secret-value")

VALID_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4//8/AAX+Av4N70a4AAAAAElFTkSuQmCC"
)

DEFAULT_TEST_CONFIG = {
    "stores.json": ["南京门东店", "南昌万寿宫店", "山城巷店"],
    "request_types.json": ["建单需求", "审单需求", "商品异常", "缺货需求", "新品需求", "系统问题", "其他"],
    "urgency_levels.json": ["普通", "加急", "当天必须处理"],
    "statuses.json": ["待处理", "处理中", "待门店补充", "已完成", "已驳回"],
    "brands.json": [],
    "handlers.json": ["总部商品", "总部运营", "采购", "财务"],
    "system.json": {
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
    },
}

EXPECTED_EXPORT_HEADERS = [
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
]


def write_config(config_dir, overrides=None):
    config_dir.mkdir(parents=True, exist_ok=True)
    config = dict(DEFAULT_TEST_CONFIG)
    config.update(overrides or {})
    for filename, value in config.items():
        (config_dir / filename).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def build_client(
    tmp_path,
    monkeypatch,
    config_overrides=None,
    write_configs=True,
    admin_auth=ADMIN_AUTH,
    admin_users=None,
):
    config_dir = tmp_path / "config"
    if write_configs:
        write_config(config_dir, config_overrides)
    monkeypatch.setenv("STORE_REQUEST_DB_PATH", str(tmp_path / "tickets.db"))
    monkeypatch.setenv("STORE_REQUEST_UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setenv("STORE_REQUEST_CONFIG_DIR", str(config_dir))
    block_dotenv_admin_users = admin_users is None
    if admin_users is None:
        monkeypatch.setenv("ADMIN_USERS", "")
    else:
        monkeypatch.setenv("ADMIN_USERS", admin_users)
    monkeypatch.setenv("ADMIN_USERNAME", admin_auth[0])
    monkeypatch.setenv("ADMIN_PASSWORD", admin_auth[1])
    monkeypatch.syspath_prepend(str(PROJECT_DIR))
    sys.modules.pop("main", None)
    main = importlib.import_module("main")
    if block_dotenv_admin_users:
        monkeypatch.delenv("ADMIN_USERS", raising=False)
    return TestClient(main.app), main


def submit_ticket(client, **overrides):
    data = {
        "store_name": "南京门东店",
        "submitter": "测试提报人",
        "request_type": "建单需求",
        "urgency": "加急",
        "brand": "测试品牌",
        "product_name": "测试商品",
        "sku_barcode": "690000000001",
        "quantity": "12",
        "description": "这是一条自动化测试工单",
        "expected_finish_date": "2026-07-04",
    }
    data.update(overrides)
    return client.post("/submit", data=data)


def submit_ticket_with_image(client, filename="issue.png", content=VALID_PNG, **overrides):
    data = {
        "store_name": "南京门东店",
        "submitter": "测试提报人",
        "request_type": "建单需求",
        "urgency": "加急",
        "brand": "测试品牌",
        "product_name": "测试商品",
        "sku_barcode": "690000000001",
        "quantity": "12",
        "description": "这是一条自动化测试工单",
        "expected_finish_date": "2026-07-04",
    }
    data.update(overrides)
    media_type = "image/png" if filename.endswith(".png") else "image/jpeg"
    return client.post("/submit", data=data, files=[("images", (filename, content, media_type))])


def submit_ticket_with_file(client, filename="need-list.xlsx", content=b"sku,qty\nA001,1\n", **overrides):
    data = {
        "store_name": "南京门东店",
        "submitter": "测试提报人",
        "request_type": "建单需求",
        "urgency": "加急",
        "brand": "测试品牌",
        "product_name": "测试商品",
        "sku_barcode": "690000000001",
        "quantity": "12",
        "description": "这是一条带普通附件的自动化测试工单",
        "expected_finish_date": "2026-07-04",
    }
    data.update(overrides)
    return client.post(
        "/submit",
        data=data,
        files=[("files", (filename, content, "application/octet-stream"))],
    )


def submit_ticket_with_image_and_file(
    client,
    image_filename="issue.png",
    file_filename="need-list.xlsx",
    file_content=b"sku,qty\nA001,1\n",
    **overrides,
):
    data = {
        "store_name": "南京门东店",
        "submitter": "测试提报人",
        "request_type": "建单需求",
        "urgency": "加急",
        "brand": "测试品牌",
        "product_name": "测试商品",
        "sku_barcode": "690000000001",
        "quantity": "12",
        "description": "这是一条同时带图片和普通附件的自动化测试工单",
        "expected_finish_date": "2026-07-04",
    }
    data.update(overrides)
    return client.post(
        "/submit",
        data=data,
        files=[
            ("images", (image_filename, VALID_PNG, "image/png")),
            ("files", (file_filename, file_content, "application/octet-stream")),
        ],
    )


def rows_for(tmp_path, table_name):
    with sqlite3.connect(tmp_path / "tickets.db") as connection:
        connection.row_factory = sqlite3.Row
        return [dict(row) for row in connection.execute(f"SELECT * FROM {table_name} ORDER BY id")]


def test_submit_page_exposes_image_and_file_upload_inputs(tmp_path, monkeypatch):
    client, _ = build_client(tmp_path, monkeypatch)

    submit_page = client.get("/submit")

    assert submit_page.status_code == 200
    assert 'name="images"' in submit_page.text
    assert "data-image-input" in submit_page.text
    assert "data-image-preview" in submit_page.text
    assert 'type="button"' in submit_page.text
    assert "data-clear-images" in submit_page.text
    assert 'name="files"' in submit_page.text
    assert "data-file-input" in submit_page.text
    assert "data-file-preview" in submit_page.text
    assert "data-clear-files" in submit_page.text
    assert "文件上传" in submit_page.text
    assert '/static/app.js?v=20260704' in submit_page.text


def test_upload_script_uses_independent_state_and_clear_selectors():
    script = (PROJECT_DIR / "static" / "app.js").read_text(encoding="utf-8")

    assert "DOMContentLoaded" in script
    assert "selectedImages" in script
    assert "selectedFiles" in script
    assert "[data-clear-images]" in script
    assert "[data-clear-files]" in script
    assert "event.preventDefault()" in script
    assert "event.stopPropagation()" in script
    assert "new DataTransfer()" in script


def test_admin_users_allows_multiple_accounts_and_logs_actual_operator(tmp_path, monkeypatch):
    client, _ = build_client(
        tmp_path,
        monkeypatch,
        admin_auth=("legacy-admin", "legacy-secret"),
        admin_users=" admin : 123456 , caigou : 123456 , baduser: , :badpass ",
    )

    assert client.get("/admin", auth=("admin", "123456")).status_code == 200
    assert client.get("/admin", auth=("caigou", "123456")).status_code == 200
    assert client.get("/admin", auth=("caigou", "wrong-password")).status_code == 401
    assert client.get("/admin", auth=("legacy-admin", "legacy-secret")).status_code == 401

    response = submit_ticket(client)
    assert response.status_code == 200

    update_response = client.post(
        "/admin/ticket/1",
        data={"status": "处理中", "assigned_to": "采购", "handler_note": "采购账号处理"},
        auth=("caigou", "123456"),
        follow_redirects=False,
    )
    assert update_response.status_code == 303

    logs = rows_for(tmp_path, "ticket_logs")
    assert len(logs) == 1
    assert logs[0]["operator"] == "caigou"


def test_legacy_admin_credentials_still_work_when_admin_users_missing(tmp_path, monkeypatch):
    client, _ = build_client(tmp_path, monkeypatch, admin_auth=("legacy-admin", "legacy-secret"))

    assert client.get("/admin", auth=("legacy-admin", "legacy-secret")).status_code == 200
    assert client.get("/admin", auth=("legacy-admin", "wrong-password")).status_code == 401


def test_file_upload_download_detail_admin_and_export(tmp_path, monkeypatch):
    client, _ = build_client(tmp_path, monkeypatch)
    file_content = b"sku,qty\nA001,1\n"

    response = submit_ticket_with_file(client, filename="采购清单.xlsx", content=file_content)
    assert response.status_code == 200

    ticket_files = rows_for(tmp_path, "ticket_files")
    assert len(ticket_files) == 1
    ticket_file = ticket_files[0]
    assert ticket_file["original_filename"] == "采购清单.xlsx"
    assert ticket_file["file_ext"] == "xlsx"
    assert ticket_file["file_size"] == len(file_content)
    assert ticket_file["stored_filename"].startswith("FILE-REQ-")
    assert ticket_file["stored_filename"].endswith(".xlsx")
    assert ticket_file["stored_filename"] != "采购清单.xlsx"
    assert ticket_file["file_path"].startswith("uploads/")
    assert (tmp_path / ticket_file["file_path"]).read_bytes() == file_content

    assert client.get("/admin/files/1").status_code == 401
    download_response = client.get("/admin/files/1", auth=ADMIN_AUTH)
    assert download_response.status_code == 200
    assert download_response.content == file_content
    assert download_response.headers["content-disposition"].lower().startswith("attachment")
    assert "%E9%87%87%E8%B4%AD%E6%B8%85%E5%8D%95.xlsx" in download_response.headers["content-disposition"]
    assert client.get("/files/1").status_code == 404

    detail_page = client.get("/admin/ticket/1", auth=ADMIN_AUTH)
    assert detail_page.status_code == 200
    assert "文件附件" in detail_page.text
    assert "采购清单.xlsx" in detail_page.text
    assert "/admin/files/1" in detail_page.text

    admin_page = client.get("/admin", auth=ADMIN_AUTH)
    assert admin_page.status_code == 200
    assert "图片 0 / 文件 1" in admin_page.text

    export_response = client.get("/admin/export", auth=ADMIN_AUTH)
    workbook = load_workbook(BytesIO(export_response.content))
    sheet = workbook.active
    assert [cell.value for cell in sheet[1]] == EXPECTED_EXPORT_HEADERS
    assert sheet["M2"].value == "采购清单.xlsx"
    assert sheet["N2"].value == "/admin/files/1"


def test_submit_admin_security_lifecycle_export_and_persistence(tmp_path, monkeypatch):
    client, main = build_client(tmp_path, monkeypatch)

    submit_page = client.get("/submit")
    assert submit_page.status_code == 200
    assert "南京门东店" in submit_page.text

    response = submit_ticket_with_image(client)
    assert response.status_code == 200
    assert "已提交" in response.text
    assert "请截图保存工单号" in response.text
    ticket_no = re.search(r"REQ-\d{8}-0001", response.text).group(0)

    tickets = rows_for(tmp_path, "tickets")
    assert tickets[0]["assigned_to"] in ("", None)
    assert tickets[0]["closed_at"] in ("", None)

    uploaded_files = list((tmp_path / "uploads").glob("*.png"))
    assert len(uploaded_files) == 1
    protected_path = f"/admin/uploads/{uploaded_files[0].name}"

    assert client.get(f"/uploads/{uploaded_files[0].name}").status_code != 200
    assert client.get(protected_path).status_code == 401
    protected_response = client.get(protected_path, auth=ADMIN_AUTH)
    assert protected_response.status_code == 200
    assert protected_response.content.startswith(b"\x89PNG")
    assert client.get("/admin/uploads/../tickets.db", auth=ADMIN_AUTH).status_code == 404

    assert client.get("/admin").status_code == 401
    assert client.get("/admin", auth=("admin", "change-me")).status_code == 401
    admin_page = client.get("/admin", auth=ADMIN_AUTH)
    assert admin_page.status_code == 200
    assert ticket_no in admin_page.text
    assert "处理人" in admin_page.text
    assert "图片数" in admin_page.text

    assert client.get("/admin/ticket/1").status_code == 401
    assert client.post("/admin/ticket/1", data={"status": "处理中"}).status_code == 401
    update_response = client.post(
        "/admin/ticket/1",
        data={"status": "已完成", "assigned_to": "采购", "handler_note": "已安排总部同事处理"},
        auth=ADMIN_AUTH,
        follow_redirects=False,
    )
    assert update_response.status_code == 303

    tickets = rows_for(tmp_path, "tickets")
    assert tickets[0]["status"] == "已完成"
    assert tickets[0]["assigned_to"] == "采购"
    assert tickets[0]["closed_at"]
    logs = rows_for(tmp_path, "ticket_logs")
    assert len(logs) == 1
    assert logs[0]["old_status"] == "待处理"
    assert logs[0]["new_status"] == "已完成"
    assert logs[0]["new_assigned_to"] == "采购"
    assert logs[0]["operator"] == ADMIN_AUTH[0]

    detail_page = client.get("/admin/ticket/1", auth=ADMIN_AUTH)
    assert detail_page.status_code == 200
    assert "已安排总部同事处理" in detail_page.text
    assert "完成时间" in detail_page.text
    assert "处理日志" in detail_page.text
    assert protected_path in detail_page.text

    reopen_response = client.post(
        "/admin/ticket/1",
        data={"status": "处理中", "assigned_to": "采购", "handler_note": "重新打开"},
        auth=ADMIN_AUTH,
        follow_redirects=False,
    )
    assert reopen_response.status_code == 303
    tickets = rows_for(tmp_path, "tickets")
    assert tickets[0]["closed_at"] in ("", None)
    assert len(rows_for(tmp_path, "ticket_logs")) == 2

    export_response = client.get("/admin/export", auth=ADMIN_AUTH)
    assert export_response.status_code == 200
    assert export_response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    workbook = load_workbook(BytesIO(export_response.content))
    sheet = workbook.active
    assert [cell.value for cell in sheet[1]] == EXPECTED_EXPORT_HEADERS
    assert sheet["A2"].value == ticket_no
    assert sheet["L2"].value == protected_path
    assert sheet["M2"].value in ("", None)
    assert sheet["N2"].value in ("", None)
    assert sheet["O2"].value == "2026-07-04"
    assert sheet["P2"].value == "处理中"
    assert sheet["Q2"].value == "采购"
    assert sheet["R2"].value == "重新打开"
    assert isinstance(sheet["U2"].value, (int, float))
    assert sheet["V2"].value in ("是", "否")

    monkeypatch.setenv("ADMIN_USERS", "")
    restarted_client = TestClient(main.create_app())
    monkeypatch.delenv("ADMIN_USERS", raising=False)
    restarted_admin_page = restarted_client.get("/admin", auth=ADMIN_AUTH)
    assert restarted_admin_page.status_code == 200
    assert ticket_no in restarted_admin_page.text


def test_upload_validation_rejects_fake_images_count_and_total_size(tmp_path, monkeypatch):
    client, _ = build_client(
        tmp_path,
        monkeypatch,
        {
            "system.json": {
                "app_name": "门店需求工单系统",
                "port": 8701,
                "max_image_mb": 10,
                "allowed_image_extensions": ["jpg", "jpeg", "png", "webp"],
                "default_status": "待处理",
                "excel_filename_prefix": "门店需求工单",
                "page_size": 50,
                "max_image_count": 1,
                "max_total_upload_mb": 1,
            }
        },
    )

    fake_jpg = submit_ticket_with_image(client, filename="fake.jpg", content=b"not an image")
    assert fake_jpg.status_code == 400
    assert "图片文件无法识别" in fake_jpg.text

    too_many = client.post(
        "/submit",
        data={
            "store_name": "南京门东店",
            "submitter": "测试提报人",
            "request_type": "建单需求",
            "urgency": "普通",
            "description": "超过图片数量限制",
        },
        files=[
            ("images", ("one.png", VALID_PNG, "image/png")),
            ("images", ("two.png", VALID_PNG, "image/png")),
        ],
    )
    assert too_many.status_code == 400
    assert "最多上传 1 张图片" in too_many.text

    too_large_total = submit_ticket_with_image(client, filename="large.png", content=VALID_PNG + b"x" * (1024 * 1024 + 1))
    assert too_large_total.status_code == 400
    assert "图片总大小不能超过 1MB" in too_large_total.text
    assert "REQ-" not in client.get("/admin", auth=ADMIN_AUTH).text


def test_file_upload_validation_rejects_extension_count_single_and_total_size(tmp_path, monkeypatch):
    limited_file_config = {
        "system.json": {
            "app_name": "门店需求工单系统",
            "port": 8701,
            "max_image_mb": 10,
            "allowed_image_extensions": ["jpg", "jpeg", "png", "webp"],
            "default_status": "待处理",
            "excel_filename_prefix": "门店需求工单",
            "page_size": 50,
            "max_image_count": 5,
            "max_total_upload_mb": 30,
            "allowed_file_extensions": ["txt"],
            "max_file_mb": 1,
            "max_file_count": 1,
            "max_total_file_upload_mb": 1,
        }
    }
    client, _ = build_client(tmp_path / "limited", monkeypatch, limited_file_config)

    executable = submit_ticket_with_file(client, filename="danger.exe", content=b"MZ")
    assert executable.status_code == 400
    assert "文件仅支持 txt 格式" in executable.text

    too_many = client.post(
        "/submit",
        data={
            "store_name": "南京门东店",
            "submitter": "测试提报人",
            "request_type": "建单需求",
            "urgency": "普通",
            "description": "超过文件数量限制",
        },
        files=[
            ("files", ("one.txt", b"one", "text/plain")),
            ("files", ("two.txt", b"two", "text/plain")),
        ],
    )
    assert too_many.status_code == 400
    assert "最多上传 1 个文件" in too_many.text

    single_size_config = dict(limited_file_config)
    single_size_config["system.json"] = dict(limited_file_config["system.json"], max_file_count=5, max_total_file_upload_mb=5)
    single_client, _ = build_client(tmp_path / "single", monkeypatch, single_size_config)
    too_large = submit_ticket_with_file(single_client, filename="large.txt", content=b"x" * (1024 * 1024 + 1))
    assert too_large.status_code == 400
    assert "单个文件不能超过 1MB" in too_large.text

    total_size_config = dict(limited_file_config)
    total_size_config["system.json"] = dict(limited_file_config["system.json"], max_file_mb=2, max_file_count=5)
    total_client, _ = build_client(tmp_path / "total", monkeypatch, total_size_config)
    too_large_total = total_client.post(
        "/submit",
        data={
            "store_name": "南京门东店",
            "submitter": "测试提报人",
            "request_type": "建单需求",
            "urgency": "普通",
            "description": "超过文件总大小限制",
        },
        files=[
            ("files", ("one.txt", b"x" * 600_000, "text/plain")),
            ("files", ("two.txt", b"x" * 600_000, "text/plain")),
        ],
    )
    assert too_large_total.status_code == 400
    assert "文件总大小不能超过 1MB" in too_large_total.text


def test_admin_can_delete_file_and_image_attachments_and_logs_actions(tmp_path, monkeypatch):
    client, _ = build_client(tmp_path, monkeypatch)
    response = submit_ticket_with_image_and_file(client)
    assert response.status_code == 200

    image = rows_for(tmp_path, "ticket_images")[0]
    ticket_file = rows_for(tmp_path, "ticket_files")[0]
    image_path = tmp_path / image["image_path"]
    file_path = tmp_path / ticket_file["file_path"]
    assert image_path.exists()
    assert file_path.exists()

    file_path.unlink()
    file_delete = client.post(
        f"/admin/ticket/1/file/{ticket_file['id']}/delete",
        auth=ADMIN_AUTH,
        follow_redirects=False,
    )
    assert file_delete.status_code == 303
    assert rows_for(tmp_path, "ticket_files") == []

    image_path.unlink()
    image_delete = client.post(
        f"/admin/ticket/1/image/{image['id']}/delete",
        auth=ADMIN_AUTH,
        follow_redirects=False,
    )
    assert image_delete.status_code == 303
    assert rows_for(tmp_path, "ticket_images") == []

    logs = rows_for(tmp_path, "ticket_logs")
    assert [log["action"] for log in logs] == ["删除文件", "删除图片"]
    assert all(log["operator"] == ADMIN_AUTH[0] for log in logs)

    detail_page = client.get("/admin/ticket/1", auth=ADMIN_AUTH)
    assert detail_page.status_code == 200
    assert "未上传图片" in detail_page.text
    assert "未上传文件附件" in detail_page.text
    assert "删除文件" in detail_page.text
    assert "删除图片" in detail_page.text


def test_admin_pagination_filters_handlers_and_keyword_search(tmp_path, monkeypatch):
    client, _ = build_client(
        tmp_path,
        monkeypatch,
        {
            "system.json": {
                "app_name": "门店需求工单系统",
                "port": 8701,
                "max_image_mb": 10,
                "allowed_image_extensions": ["jpg", "jpeg", "png", "webp"],
                "default_status": "待处理",
                "excel_filename_prefix": "门店需求工单",
                "page_size": 2,
                "max_image_count": 5,
                "max_total_upload_mb": 30,
            },
            "handlers.json": ["张三", "李四"],
        },
    )

    for index in range(3):
        response = submit_ticket(client, description=f"分页测试工单 {index}", product_name=f"测试商品{index}")
        assert response.status_code == 200

    client.post(
        "/admin/ticket/1",
        data={"status": "处理中", "assigned_to": "张三", "handler_note": "分配给张三"},
        auth=ADMIN_AUTH,
    )

    page_1 = client.get("/admin?page=1&sort=newest", auth=ADMIN_AUTH)
    assert page_1.status_code == 200
    assert "当前第 1 页" in page_1.text
    assert "共 2 页" in page_1.text
    assert "共 3 条工单" in page_1.text
    assert "sort=newest" in page_1.text
    assert "page=2" in page_1.text

    page_2 = client.get("/admin?page=2&sort=newest", auth=ADMIN_AUTH)
    assert page_2.status_code == 200
    assert "当前第 2 页" in page_2.text

    assigned_filter = client.get(f"/admin?assigned_to={quote('张三')}&keyword={quote('张三')}", auth=ADMIN_AUTH)
    assert assigned_filter.status_code == 200
    assert "分配给张三" in assigned_filter.text or "张三" in assigned_filter.text
    assert "共 1 条工单" in assigned_filter.text

    export_response = client.get("/admin/export?page=2", auth=ADMIN_AUTH)
    workbook = load_workbook(BytesIO(export_response.content))
    assert workbook.active.max_row == 4


def test_export_without_data_still_contains_headers(tmp_path, monkeypatch):
    client, _ = build_client(tmp_path, monkeypatch)

    assert client.get("/admin/export").status_code == 401

    response = client.get("/admin/export", auth=ADMIN_AUTH)
    assert response.status_code == 200
    workbook = load_workbook(BytesIO(response.content))
    sheet = workbook.active

    assert [cell.value for cell in sheet[1]] == EXPECTED_EXPORT_HEADERS
    assert sheet.max_row == 1


def test_legacy_database_is_migrated_without_losing_existing_rows(tmp_path, monkeypatch):
    config_dir = tmp_path / "config"
    write_config(config_dir)
    db_path = tmp_path / "tickets.db"
    with sqlite3.connect(db_path) as connection:
        connection.execute(
            """
            CREATE TABLE tickets (
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
            CREATE TABLE ticket_images (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_id INTEGER NOT NULL,
                image_path TEXT NOT NULL,
                uploaded_at TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            INSERT INTO tickets (
                ticket_no, created_at, updated_at, store_name, submitter, request_type, urgency,
                brand, product_name, sku_barcode, quantity, description, expected_finish_date, status, handler_note
            )
            VALUES (
                'REQ-20260703-0001', '2026-07-03 10:00:00', '2026-07-03 10:00:00',
                '南京门东店', '旧数据', '建单需求', '普通', '', '旧商品', '', 1,
                '旧库里的工单', '2026-07-04', '待处理', ''
            )
            """
        )
        connection.execute(
            "INSERT INTO ticket_images (ticket_id, image_path, uploaded_at) VALUES (1, 'uploads/old.png', '2026-07-03 10:00:00')"
        )

    monkeypatch.setenv("STORE_REQUEST_DB_PATH", str(db_path))
    monkeypatch.setenv("STORE_REQUEST_UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setenv("STORE_REQUEST_CONFIG_DIR", str(config_dir))
    monkeypatch.setenv("ADMIN_USERS", "")
    monkeypatch.setenv("ADMIN_USERNAME", ADMIN_AUTH[0])
    monkeypatch.setenv("ADMIN_PASSWORD", ADMIN_AUTH[1])
    monkeypatch.syspath_prepend(str(PROJECT_DIR))
    sys.modules.pop("main", None)
    main = importlib.import_module("main")
    monkeypatch.delenv("ADMIN_USERS", raising=False)
    client = TestClient(main.app)

    with sqlite3.connect(db_path) as connection:
        ticket_columns = {row[1] for row in connection.execute("PRAGMA table_info(tickets)").fetchall()}
        tables = {row[0] for row in connection.execute("SELECT name FROM sqlite_master WHERE type = 'table'").fetchall()}
        ticket_count = connection.execute("SELECT COUNT(*) FROM tickets").fetchone()[0]
        image_count = connection.execute("SELECT COUNT(*) FROM ticket_images").fetchone()[0]

    assert "assigned_to" in ticket_columns
    assert "closed_at" in ticket_columns
    assert "ticket_logs" in tables
    assert "ticket_files" in tables
    assert ticket_count == 1
    assert image_count == 1

    admin_page = client.get("/admin", auth=ADMIN_AUTH)
    assert admin_page.status_code == 200
    assert "旧库里的工单" in admin_page.text


def test_config_files_drive_options_validation_images_and_export_name(tmp_path, monkeypatch):
    client, _ = build_client(
        tmp_path,
        monkeypatch,
        {
            "stores.json": ["测试门店"],
            "request_types.json": ["陈列需求"],
            "urgency_levels.json": ["立刻处理"],
            "statuses.json": ["新建", "跟进中"],
            "brands.json": ["自有品牌"],
            "system.json": {
                "app_name": "配置化工单",
                "port": 8701,
                "max_image_mb": 1,
                "allowed_image_extensions": ["png"],
                "default_status": "新建",
                "excel_filename_prefix": "配置导出",
                "page_size": 25,
                "max_image_count": 2,
                "max_total_upload_mb": 3,
            },
        },
    )

    submit_page = client.get("/submit")
    assert submit_page.status_code == 200
    assert "测试门店" in submit_page.text
    assert "陈列需求" in submit_page.text
    assert "立刻处理" in submit_page.text
    assert "南京门东店" not in submit_page.text
    assert "建单需求" not in submit_page.text

    bad_request_type = client.post(
        "/submit",
        data={
            "store_name": "测试门店",
            "submitter": "配置测试",
            "request_type": "建单需求",
            "urgency": "立刻处理",
            "description": "配置文件校验应拒绝未配置的需求类型",
        },
    )
    assert bad_request_type.status_code == 400
    assert "请选择有效需求类型" in bad_request_type.text

    oversized_image = client.post(
        "/submit",
        data={
            "store_name": "测试门店",
            "submitter": "配置测试",
            "request_type": "陈列需求",
            "urgency": "立刻处理",
            "description": "图片大小限制来自 system.json",
        },
        files=[("images", ("large.png", VALID_PNG + b"0" * (1024 * 1024 + 1), "image/png"))],
    )
    assert oversized_image.status_code == 400
    assert "单张图片不能超过 1MB" in oversized_image.text

    response = client.post(
        "/submit",
        data={
            "store_name": "测试门店",
            "submitter": "配置测试",
            "request_type": "陈列需求",
            "urgency": "立刻处理",
            "brand": "自有品牌",
            "description": "配置化提交成功",
        },
    )
    assert response.status_code == 200

    admin_page = client.get("/admin?status=新建", auth=ADMIN_AUTH)
    assert admin_page.status_code == 200
    assert "新建" in admin_page.text
    assert "跟进中" in admin_page.text
    assert "配置化提交成功" in admin_page.text

    detail_page = client.get("/admin/ticket/1", auth=ADMIN_AUTH)
    assert detail_page.status_code == 200
    assert "跟进中" in detail_page.text

    export_response = client.get("/admin/export", auth=ADMIN_AUTH)
    assert export_response.status_code == 200
    assert "%E9%85%8D%E7%BD%AE%E5%AF%BC%E5%87%BA_" in export_response.headers["content-disposition"]


def test_missing_and_invalid_config_files_fall_back_to_defaults(tmp_path, monkeypatch):
    client, main = build_client(tmp_path, monkeypatch, write_configs=False)
    submit_page = client.get("/submit")
    assert submit_page.status_code == 200
    assert "南京门东店" in submit_page.text
    assert "建单需求" in submit_page.text
    assert main.load_app_config().page_size == 50
    assert main.load_app_config().max_image_count == 5
    assert main.load_app_config().max_total_upload_mb == 30
    assert main.load_app_config().allowed_file_extensions == ["pdf", "doc", "docx", "xls", "xlsx", "csv", "txt", "zip", "rar"]
    assert main.load_app_config().max_file_mb == 20
    assert main.load_app_config().max_file_count == 5
    assert main.load_app_config().max_total_file_upload_mb == 50

    config_dir = tmp_path / "config"
    config_dir.mkdir(parents=True, exist_ok=True)
    (config_dir / "stores.json").write_text("{bad json", encoding="utf-8")
    (config_dir / "system.json").write_text(
        json.dumps(
            {
                "page_size": -1,
                "max_image_count": "bad",
                "max_total_upload_mb": 0,
                "allowed_file_extensions": "bad",
                "max_file_mb": -1,
                "max_file_count": "bad",
                "max_total_file_upload_mb": 0,
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    sys.modules.pop("main", None)
    main = importlib.import_module("main")
    invalid_config_client = TestClient(main.app)

    invalid_submit_page = invalid_config_client.get("/submit")
    assert invalid_submit_page.status_code == 200
    assert "南京门东店" in invalid_submit_page.text
    assert main.load_app_config().page_size == 50
    assert main.load_app_config().max_image_count == 5
    assert main.load_app_config().max_total_upload_mb == 30
    assert main.load_app_config().allowed_file_extensions == ["pdf", "doc", "docx", "xls", "xlsx", "csv", "txt", "zip", "rar"]
    assert main.load_app_config().max_file_mb == 20
    assert main.load_app_config().max_file_count == 5
    assert main.load_app_config().max_total_file_upload_mb == 50
