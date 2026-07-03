import importlib
import re
import sys
from io import BytesIO
from pathlib import Path

from fastapi.testclient import TestClient
from openpyxl import load_workbook


PROJECT_DIR = Path(__file__).resolve().parents[1]
ADMIN_AUTH = ("regional-admin", "very-secret-value")


def build_client(tmp_path, monkeypatch):
    monkeypatch.setenv("STORE_REQUEST_DB_PATH", str(tmp_path / "tickets.db"))
    monkeypatch.setenv("STORE_REQUEST_UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setenv("ADMIN_USERNAME", ADMIN_AUTH[0])
    monkeypatch.setenv("ADMIN_PASSWORD", ADMIN_AUTH[1])
    monkeypatch.syspath_prepend(str(PROJECT_DIR))
    sys.modules.pop("main", None)
    main = importlib.import_module("main")
    return TestClient(main.app), main


def test_submit_ticket_with_image_admin_update_export_and_persistence(tmp_path, monkeypatch):
    client, main = build_client(tmp_path, monkeypatch)

    submit_page = client.get("/submit")
    assert submit_page.status_code == 200
    assert "南京门东店" in submit_page.text

    image_bytes = b"\x89PNG\r\n\x1a\n" + b"0" * 128
    response = client.post(
        "/submit",
        data={
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
        },
        files=[("images", ("issue.png", image_bytes, "image/png"))],
    )
    assert response.status_code == 200
    assert "已提交" in response.text
    ticket_no = re.search(r"REQ-\d{8}-0001", response.text).group(0)

    uploaded_files = list((tmp_path / "uploads").glob("*.png"))
    assert len(uploaded_files) == 1

    assert client.get("/admin").status_code == 401
    assert client.get("/admin", auth=("admin", "change-me")).status_code == 401
    admin_page = client.get("/admin", auth=ADMIN_AUTH)
    assert admin_page.status_code == 200
    assert ticket_no in admin_page.text
    assert "测试商品" in admin_page.text

    assert client.get("/admin/ticket/1").status_code == 401
    assert client.post("/admin/ticket/1", data={"status": main.STATUSES[1]}).status_code == 401
    update_response = client.post(
        "/admin/ticket/1",
        data={"status": "处理中", "handler_note": "已安排总部同事处理"},
        auth=ADMIN_AUTH,
        follow_redirects=False,
    )
    assert update_response.status_code == 303

    detail_page = client.get("/admin/ticket/1", auth=ADMIN_AUTH)
    assert detail_page.status_code == 200
    assert "处理中" in detail_page.text
    assert "已安排总部同事处理" in detail_page.text
    assert "/uploads/" in detail_page.text

    assert client.get("/admin/export").status_code == 401
    export_response = client.get("/admin/export", auth=ADMIN_AUTH)
    assert export_response.status_code == 200
    assert export_response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    workbook = load_workbook(BytesIO(export_response.content))
    sheet = workbook.active
    assert sheet["A2"].value == ticket_no
    assert sheet["K2"].value == "这是一条自动化测试工单"
    assert "uploads/" in sheet["L2"].value
    assert sheet["M2"].value == "处理中"
    assert sheet["N2"].value == "已安排总部同事处理"

    restarted_client = TestClient(main.create_app())
    restarted_admin_page = restarted_client.get("/admin", auth=ADMIN_AUTH)
    assert restarted_admin_page.status_code == 200
    assert ticket_no in restarted_admin_page.text


def test_submit_rejects_bad_quantity_and_bad_image_type(tmp_path, monkeypatch):
    client, _ = build_client(tmp_path, monkeypatch)

    bad_quantity = client.post(
        "/submit",
        data={
            "store_name": "南京门东店",
            "submitter": "测试提报人",
            "request_type": "建单需求",
            "urgency": "普通",
            "quantity": "12a",
            "description": "数量非法时不能提交",
        },
    )
    assert bad_quantity.status_code == 400
    assert "数量只能填写数字" in bad_quantity.text
    assert "REQ-" not in client.get("/admin", auth=ADMIN_AUTH).text

    bad_image = client.post(
        "/submit",
        data={
            "store_name": "南京门东店",
            "submitter": "测试提报人",
            "request_type": "建单需求",
            "urgency": "普通",
            "description": "图片格式非法时不能提交",
        },
        files=[("images", ("bad.gif", b"GIF89a", "image/gif"))],
    )
    assert bad_image.status_code == 400
    assert "图片仅支持" in bad_image.text
    assert "REQ-" not in client.get("/admin", auth=ADMIN_AUTH).text


def test_export_without_data_still_contains_headers(tmp_path, monkeypatch):
    client, _ = build_client(tmp_path, monkeypatch)

    assert client.get("/admin/export").status_code == 401

    response = client.get("/admin/export", auth=ADMIN_AUTH)
    assert response.status_code == 200
    workbook = load_workbook(BytesIO(response.content))
    sheet = workbook.active

    headers = [cell.value for cell in sheet[1]]
    assert headers == [
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
    assert sheet.max_row == 1
