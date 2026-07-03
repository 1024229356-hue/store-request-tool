import importlib
import json
import re
import sys
from io import BytesIO
from pathlib import Path

from fastapi.testclient import TestClient
from openpyxl import load_workbook


PROJECT_DIR = Path(__file__).resolve().parents[1]
ADMIN_AUTH = ("regional-admin", "very-secret-value")

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
    },
}


def write_config(config_dir, overrides=None):
    config_dir.mkdir(parents=True, exist_ok=True)
    config = dict(DEFAULT_TEST_CONFIG)
    config.update(overrides or {})
    for filename, value in config.items():
        (config_dir / filename).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def build_client(tmp_path, monkeypatch, config_overrides=None, write_configs=True):
    config_dir = tmp_path / "config"
    if write_configs:
        write_config(config_dir, config_overrides)
    monkeypatch.setenv("STORE_REQUEST_DB_PATH", str(tmp_path / "tickets.db"))
    monkeypatch.setenv("STORE_REQUEST_UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setenv("STORE_REQUEST_CONFIG_DIR", str(config_dir))
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
    assert client.post("/admin/ticket/1", data={"status": "处理中"}).status_code == 401
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
        files=[("images", ("large.png", b"\x89PNG\r\n\x1a\n" + b"0" * (1024 * 1024 + 1), "image/png"))],
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
    client, _ = build_client(tmp_path, monkeypatch, write_configs=False)
    submit_page = client.get("/submit")
    assert submit_page.status_code == 200
    assert "南京门东店" in submit_page.text
    assert "建单需求" in submit_page.text

    config_dir = tmp_path / "config"
    config_dir.mkdir(parents=True, exist_ok=True)
    (config_dir / "stores.json").write_text("{bad json", encoding="utf-8")
    sys.modules.pop("main", None)
    main = importlib.import_module("main")
    invalid_config_client = TestClient(main.app)

    invalid_submit_page = invalid_config_client.get("/submit")
    assert invalid_submit_page.status_code == 200
    assert "南京门东店" in invalid_submit_page.text
