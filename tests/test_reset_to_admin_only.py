import importlib
import json
import sqlite3
import sys
from pathlib import Path

import pytest
from fastapi.testclient import TestClient


PROJECT_DIR = Path(__file__).resolve().parents[1]


def write_minimal_config(config_dir: Path) -> None:
    config_dir.mkdir(parents=True, exist_ok=True)
    (config_dir / "stores.json").write_text(json.dumps(["Reset Store"], ensure_ascii=False), encoding="utf-8")
    (config_dir / "brands.json").write_text(json.dumps(["Reset Brand"], ensure_ascii=False), encoding="utf-8")
    (config_dir / "request_types.json").write_text(json.dumps(["Reset Type"], ensure_ascii=False), encoding="utf-8")
    (config_dir / "holidays.json").write_text(json.dumps({"2026-07-08": "Reset Holiday"}, ensure_ascii=False), encoding="utf-8")


def import_main_and_reset(tmp_path, monkeypatch, admin_users="admin:123456,ops:123456"):
    config_dir = tmp_path / "config"
    write_minimal_config(config_dir)
    monkeypatch.setenv("STORE_REQUEST_DB_PATH", str(tmp_path / "tickets.db"))
    monkeypatch.setenv("STORE_REQUEST_UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setenv("STORE_REQUEST_CONFIG_DIR", str(config_dir))
    monkeypatch.setenv("STORE_REQUEST_RESET_REPORT_DIR", str(tmp_path / "reports"))
    monkeypatch.setenv("ADMIN_USERS", admin_users)
    monkeypatch.setenv("SESSION_SECRET", "reset-test-session-secret")
    monkeypatch.syspath_prepend(str(PROJECT_DIR))
    sys.modules.pop("main", None)
    sys.modules.pop("scripts.reset_to_admin_only", None)
    main = importlib.import_module("main")
    reset = importlib.import_module("scripts.reset_to_admin_only")
    return main, reset


def rows_for(db_path: Path, table_name: str):
    with sqlite3.connect(db_path) as connection:
        connection.row_factory = sqlite3.Row
        return [dict(row) for row in connection.execute(f"SELECT * FROM {table_name} ORDER BY id")]


def count_rows(db_path: Path, table_name: str) -> int:
    with sqlite3.connect(db_path) as connection:
        return int(connection.execute(f"SELECT COUNT(*) FROM {table_name}").fetchone()[0])


def fake_backup_runner_factory(tmp_path: Path):
    calls = []

    def fake_backup_runner(project_root: Path) -> Path:
        calls.append(project_root)
        backup_dir = tmp_path / "backups" / "20260708_120000"
        (backup_dir / "data").mkdir(parents=True, exist_ok=True)
        (backup_dir / "uploads").mkdir(parents=True, exist_ok=True)
        return backup_dir

    fake_backup_runner.calls = calls
    return fake_backup_runner


def seed_reset_history(tmp_path: Path, main_module):
    db_path = tmp_path / "tickets.db"
    timestamp = "2026-07-08 12:00:00"
    system_role_id = rows_for(db_path, "admin_roles")[0]["id"]
    ops_role_id = next(row["id"] for row in rows_for(db_path, "admin_roles") if row["role_name"] != main_module.SYSTEM_ADMIN_ROLE_NAME)
    with sqlite3.connect(db_path) as connection:
        connection.execute(
            "UPDATE admin_users SET role_id = ?, allow_login = 1, is_active = 1, data_scope = 'stores' WHERE username = ?",
            (system_role_id, "admin"),
        )
        connection.execute(
            """
            INSERT OR IGNORE INTO admin_users (
                username, display_name, password_hash, role_id, allow_login, participate_schedule,
                is_active, is_assignable, data_scope, store_names, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, 1, 0, 1, 1, 'all', '', ?, ?)
            """,
            ("staff", "Staff", main_module.hash_password("pw"), ops_role_id, timestamp, timestamp),
        )
        connection.execute(
            """
            INSERT INTO tickets (
                ticket_no, created_at, updated_at, store_name, submitter, request_type, urgency,
                brand, quantity, description, status
            )
            VALUES ('RESET-001', ?, ?, 'Reset Store', 'tester', 'Reset Type', 'normal', 'Reset Brand', 1, 'history', 'open')
            """,
            (timestamp, timestamp),
        )
        connection.execute("INSERT INTO ticket_stores (ticket_id, store_name, created_at) VALUES (1, 'Reset Store', ?)", (timestamp,))
        connection.execute("INSERT INTO ticket_brands (ticket_id, brand, created_at) VALUES (1, 'Reset Brand', ?)", (timestamp,))
        connection.execute("INSERT INTO ticket_images (ticket_id, image_path, uploaded_at) VALUES (1, 'uploads/old.png', ?)", (timestamp,))
        connection.execute(
            """
            INSERT INTO ticket_files (ticket_id, original_filename, stored_filename, file_path, file_ext, file_size, uploaded_at)
            VALUES (1, 'old.txt', 'old.txt', 'uploads/old.txt', 'txt', 3, ?)
            """,
            (timestamp,),
        )
        connection.execute("INSERT INTO ticket_supplements (ticket_id, store_name, submitter, note, created_at) VALUES (1, 'Reset Store', 'tester', 'note', ?)", (timestamp,))
        connection.execute("INSERT INTO ticket_participants (ticket_id, participant_type, participant_name, created_at) VALUES (1, 'user', 'staff', ?)", (timestamp,))
        connection.execute("INSERT INTO ticket_comments (ticket_id, author_type, author_name, content, visibility, created_at) VALUES (1, 'admin', 'staff', 'comment', 'public', ?)", (timestamp,))
        connection.execute("INSERT INTO ticket_tasks (ticket_id, title, status, created_at, updated_at) VALUES (1, 'task', 'open', ?, ?)", (timestamp, timestamp))
        connection.execute("INSERT INTO ticket_logs (ticket_id, action, operator, created_at) VALUES (1, 'create', 'staff', ?)", (timestamp,))
        connection.execute(
            """
            INSERT INTO employees (employee_name, store_name, participate_schedule, created_at, updated_at)
            VALUES ('Reset Employee', 'Reset Store', 1, ?, ?)
            """,
            (timestamp, timestamp),
        )
        connection.execute("INSERT INTO employee_store_map (employee_id, store_name, created_at) VALUES (1, 'Reset Store', ?)", (timestamp,))
        connection.execute("INSERT INTO shift_types (shift_name, is_global, created_at, updated_at) VALUES ('Reset Shift', 1, ?, ?)", (timestamp, timestamp))
        connection.execute("INSERT INTO store_schedules (store_name, employee_id, schedule_date, shift_type_id, created_at, updated_at) VALUES ('Reset Store', 1, '2026-07-08', 1, ?, ?)", (timestamp, timestamp))
        connection.execute("INSERT INTO schedule_logs (schedule_id, action, created_at) VALUES (1, 'create', ?)", (timestamp,))
        connection.execute("INSERT INTO store_business_hours (store_name, business_start_time, business_end_time, created_at, updated_at) VALUES ('Reset Store', '09:00', '18:00', ?, ?)", (timestamp, timestamp))
        connection.execute("INSERT INTO personnel_match_ignores (user_id, employee_id, reason, ignored_at) VALUES (2, 1, 'ignore', ?)", (timestamp,))
        connection.execute("INSERT INTO notification_events (event_type, ticket_id, ticket_no, title, severity, created_at) VALUES ('ticket', 1, 'RESET-001', 'notice', 'info', ?)", (timestamp,))
        connection.execute("INSERT INTO notification_reads (event_id, username, read_at) VALUES (1, 'staff', ?)", (timestamp,))
        connection.execute("INSERT INTO embedded_pages (page_key, title, nav_label, filename, created_at, updated_at) VALUES ('reset-page', 'Reset Page', 'Reset', 'index.html', ?, ?)", (timestamp, timestamp))
        connection.execute("INSERT INTO ticket_assignment_rules (request_type, default_handler, created_at, updated_at) VALUES ('Reset Type', 'staff', ?, ?)", (timestamp, timestamp))
        connection.execute("INSERT INTO ticket_sla_rules (request_type, urgency_level, due_hours, created_at, updated_at) VALUES ('Reset Type', 'normal', 24, ?, ?)", (timestamp, timestamp))
        connection.execute("INSERT INTO request_type_templates (request_type, template_name, created_at, updated_at) VALUES ('Reset Type', 'Reset Template', ?, ?)", (timestamp, timestamp))
        connection.execute("INSERT INTO admin_login_logs (username, success, created_at) VALUES ('staff', 1, ?)", (timestamp,))
        connection.execute("INSERT INTO admin_operation_logs (username, action, created_at) VALUES ('staff', 'reset.fixture', ?)", (timestamp,))

    upload_dir = tmp_path / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)
    (upload_dir / "old.txt").write_text("old", encoding="utf-8")
    (upload_dir / "nested").mkdir()
    (upload_dir / "nested" / "old2.txt").write_text("old2", encoding="utf-8")
    embedded_dir = tmp_path / "embedded_pages"
    (embedded_dir / "reset-page").mkdir(parents=True, exist_ok=True)
    (embedded_dir / "reset-page" / "index.html").write_text("<html>old</html>", encoding="utf-8")


def snapshot_counts(db_path: Path, table_names):
    return {table_name: count_rows(db_path, table_name) for table_name in table_names}


def test_reset_to_admin_only_dry_run_does_not_modify_database(tmp_path, monkeypatch):
    main, reset = import_main_and_reset(tmp_path, monkeypatch)
    seed_reset_history(tmp_path, main)
    db_path = tmp_path / "tickets.db"
    before = snapshot_counts(db_path, ["admin_users", "tickets", "employees", "store_schedules"])
    fake_backup_runner = fake_backup_runner_factory(tmp_path)

    report = reset.run_reset(["--dry-run", "--keep-admin", "admin"], backup_runner=fake_backup_runner)

    assert report.mode == "dry-run"
    assert fake_backup_runner.calls
    assert snapshot_counts(db_path, ["admin_users", "tickets", "employees", "store_schedules"]) == before
    assert (tmp_path / "uploads" / "old.txt").is_file()
    assert (tmp_path / "embedded_pages" / "reset-page" / "index.html").is_file()
    assert report.after_counts["tickets"] == 0
    assert report.after_counts["admin_users"] == 1
    assert report.deleted_accounts == ["ops", "staff"]


def test_reset_to_admin_only_execute_requires_confirm(tmp_path, monkeypatch):
    main, reset = import_main_and_reset(tmp_path, monkeypatch)
    seed_reset_history(tmp_path, main)
    fake_backup_runner = fake_backup_runner_factory(tmp_path)

    with pytest.raises(reset.ResetSafetyError, match="RESET_TO_ADMIN_ONLY"):
        reset.run_reset(["--execute", "--keep-admin", "admin"], backup_runner=fake_backup_runner)

    assert count_rows(tmp_path / "tickets.db", "tickets") == 1
    assert not fake_backup_runner.calls


def test_reset_to_admin_only_execute_keeps_admin_and_clears_history(tmp_path, monkeypatch):
    main, reset = import_main_and_reset(tmp_path, monkeypatch)
    seed_reset_history(tmp_path, main)
    fake_backup_runner = fake_backup_runner_factory(tmp_path)

    report = reset.run_reset(
        ["--execute", "--confirm", "RESET_TO_ADMIN_ONLY", "--keep-admin", "admin"],
        backup_runner=fake_backup_runner,
    )

    db_path = tmp_path / "tickets.db"
    users = rows_for(db_path, "admin_users")
    assert [user["username"] for user in users] == ["admin"]
    admin_user = users[0]
    assert admin_user["is_active"] == 1
    assert admin_user["allow_login"] == 1
    assert admin_user["is_assignable"] == 0
    assert admin_user["data_scope"] == "all"
    role = rows_for(db_path, "admin_roles")
    admin_role_id = next(row["id"] for row in role if row["role_name"] == main.SYSTEM_ADMIN_ROLE_NAME)
    assert admin_user["role_id"] == admin_role_id
    assert count_rows(db_path, "tickets") == 0
    assert count_rows(db_path, "employees") == 0
    assert count_rows(db_path, "store_schedules") == 0
    assert count_rows(db_path, "admin_login_logs") == 0
    assert count_rows(db_path, "admin_operation_logs") == 0
    assert db_path.is_file()
    assert sorted(row["permission_key"] for row in rows_for(db_path, "admin_role_permissions") if row["role_id"] == admin_role_id) == sorted(main.ADMIN_PERMISSION_KEYS)
    assert (tmp_path / "uploads").is_dir()
    assert list((tmp_path / "uploads").rglob("*")) == []
    assert (tmp_path / "embedded_pages").is_dir()
    assert list((tmp_path / "embedded_pages").rglob("*")) == []
    assert (report.backup_dir / "reset_file_backup" / "uploads" / "old.txt").is_file()
    assert (report.backup_dir / "reset_file_backup" / "data" / "embedded_pages" / "reset-page" / "index.html").is_file()
    assert report.upload_files_cleared == 2
    assert report.embedded_page_files_cleared == 1
    assert report.login_verification_ok is True
    assert report.sqlite_sequence_reset is True
    assert report.residual_non_admin_accounts == 0
    assert report.residual_tickets == 0
    assert report.residual_employees == 0
    assert report.residual_schedules == 0
    assert report.residual_business_config == 0
    report_text = report.report_path.read_text(encoding="utf-8")
    assert "是否重置 sqlite_sequence" in report_text
    assert "清理后是否仍有非管理员账号：否" in report_text
    assert "清理后是否仍有工单数据：否" in report_text
    assert "清理后是否仍有人员数据：否" in report_text
    assert "清理后是否仍有排班数据：否" in report_text
    assert "清理后是否仍有业务配置数据：否" in report_text

    client = TestClient(main.app)
    response = client.post("/admin/login", data={"username": "admin", "password": "123456"}, follow_redirects=False)
    assert response.status_code == 303
    assert response.headers["location"].startswith("/admin")
    overview = client.get("/admin/permission-overview")
    assert overview.status_code == 200
    assert "高风险 POST 未接入：0" in overview.text


def test_reset_to_admin_only_execute_rolls_back_database_on_failure(tmp_path, monkeypatch):
    main, reset = import_main_and_reset(tmp_path, monkeypatch)
    seed_reset_history(tmp_path, main)
    db_path = tmp_path / "tickets.db"
    with sqlite3.connect(db_path) as connection:
        connection.execute(
            """
            CREATE TRIGGER fail_reset_employee_delete
            BEFORE DELETE ON employees
            BEGIN
                SELECT RAISE(ABORT, 'forced rollback');
            END
            """
        )
    fake_backup_runner = fake_backup_runner_factory(tmp_path)

    with pytest.raises(sqlite3.DatabaseError, match="forced rollback"):
        reset.run_reset(
            ["--execute", "--confirm", "RESET_TO_ADMIN_ONLY", "--keep-admin", "admin"],
            backup_runner=fake_backup_runner,
        )

    assert count_rows(db_path, "tickets") == 1
    assert count_rows(db_path, "employees") == 1
    assert sorted(user["username"] for user in rows_for(db_path, "admin_users")) == ["admin", "ops", "staff"]
    assert (tmp_path / "uploads" / "old.txt").is_file()
    assert (tmp_path / "embedded_pages" / "reset-page" / "index.html").is_file()
