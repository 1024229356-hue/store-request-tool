from __future__ import annotations

import argparse
import ast
import hashlib
import hmac
import os
import secrets
import shutil
import sqlite3
import subprocess
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple


PROJECT_ROOT = Path(__file__).resolve().parents[1]
MAIN_PATH = PROJECT_ROOT / "main.py"
ENV_PATH = PROJECT_ROOT / ".env"
DEFAULT_DB_PATH = PROJECT_ROOT / "data" / "tickets.db"
DEFAULT_UPLOAD_DIR = PROJECT_ROOT / "uploads"
REPORT_DIR = PROJECT_ROOT / "docs"
CONFIRM_TOKEN = "RESET_TO_ADMIN_ONLY"

PASSWORD_HASH_ITERATIONS = 260000

TICKET_TABLES = [
    "ticket_stores",
    "ticket_brands",
    "ticket_images",
    "ticket_files",
    "ticket_supplements",
    "ticket_participants",
    "ticket_comments",
    "ticket_tasks",
    "ticket_logs",
    "tickets",
]

SCHEDULE_TABLES = [
    "schedule_logs",
    "store_schedules",
    "employee_store_map",
    "employees",
    "shift_types",
    "store_business_hours",
]

GOVERNANCE_TABLES = ["personnel_match_ignores"]
NOTIFICATION_TABLES = ["notification_reads", "notification_events"]
EMBEDDED_TABLES = ["embedded_pages"]
CONFIG_TABLES = [
    "ticket_assignment_rules",
    "ticket_sla_rules",
    "request_type_templates",
    "stores",
    "brands",
    "holidays",
    "request_types",
]
AUDIT_TABLES = ["admin_login_logs", "admin_operation_logs"]

FULL_CLEAR_TABLES = [
    *TICKET_TABLES,
    *SCHEDULE_TABLES,
    *GOVERNANCE_TABLES,
    *NOTIFICATION_TABLES,
    *EMBEDDED_TABLES,
    *CONFIG_TABLES,
    *AUDIT_TABLES,
]

REPORT_TABLES = [*FULL_CLEAR_TABLES, "admin_users", "admin_roles", "admin_role_permissions"]


class ResetSafetyError(RuntimeError):
    """Raised when a safety precondition blocks the reset."""


@dataclass
class MainConstants:
    admin_permission_keys: List[str]
    system_admin_role_name: str


@dataclass
class DirectoryResetPlan:
    source_dir: Path
    backup_dir: Path
    quarantine_dir: Path
    file_count: int
    moved_children: List[Path] = field(default_factory=list)


@dataclass
class ResetReport:
    mode: str
    keep_admin: str
    backup_dir: Path
    report_path: Path
    before_counts: Dict[str, int]
    after_counts: Dict[str, int]
    deleted_accounts: List[str]
    kept_accounts: List[str]
    clear_tables: List[str]
    upload_files_cleared: int
    embedded_page_files_cleared: int
    permissions_completed: bool
    permissions_added: int
    login_verification_ok: bool
    sqlite_sequence_reset: bool
    residual_non_admin_accounts: int
    residual_tickets: int
    residual_employees: int
    residual_schedules: int
    residual_business_config: int
    config_reinitializes_from_files: bool
    db_path: Path
    upload_dir: Path
    embedded_pages_dir: Path
    warnings: List[str] = field(default_factory=list)


BackupRunner = Callable[[Path], Path]


def configure_output() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def table_exists(connection: sqlite3.Connection, table_name: str) -> bool:
    row = connection.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
        (table_name,),
    ).fetchone()
    return row is not None


def column_names(connection: sqlite3.Connection, table_name: str) -> List[str]:
    if not table_exists(connection, table_name):
        return []
    return [str(row[1]) for row in connection.execute(f"PRAGMA table_info({table_name})").fetchall()]


def row_count(connection: sqlite3.Connection, table_name: str) -> int:
    if not table_exists(connection, table_name):
        return 0
    row = connection.execute(f"SELECT COUNT(*) AS count FROM {table_name}").fetchone()
    return int(row["count"] if isinstance(row, sqlite3.Row) else row[0])


def collect_counts(connection: sqlite3.Connection, table_names: Iterable[str]) -> Dict[str, int]:
    return {table_name: row_count(connection, table_name) for table_name in table_names}


def load_main_constants(main_path: Path = MAIN_PATH) -> MainConstants:
    tree = ast.parse(main_path.read_text(encoding="utf-8"))
    values: Dict[str, object] = {}
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id in {"ADMIN_PERMISSION_KEYS", "SYSTEM_ADMIN_ROLE_NAME"}:
                    values[target.id] = ast.literal_eval(node.value)
    permission_keys = values.get("ADMIN_PERMISSION_KEYS")
    role_name = values.get("SYSTEM_ADMIN_ROLE_NAME")
    if not isinstance(permission_keys, list) or not all(isinstance(item, str) for item in permission_keys):
        raise ResetSafetyError("无法从 main.py 读取 ADMIN_PERMISSION_KEYS，禁止继续。")
    if not isinstance(role_name, str) or not role_name.strip():
        raise ResetSafetyError("无法从 main.py 读取 SYSTEM_ADMIN_ROLE_NAME，禁止继续。")
    return MainConstants(admin_permission_keys=list(permission_keys), system_admin_role_name=role_name)


def load_env_values(env_path: Path = ENV_PATH) -> Dict[str, str]:
    values = dict(os.environ)
    if not env_path.exists():
        return values
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        clean_key = key.strip()
        if clean_key and clean_key not in values:
            values[clean_key] = value.strip().strip("\"'")
    return values


def parse_admin_users(raw_users: str) -> List[Tuple[str, str]]:
    users: List[Tuple[str, str]] = []
    for raw_item in raw_users.split(","):
        username, separator, password = raw_item.partition(":")
        if not separator:
            continue
        clean_username = username.strip()
        clean_password = password.strip()
        if clean_username and clean_password:
            users.append((clean_username, clean_password))
    return users


def env_admin_credentials(env_values: Dict[str, str]) -> List[Tuple[str, str]]:
    admin_users = env_values.get("ADMIN_USERS")
    if admin_users is not None:
        users = parse_admin_users(admin_users)
        if users:
            return users
    username = env_values.get("ADMIN_USERNAME", "").strip()
    password = env_values.get("ADMIN_PASSWORD", "")
    if username and password:
        return [(username, password)]
    return []


def hash_password(password: str, iterations: int = PASSWORD_HASH_ITERATIONS) -> str:
    salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("ascii"), iterations)
    return f"pbkdf2_sha256${iterations}${salt}${digest.hex()}"


def verify_password(password: str, password_hash: str) -> bool:
    parts = (password_hash or "").split("$")
    if len(parts) != 4 or parts[0] != "pbkdf2_sha256":
        return False
    try:
        iterations = int(parts[1])
        salt = parts[2]
        expected_hash = parts[3]
        digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("ascii"), iterations)
    except (TypeError, ValueError):
        return False
    return hmac.compare_digest(digest.hex(), expected_hash)


def resolve_db_path(env_values: Dict[str, str]) -> Path:
    return Path(env_values.get("STORE_REQUEST_DB_PATH") or DEFAULT_DB_PATH)


def resolve_upload_dir(env_values: Dict[str, str]) -> Path:
    return Path(env_values.get("STORE_REQUEST_UPLOAD_DIR") or DEFAULT_UPLOAD_DIR)


def resolve_embedded_pages_dir(db_path: Path) -> Path:
    return db_path.parent / "embedded_pages"


def resolve_report_dir(env_values: Dict[str, str]) -> Path:
    return Path(env_values.get("STORE_REQUEST_RESET_REPORT_DIR") or REPORT_DIR)


def default_keep_admin(credentials: List[Tuple[str, str]]) -> str:
    if not credentials:
        raise ResetSafetyError("未指定 --keep-admin，且 .env/环境变量中没有可用 ADMIN_USERS。")
    return credentials[0][0].strip()


def newest_backup_dir(backups_root: Path) -> Optional[Path]:
    if not backups_root.exists():
        return None
    dirs = [path for path in backups_root.iterdir() if path.is_dir()]
    if not dirs:
        return None
    return max(dirs, key=lambda path: path.stat().st_mtime)


def parse_backup_dir_from_output(output: str) -> Optional[Path]:
    lines = [line.strip() for line in output.splitlines() if line.strip()]
    for line in reversed(lines):
        path = Path(line)
        if path.exists() and path.is_dir():
            return path
    return None


def run_backup_bat(project_root: Path = PROJECT_ROOT) -> Path:
    backup_script = project_root / "backup.bat"
    if not backup_script.exists():
        raise ResetSafetyError("未找到 backup.bat，禁止继续。")
    completed = subprocess.run(
        str(backup_script),
        cwd=project_root,
        shell=True,
        capture_output=True,
        text=True,
        errors="replace",
    )
    output = "\n".join(part for part in (completed.stdout, completed.stderr) if part)
    if completed.returncode != 0:
        raise ResetSafetyError(f"backup.bat 执行失败，禁止继续。\n{output.strip()}")
    parsed = parse_backup_dir_from_output(output)
    if parsed:
        return parsed
    newest = newest_backup_dir(project_root / "backups")
    if newest:
        return newest
    raise ResetSafetyError("backup.bat 已执行，但无法确认备份目录，禁止继续。")


def count_files(root: Path) -> int:
    if not root.exists():
        return 0
    return sum(1 for path in root.rglob("*") if path.is_file())


def copy_directory_snapshot(source_dir: Path, target_dir: Path) -> None:
    target_dir.parent.mkdir(parents=True, exist_ok=True)
    if target_dir.exists():
        shutil.rmtree(target_dir)
    if source_dir.exists():
        shutil.copytree(source_dir, target_dir)
    else:
        target_dir.mkdir(parents=True, exist_ok=True)


def prepare_directory_reset(source_dir: Path, backup_target: Path, quarantine_target: Path) -> DirectoryResetPlan:
    source_dir.mkdir(parents=True, exist_ok=True)
    copy_directory_snapshot(source_dir, backup_target)
    quarantine_target.parent.mkdir(parents=True, exist_ok=True)
    if quarantine_target.exists():
        shutil.rmtree(quarantine_target)
    quarantine_target.mkdir(parents=True, exist_ok=True)
    plan = DirectoryResetPlan(
        source_dir=source_dir,
        backup_dir=backup_target,
        quarantine_dir=quarantine_target,
        file_count=count_files(source_dir),
    )
    try:
        for child in list(source_dir.iterdir()):
            destination = quarantine_target / child.name
            shutil.move(str(child), str(destination))
            plan.moved_children.append(destination)
    except Exception:
        restore_directory_reset(plan)
        raise
    source_dir.mkdir(parents=True, exist_ok=True)
    return plan


def restore_directory_reset(plan: DirectoryResetPlan) -> None:
    plan.source_dir.mkdir(parents=True, exist_ok=True)
    for moved_child in reversed(plan.moved_children):
        if not moved_child.exists():
            continue
        destination = plan.source_dir / moved_child.name
        if destination.exists():
            if destination.is_dir():
                shutil.rmtree(destination)
            else:
                destination.unlink()
        shutil.move(str(moved_child), str(destination))


def role_id_for_name(connection: sqlite3.Connection, role_name: str) -> Optional[int]:
    if not table_exists(connection, "admin_roles"):
        return None
    row = connection.execute("SELECT id FROM admin_roles WHERE role_name = ?", (role_name,)).fetchone()
    return int(row["id"]) if row else None


def ensure_system_admin_role(connection: sqlite3.Connection, constants: MainConstants) -> Tuple[int, int, bool]:
    if not table_exists(connection, "admin_roles") or not table_exists(connection, "admin_role_permissions"):
        raise ResetSafetyError("admin_roles/admin_role_permissions 表不存在，禁止继续。")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    role_id = role_id_for_name(connection, constants.system_admin_role_name)
    if role_id is None:
        cursor = connection.execute(
            """
            INSERT INTO admin_roles (role_name, description, is_system, created_at, updated_at)
            VALUES (?, ?, 1, ?, ?)
            """,
            (constants.system_admin_role_name, "系统管理员角色由 reset_to_admin_only.py 补齐。", timestamp, timestamp),
        )
        role_id = int(cursor.lastrowid)
    else:
        connection.execute(
            "UPDATE admin_roles SET is_system = 1, updated_at = ? WHERE id = ?",
            (timestamp, role_id),
        )
    before = row_count(connection, "admin_role_permissions")
    for permission_key in constants.admin_permission_keys:
        connection.execute(
            """
            INSERT OR IGNORE INTO admin_role_permissions (role_id, permission_key, created_at)
            VALUES (?, ?, ?)
            """,
            (role_id, permission_key, timestamp),
        )
    permissions = [
        str(row["permission_key"])
        for row in connection.execute(
            "SELECT permission_key FROM admin_role_permissions WHERE role_id = ?",
            (role_id,),
        ).fetchall()
    ]
    permission_set = set(permissions)
    complete = set(constants.admin_permission_keys).issubset(permission_set)
    return role_id, row_count(connection, "admin_role_permissions") - before, complete


def credential_password_for(credentials: List[Tuple[str, str]], username: str) -> Optional[str]:
    for candidate_username, password in credentials:
        if candidate_username == username:
            return password
    return None


def ensure_keep_admin_account(
    connection: sqlite3.Connection,
    keep_admin: str,
    system_role_id: int,
    credentials: List[Tuple[str, str]],
) -> None:
    if not table_exists(connection, "admin_users"):
        raise ResetSafetyError("admin_users 表不存在，禁止继续。")
    columns = set(column_names(connection, "admin_users"))
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    existing = connection.execute("SELECT * FROM admin_users WHERE username = ?", (keep_admin,)).fetchone()
    assignments = []
    params: List[object] = []
    desired_values: Dict[str, object] = {
        "role_id": system_role_id,
        "employee_id": None,
        "allow_login": 1,
        "participate_schedule": 0,
        "is_active": 1,
        "is_assignable": 0,
        "data_scope": "all",
        "store_names": "",
        "updated_at": timestamp,
        "updated_by": "reset_to_admin_only",
    }
    if "show_in_account_management" in columns:
        desired_values["show_in_account_management"] = 1
    for column_name, value in desired_values.items():
        if column_name in columns:
            assignments.append(f"{column_name} = ?")
            params.append(value)
    if existing:
        connection.execute(
            f"UPDATE admin_users SET {', '.join(assignments)} WHERE username = ?",
            (*params, keep_admin),
        )
        return

    password = credential_password_for(credentials, keep_admin)
    if not password:
        raise ResetSafetyError(f"保留管理员账号 {keep_admin} 不存在，且 .env/环境变量中没有该账号密码，禁止继续。")
    insert_values: Dict[str, object] = {
        "username": keep_admin,
        "display_name": keep_admin,
        "password_hash": hash_password(password),
        "role_id": system_role_id,
        "employee_id": None,
        "allow_login": 1,
        "participate_schedule": 0,
        "is_active": 1,
        "is_assignable": 0,
        "show_in_account_management": 1,
        "data_scope": "all",
        "store_names": "",
        "created_at": timestamp,
        "updated_at": timestamp,
        "created_by": "reset_to_admin_only",
        "updated_by": "reset_to_admin_only",
    }
    clean_columns = [column_name for column_name in insert_values if column_name in columns]
    placeholders = ", ".join("?" for _ in clean_columns)
    connection.execute(
        f"INSERT INTO admin_users ({', '.join(clean_columns)}) VALUES ({placeholders})",
        tuple(insert_values[column_name] for column_name in clean_columns),
    )


def admin_accounts(connection: sqlite3.Connection) -> List[str]:
    if not table_exists(connection, "admin_users"):
        return []
    rows = connection.execute("SELECT username FROM admin_users ORDER BY username").fetchall()
    return [str(row["username"]) for row in rows]


def active_system_admin_count(connection: sqlite3.Connection, system_role_name: str) -> int:
    if not table_exists(connection, "admin_users") or not table_exists(connection, "admin_roles"):
        return 0
    row = connection.execute(
        """
        SELECT COUNT(*) AS count
        FROM admin_users
        JOIN admin_roles ON admin_roles.id = admin_users.role_id
        WHERE admin_users.is_active = 1
          AND COALESCE(admin_users.allow_login, 1) = 1
          AND admin_roles.role_name = ?
        """,
        (system_role_name,),
    ).fetchone()
    return int(row["count"] or 0)


def verify_admin_login(connection: sqlite3.Connection, keep_admin: str, credentials: List[Tuple[str, str]]) -> bool:
    password = credential_password_for(credentials, keep_admin)
    if not password or not table_exists(connection, "admin_users"):
        return False
    row = connection.execute(
        """
        SELECT admin_users.*, admin_roles.role_name
        FROM admin_users
        LEFT JOIN admin_roles ON admin_roles.id = admin_users.role_id
        WHERE admin_users.username = ?
        """,
        (keep_admin,),
    ).fetchone()
    if not row:
        return False
    if int(row["is_active"] or 0) != 1 or int(row["allow_login"] or 0) != 1:
        return False
    return verify_password(password, str(row["password_hash"] or ""))


def clear_sqlite_sequence(connection: sqlite3.Connection, cleared_tables: List[str], partial_tables: List[str]) -> None:
    if not table_exists(connection, "sqlite_sequence"):
        return
    for table_name in cleared_tables:
        connection.execute("DELETE FROM sqlite_sequence WHERE name = ?", (table_name,))
    for table_name in partial_tables:
        if not table_exists(connection, table_name):
            continue
        max_row = connection.execute(f"SELECT COALESCE(MAX(id), 0) AS max_id FROM {table_name}").fetchone()
        max_id = int(max_row["max_id"] if isinstance(max_row, sqlite3.Row) else max_row[0])
        if max_id <= 0:
            connection.execute("DELETE FROM sqlite_sequence WHERE name = ?", (table_name,))
        else:
            connection.execute("UPDATE sqlite_sequence SET seq = ? WHERE name = ?", (max_id, table_name))


def projected_after_counts(before_counts: Dict[str, int], kept_account_count: int) -> Dict[str, int]:
    after = dict(before_counts)
    for table_name in FULL_CLEAR_TABLES:
        after[table_name] = 0
    after["admin_users"] = kept_account_count
    return after


def residual_group_count(counts: Dict[str, int], table_names: Iterable[str]) -> int:
    return sum(counts.get(table_name, 0) for table_name in table_names)


def residual_non_admin_count(counts: Dict[str, int], kept_accounts: Sequence[str]) -> int:
    kept_count = 1 if kept_accounts else 0
    return max(counts.get("admin_users", 0) - kept_count, 0)


def write_report(report: ResetReport) -> None:
    report.report_path.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        f"# Reset To Admin Only Report",
        "",
        f"- 执行时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"- 模式：{report.mode}",
        f"- 保留管理员账号：{report.keep_admin}",
        f"- 备份目录：{report.backup_dir}",
        f"- 数据库文件：{report.db_path}",
        f"- 上传目录：{report.upload_dir}",
        f"- 嵌入页面目录：{report.embedded_pages_dir}",
        f"- 删除的账号数量：{len(report.deleted_accounts)}",
        f"- 将删除/已删除账号：{', '.join(report.deleted_accounts) if report.deleted_accounts else '无'}",
        f"- 保留账号：{', '.join(report.kept_accounts) if report.kept_accounts else report.keep_admin}",
        f"- 清空的上传文件数量：{report.upload_files_cleared}",
        f"- 清空的嵌入页面文件数量：{report.embedded_page_files_cleared}",
        f"- 是否补齐系统管理员权限：{'是' if report.permissions_completed else '否'}（新增 {report.permissions_added} 个权限点）",
        f"- 是否可登录验证：{'是' if report.login_verification_ok else '否'}",
        f"- 是否重置 sqlite_sequence：{'是' if report.sqlite_sequence_reset else '否'}（dry-run 为预计；execute 时仅重置已清空表和 admin_users 序列）",
        f"- 清理后是否仍有非管理员账号：{'是' if report.residual_non_admin_accounts else '否'}（{report.residual_non_admin_accounts} 个）",
        f"- 清理后是否仍有工单数据：{'是' if report.residual_tickets else '否'}（{report.residual_tickets} 行）",
        f"- 清理后是否仍有人员数据：{'是' if report.residual_employees else '否'}（{report.residual_employees} 行）",
        f"- 清理后是否仍有排班数据：{'是' if report.residual_schedules else '否'}（{report.residual_schedules} 行）",
        f"- 清理后是否仍有业务配置数据：{'是' if report.residual_business_config else '否'}（{report.residual_business_config} 行）",
        f"- 配置表清空后是否可由 config/*.json 自动初始化：{'是' if report.config_reinitializes_from_files else '否'}",
        "",
        "## 清空表",
        "",
        ", ".join(report.clear_tables),
        "",
        "## 表行数",
        "",
        "| 表 | 清理前 | 清理后/预计 |",
        "|---|---:|---:|",
    ]
    for table_name in REPORT_TABLES:
        lines.append(f"| {table_name} | {report.before_counts.get(table_name, 0)} | {report.after_counts.get(table_name, 0)} |")
    lines.extend(
        [
            "",
            "## 风险提示",
            "",
            "- 本脚本不会删除 data/tickets.db 文件本身，不会删除 .env、config、static、templates、tests 或代码文件。",
            "- dry-run 模式只生成预览和备份，不执行数据库删除，也不清空 uploads 或 data/embedded_pages。",
            "- execute 模式必须显式传入 --confirm RESET_TO_ADMIN_ONLY；执行前已先运行 backup.bat。",
            "- stores、brands、holidays、request_types 清空后，应用启动时可按 config/*.json 的既有逻辑重新初始化或回退显示配置项。",
        ]
    )
    if report.warnings:
        lines.extend(["", "## 警告", ""])
        lines.extend(f"- {warning}" for warning in report.warnings)
    lines.extend(
        [
            "",
            "## 下一步建议",
            "",
            "- 先人工审阅本报告和备份目录。",
            "- 需要真实清理时，使用 --execute --confirm RESET_TO_ADMIN_ONLY --keep-admin <账号>。",
            "- 真实清理后再运行 /healthz、/__version、/admin/login、/admin/dashboard、/admin/account、/admin/employees、/admin/schedules、/admin/settings、/admin/permission-overview 和 pytest 验证。",
            "",
        ]
    )
    report.report_path.write_text("\n".join(lines), encoding="utf-8")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Reset store_request_tool history data and keep only one system admin.")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--dry-run", action="store_true", help="Preview reset actions without deleting data. This is the default.")
    mode.add_argument("--execute", action="store_true", help="Execute destructive reset after backup and confirmation.")
    parser.add_argument("--confirm", default="", help=f"Required for --execute: {CONFIRM_TOKEN}")
    parser.add_argument("--keep-admin", default="", help="Admin username to keep. Defaults to first ADMIN_USERS account in .env.")
    return parser


def run_reset(
    argv: Optional[Sequence[str]] = None,
    backup_runner: BackupRunner = run_backup_bat,
) -> ResetReport:
    args = build_parser().parse_args(list(argv) if argv is not None else None)
    mode = "execute" if args.execute else "dry-run"
    if mode == "execute" and args.confirm != CONFIRM_TOKEN:
        raise ResetSafetyError(f"真正执行必须传入 --confirm {CONFIRM_TOKEN}。")

    constants = load_main_constants()
    env_values = load_env_values()
    credentials = env_admin_credentials(env_values)
    keep_admin = (args.keep_admin or default_keep_admin(credentials)).strip()
    if not keep_admin:
        raise ResetSafetyError("保留管理员账号为空，禁止继续。")

    db_path = resolve_db_path(env_values)
    upload_dir = resolve_upload_dir(env_values)
    embedded_pages_dir = resolve_embedded_pages_dir(db_path)
    report_dir = resolve_report_dir(env_values)
    if not db_path.exists():
        raise ResetSafetyError(f"数据库文件不存在：{db_path}")

    backup_dir = backup_runner(PROJECT_ROOT)
    if not backup_dir.exists() or not backup_dir.is_dir():
        raise ResetSafetyError(f"备份目录不存在，禁止继续：{backup_dir}")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    report_path = report_dir / f"reset_to_admin_only_report_{timestamp}.md"
    clear_tables = [table_name for table_name in FULL_CLEAR_TABLES]

    upload_file_count = count_files(upload_dir)
    embedded_file_count = count_files(embedded_pages_dir)
    permissions_added = 0
    permissions_completed = False
    login_verification_ok = False
    sqlite_sequence_reset = False
    directory_plans: List[DirectoryResetPlan] = []

    with sqlite3.connect(db_path) as connection:
        connection.row_factory = sqlite3.Row
        before_counts = collect_counts(connection, REPORT_TABLES)
        existing_accounts = admin_accounts(connection)
        deleted_accounts = [username for username in existing_accounts if username != keep_admin]
        if keep_admin not in existing_accounts and credential_password_for(credentials, keep_admin):
            kept_accounts = [keep_admin]
        else:
            kept_accounts = [username for username in existing_accounts if username == keep_admin]

        if mode == "dry-run":
            system_role_id = role_id_for_name(connection, constants.system_admin_role_name)
            if system_role_id is not None and table_exists(connection, "admin_role_permissions"):
                permission_rows = connection.execute(
                    "SELECT permission_key FROM admin_role_permissions WHERE role_id = ?",
                    (system_role_id,),
                ).fetchall()
                permissions_completed = set(constants.admin_permission_keys).issubset({str(row["permission_key"]) for row in permission_rows})
            after_counts = projected_after_counts(before_counts, len(kept_accounts))
            sqlite_sequence_reset = True
        else:
            try:
                directory_plans = [
                    prepare_directory_reset(
                        upload_dir,
                        backup_dir / "reset_file_backup" / "uploads",
                        backup_dir / "reset_file_quarantine" / "uploads",
                    ),
                    prepare_directory_reset(
                        embedded_pages_dir,
                        backup_dir / "reset_file_backup" / "data" / "embedded_pages",
                        backup_dir / "reset_file_quarantine" / "data" / "embedded_pages",
                    ),
                ]
                connection.execute("BEGIN")
                system_role_id, permissions_added, permissions_completed = ensure_system_admin_role(connection, constants)
                ensure_keep_admin_account(connection, keep_admin, system_role_id, credentials)
                for table_name in FULL_CLEAR_TABLES:
                    if table_exists(connection, table_name):
                        connection.execute(f"DELETE FROM {table_name}")
                if table_exists(connection, "admin_users"):
                    connection.execute("DELETE FROM admin_users WHERE username != ?", (keep_admin,))
                system_role_id, more_permissions_added, permissions_completed = ensure_system_admin_role(connection, constants)
                permissions_added += more_permissions_added
                ensure_keep_admin_account(connection, keep_admin, system_role_id, credentials)
                if active_system_admin_count(connection, constants.system_admin_role_name) <= 0:
                    raise ResetSafetyError("清理后没有 active 系统管理员账号，已回滚。")
                clear_sqlite_sequence(
                    connection,
                    [table_name for table_name in FULL_CLEAR_TABLES if table_exists(connection, table_name)],
                    ["admin_users"],
                )
                sqlite_sequence_reset = True
                after_counts = collect_counts(connection, REPORT_TABLES)
                login_verification_ok = verify_admin_login(connection, keep_admin, credentials)
                connection.commit()
            except Exception:
                connection.rollback()
                for plan in reversed(directory_plans):
                    restore_directory_reset(plan)
                raise
            kept_accounts = admin_accounts(connection)

    if mode == "dry-run":
        login_verification_ok = credential_password_for(credentials, keep_admin) is not None
    else:
        upload_file_count = sum(plan.file_count for plan in directory_plans if plan.source_dir == upload_dir)
        embedded_file_count = sum(plan.file_count for plan in directory_plans if plan.source_dir == embedded_pages_dir)

    residual_non_admin_accounts = residual_non_admin_count(after_counts, kept_accounts)
    residual_tickets = residual_group_count(after_counts, TICKET_TABLES)
    residual_employees = residual_group_count(after_counts, ["employees", "employee_store_map", "personnel_match_ignores"])
    residual_schedules = residual_group_count(
        after_counts,
        ["shift_types", "store_schedules", "schedule_logs", "store_business_hours"],
    )
    residual_business_config = residual_group_count(after_counts, CONFIG_TABLES)

    report = ResetReport(
        mode=mode,
        keep_admin=keep_admin,
        backup_dir=backup_dir,
        report_path=report_path,
        before_counts=before_counts,
        after_counts=after_counts,
        deleted_accounts=deleted_accounts,
        kept_accounts=kept_accounts,
        clear_tables=clear_tables,
        upload_files_cleared=upload_file_count,
        embedded_page_files_cleared=embedded_file_count,
        permissions_completed=permissions_completed,
        permissions_added=permissions_added,
        login_verification_ok=login_verification_ok,
        sqlite_sequence_reset=sqlite_sequence_reset,
        residual_non_admin_accounts=residual_non_admin_accounts,
        residual_tickets=residual_tickets,
        residual_employees=residual_employees,
        residual_schedules=residual_schedules,
        residual_business_config=residual_business_config,
        config_reinitializes_from_files=True,
        db_path=db_path,
        upload_dir=upload_dir,
        embedded_pages_dir=embedded_pages_dir,
    )
    write_report(report)
    return report


def main(argv: Optional[Sequence[str]] = None, backup_runner: BackupRunner = run_backup_bat) -> int:
    configure_output()
    try:
        report = run_reset(argv, backup_runner=backup_runner)
    except ResetSafetyError as exc:
        print(f"[拒绝执行] {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"[执行失败] {exc}", file=sys.stderr)
        return 1
    print(f"模式：{report.mode}")
    print(f"保留管理员账号：{report.keep_admin}")
    print(f"备份目录：{report.backup_dir}")
    print(f"报告路径：{report.report_path}")
    print(f"将删除/已删除账号：{', '.join(report.deleted_accounts) if report.deleted_accounts else '无'}")
    print(f"上传文件清空数量：{report.upload_files_cleared}")
    print(f"嵌入页面文件清空数量：{report.embedded_page_files_cleared}")
    if report.mode == "dry-run":
        print("dry-run 已完成：未执行数据库删除，未清空 uploads 或 data/embedded_pages。")
    else:
        print("execute 已完成：历史数据已清理，仅保留指定系统管理员账号。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
