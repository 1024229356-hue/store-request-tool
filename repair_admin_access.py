from __future__ import annotations

import sys

import main


def configure_output() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def run() -> int:
    configure_output()
    credentials = main.get_env_admin_credentials()
    if not credentials:
        print("未找到 .env ADMIN_USERS，请先配置至少一个后台账号。")
        return 1

    username = credentials[0][0].strip()
    result = main.ensure_admin_access_safeguard()
    if result.get("active_system_admin_count", 0) <= 0:
        print("修复失败：仍然没有 active 系统管理员账号。")
        return 1

    print(f"已检查系统管理员角色权限：补齐 {result.get('permissions_added', 0)} 个权限点。")
    if result.get("account_created"):
        print(f"已创建 .env 第一个 ADMIN_USERS 账号为系统管理员：{username}")
    elif result.get("account_recovered"):
        print(f"已恢复 .env 第一个 ADMIN_USERS 账号为系统管理员：{username}")
    else:
        print("active 系统管理员账号已存在，仅确认权限完整。")
    print("请重启 run.bat，并使用第一个 ADMIN_USERS 账号登录。")
    return 0


if __name__ == "__main__":
    raise SystemExit(run())
