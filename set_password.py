"""
设置/更新启动密码（供 scan_orders_main.py 与 mhg_batch_query.py 共用）
使用：python set_password.py
生成文件：auth_config.json（放在当前目录）
"""

import os
import json
import hashlib
import getpass

AUTH_CONFIG_FILE = "auth_config.json"
AUTH_PEPPER = "MHG@2026#scan-query#pepper"


def sha256_hex(text: str) -> str:
    return hashlib.sha256(str(text or "").encode("utf-8")).hexdigest()


def hash_password(password: str) -> str:
    mixed = f"{AUTH_PEPPER}{password}{AUTH_PEPPER}"
    return sha256_hex(mixed)


def main():
    print("=" * 50)
    print("  设置启动密码")
    print("=" * 50)

    try:
        pwd1 = getpass.getpass("请输入新密码: ")
        pwd2 = getpass.getpass("请再次输入新密码: ")
    except Exception:
        pwd1 = input("请输入新密码: ").strip()
        pwd2 = input("请再次输入新密码: ").strip()

    if not pwd1:
        print("❌ 密码不能为空")
        return

    if pwd1 != pwd2:
        print("❌ 两次密码不一致")
        return

    if len(pwd1) < 4:
        print("❌ 密码长度至少 4 位")
        return

    data = {
        "password_sha256": hash_password(pwd1),
        "hint": "Do not store plaintext password.",
    }

    out_path = os.path.join(os.getcwd(), AUTH_CONFIG_FILE)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print("\n✅ 密码已设置成功")
    print(f"   配置文件: {out_path}")
    print("   请将该 auth_config.json 与 exe/py 放在同一目录")


if __name__ == "__main__":
    main()
