"""
打包说明：在有Python的电脑上运行此脚本，生成 scan_orders.exe
步骤：
  1. pip install pyinstaller openpyxl
  2. python build_exe.py
  3. 在 dist/ 文件夹找到 scan_orders*.exe，拷到公司电脑使用

注意：
- 本脚本直接打包当前目录下的 scan_orders_main.py
- 若旧的 dist/scan_orders.exe 正在运行或被占用，会自动回退为带时间戳的新文件名
"""

import os
import sys
import subprocess
from datetime import datetime


def run_pyinstaller(target_script: str, exe_name: str) -> int:
    return subprocess.run([
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--console",
        "--noconfirm",
        "--name", exe_name,
        target_script,
    ], check=False).returncode


def main():
    target_script = "scan_orders_main.py"

    if not os.path.isfile(target_script):
        print(f"❌ 未找到 {target_script}，请确认文件在当前目录")
        return

    print(f"准备打包: {target_script}")

    # 安装依赖
    subprocess.run([
        sys.executable, "-m", "pip", "install", "pyinstaller", "openpyxl",
        "-i", "https://mirrors.aliyun.com/pypi/simple/", "-q"
    ], check=False)

    # 先尝试固定名称，便于你习惯使用
    exe_name = "scan_orders"
    rc = run_pyinstaller(target_script, exe_name)

    if rc != 0:
        # 若失败（常见是旧exe被占用），自动回退到时间戳名称
        ts_name = f"scan_orders_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        print()
        print("⚠ 固定文件名打包失败，可能是旧版 scan_orders.exe 正在运行或被占用。")
        print(f"  自动改用新文件名重试: {ts_name}.exe")
        rc = run_pyinstaller(target_script, ts_name)
        exe_name = ts_name if rc == 0 else exe_name

    if rc == 0:
        print()
        print("=" * 50)
        print("✅ 打包成功！")
        print(f"   exe文件位置: dist/{exe_name}.exe")
        print("   拷贝到公司电脑即可双击运行")
        print("=" * 50)
    else:
        print("❌ 打包失败，请先关闭正在运行的 scan_orders.exe 后重试")


if __name__ == "__main__":
    main()
