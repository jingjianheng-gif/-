#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
TIA Portal Openness 环境检查脚本
=================================
检查运行 TIA Portal 变量导出脚本所需的全部前置条件。

用法:
    python check_environment.py
    python check_environment.py --install

作者: CodeWhale
日期: 2026-06-30
"""

import sys
import subprocess
import os
from pathlib import Path

# ────────────────────────────────────────────────────────────
#  配置
# ────────────────────────────────────────────────────────────

# TIA Portal 可能的安装路径
TIA_POSSIBLE_PATHS = [
    (r"C:\Program Files\Siemens\Automation\Portal V19\PublicAPI", "V19"),
    (r"C:\Program Files\Siemens\Automation\Portal V17\PublicAPI", "V17"),
    (r"C:\Program Files\Siemens\Automation\Portal V16\PublicAPI", "V16"),
    (r"C:\Program Files\Siemens\Automation\Portal V18\PublicAPI", "V18"),
]

# TIA Portal 可能的可执行文件
TIA_EXE_PATHS = [
    r"C:\Program Files\Siemens\Automation\Portal V19\bin\TiaPortal.exe",
    r"C:\Program Files\Siemens\Automation\Portal V17\bin\TiaPortal.exe",
    r"C:\Program Files\Siemens\Automation\Portal V16\bin\TiaPortal.exe",
]

# 必需的 Python 包
REQUIRED_PACKAGES = [
    "pythonnet>=3.0.0",
    "openpyxl>=3.0.0",
]


# ════════════════════════════════════════════════════════════
#  检查函数
# ════════════════════════════════════════════════════════════

def check_python_version():
    """检查 Python 版本和架构"""
    version = sys.version_info
    bitness = "64位" if sys.maxsize > 2**32 else "32位"
    print(f"  Python 版本: {version.major}.{version.minor}.{version.micro}")
    print(f"  Python 架构: {bitness}")
    print(f"  Python 路径: {sys.executable}")

    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print("  [!] 警告: 建议 Python >= 3.8")
        return False
    return True


def check_python_package(package_name):
    """检查 Python 包是否已安装"""
    pkg_base = package_name.split(">=")[0].split("==")[0].strip()
    try:
        __import__(pkg_base)
        # 对于 pythonnet，导入名是 clr
        if pkg_base == "pythonnet":
            import clr
        print(f"  [✓] {pkg_base} 已安装")
        return True
    except ImportError:
        print(f"  [✗] {pkg_base} 未安装")
        return False


def check_tia_portal_installed():
    """检查 TIA Portal 是否已安装"""
    print("\n--- 检查 TIA Portal 安装 ---")
    found = False

    for api_path, version in TIA_POSSIBLE_PATHS:
        api_dir = Path(api_path)
        dll = api_dir / "Siemens.Engineering.dll"
        if dll.exists():
            print(f"  [✓] 找到 TIA Portal {version} Openness API")
            print(f"      路径: {api_path}")
            found = True
        else:
            print(f"  [ ] TIA Portal {version} Openness API 未找到 ({api_path})")

    if not found:
        # 搜索一下
        print("\n  [搜索] 在系统中搜索 Siemens.Engineering.dll...")
        found_paths = list(Path(r"C:\Program Files\Siemens").rglob("Siemens.Engineering.dll"))
        found_paths += list(Path(r"C:\Program Files (x86)\Siemens").rglob("Siemens.Engineering.dll"))
        if found_paths:
            print(f"  [!] 找到 DLL 但不在预期路径:")
            for p in found_paths[:3]:
                print(f"      {p}")
            print("  请更新 export_tia_variables.py 中的 TIA_VERSIONS 配置")
            found = True

    return found


def check_tia_running():
    """检查 TIA Portal 进程是否正在运行"""
    import subprocess
    print("\n--- 检查 TIA Portal 进程 ---")
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq TiaPortal.exe"],
            capture_output=True, text=True, timeout=10
        )
        if "TiaPortal.exe" in result.stdout:
            print("  [✓] TIA Portal 正在运行")
            return True
        else:
            print("  [ ] TIA Portal 未运行")
            print("     (导出变量前需要先启动 TIA Portal 并打开项目)")
            return False
    except Exception:
        print("  [!] 无法检查 TIA Portal 进程")
        return None


def check_pip():
    """检查 pip 可用性"""
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "--version"],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            print(f"  [✓] pip 可用: {result.stdout.strip()}")
            return True
    except Exception:
        pass
    print("  [✗] pip 不可用")
    return False


# ════════════════════════════════════════════════════════════
#  安装依赖
# ════════════════════════════════════════════════════════════

def install_dependencies():
    """安装所需的 Python 包"""
    print("\n" + "=" * 60)
    print("  安装所需依赖...")
    print("=" * 60)
    for pkg in REQUIRED_PACKAGES:
        print(f"\n安装 {pkg} ...")
        try:
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", pkg],
                capture_output=True, text=True, timeout=120
            )
            if result.returncode == 0:
                print(f"  [✓] {pkg} 安装成功")
            else:
                print(f"  [✗] {pkg} 安装失败")
                print(f"      {result.stderr.strip()[-300:]}")
        except Exception as e:
            print(f"  [✗] 安装出错: {e}")


# ════════════════════════════════════════════════════════════
#  主函数
# ════════════════════════════════════════════════════════════

def main():
    install_mode = "--install" in sys.argv or "-i" in sys.argv

    print("=" * 60)
    print("  TIA Portal Openness 环境检查")
    print("=" * 60)

    # 1. Python 版本检查
    print("\n--- Python 环境 ---")
    py_ok = check_python_version()

    # 2. pip 检查
    print("\n--- pip 状态 ---")
    pip_ok = check_pip()

    # 3. 包依赖检查
    print("\n--- Python 包依赖 ---")
    all_pkgs_ok = True
    for pkg in REQUIRED_PACKAGES:
        if not check_python_package(pkg):
            all_pkgs_ok = False

    # 4. TIA Portal 安装检查
    tia_installed = check_tia_portal_installed()

    # 5. TIA Portal 进程检查
    tia_running = check_tia_running()

    # 6. 安装模式
    if install_mode and not all_pkgs_ok:
        install_dependencies()
        # 重新检查
        print("\n--- 重新检查包依赖 ---")
        all_pkgs_ok = True
        for pkg in REQUIRED_PACKAGES:
            if not check_python_package(pkg):
                all_pkgs_ok = False

    # 7. 汇总
    print("\n" + "=" * 60)
    print("  检查汇总")
    print("=" * 60)

    checks = [
        ("Python 环境", py_ok),
        ("pip", pip_ok),
        ("Python 包依赖", all_pkgs_ok),
        ("TIA Portal 已安装", tia_installed),
    ]
    if tia_running is not None:
        checks.append(("TIA Portal 正在运行", tia_running))

    all_ok = True
    for name, status in checks:
        icon = "[✓]" if status else "[✗]"
        print(f"  {icon} {name}")
        if not status:
            all_ok = False

    print()
    if all_ok:
        print("[✓] 所有检查通过！可以运行导出脚本:")
        print(f"    python {Path(__file__).parent / 'export_tia_variables.py'}")
    else:
        print("[✗] 存在未满足的条件。")
        if not tia_installed:
            print("\n  TIA Portal Openness 安装说明:")
            print("  1. 运行 TIA Portal 安装程序")
            print("  2. 在组件选择中，勾选 'Openness' 选项")
            print("  3. 完成安装后重试")
        if not all_pkgs_ok:
            print(f"\n  请运行以下命令安装依赖:")
            print(f"    {sys.executable} -m pip install pythonnet openpyxl")
            print(f"  或使用本脚本的安装模式:")
            print(f"    python {__file__} --install")
        if not tia_running and tia_installed:
            print("\n  请先启动 TIA Portal 并打开项目，再运行导出脚本。")

    print()

    # 8. 项目文件路径提示
    print("--- 项目文件检查 ---")
    project_dirs = [
        Path(r"C:\Users\Administrator\Documents\Project\JSB-25-081B（瑞源橡塑）\JSB-25--081B(瑞源橡塑）TPV包纱管1.0"),
        Path(r"C:\Users\Administrator\Documents\Project\JSB-25-081B（瑞源橡塑）\针织机_V19（1214C+KTP700）"),
    ]
    for proj_dir in project_dirs:
        if proj_dir.exists():
            ap_files = list(proj_dir.glob("*.ap1*"))
            if ap_files:
                tia_ver = ap_files[0].suffix.replace(".ap", "V")
                print(f"  [✓] {proj_dir.name} (TIA Portal {tia_ver})")
            else:
                print(f"  [ ] {proj_dir.name} (无 .ap 文件)")
        else:
            print(f"  [ ] {proj_dir.name} (目录不存在)")


if __name__ == "__main__":
    main()
