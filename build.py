#!/usr/bin/env python3
"""一键打包: python build.py"""
import sys, shutil, subprocess, sysconfig
from pathlib import Path

BASE = Path(__file__).resolve().parent
EXE_NAME = "SUATCatManager"
PIP_MIRROR = "https://mirrors.sustech.edu.cn/pypi/web/simple"

REQUIRED = [
    ("PIL","Pillow"), ("webview","pywebview"), ("numpy","numpy"),
    ("pythonnet","pythonnet"), ("PyInstaller","pyinstaller"),
]

def step(msg): print(f"\n{'='*50}\n  {msg}\n{'='*50}")

def check_deps():
    missing = []
    for mod, pkg in REQUIRED:
        try: __import__(mod); print(f"  \u2713 {pkg}")
        except ImportError: print(f"  \u2717 {pkg}"); missing.append(pkg)
    if not missing: print("  全部就绪"); return
    if sys.stdin.isatty():
        ans = input(f"  安装 {missing}？[Y/n] ").strip().lower()
        if ans in ("","y","yes"):
            for p in missing:
                subprocess.run([sys.executable,"-m","pip","install",p,
                    "-i",PIP_MIRROR,"--trusted-host","mirrors.sustech.edu.cn"], check=True)

def collect_conda_dlls():
    """conda 把 OpenSSL / OpenBLAS 等运行时 DLL 放在 <env>/Library/bin，
    而标准 CPython 没有这个目录。这些 DLL 不会被 --collect-all 自动收集，
    缺失会导致 _ssl / numpy 等 ImportError: DLL load failed。
    仅在 conda 环境下(存在 Library/bin)才收集,venv 下不触发、保持精简。"""
    lib_bin = Path(sys.prefix) / "Library" / "bin"
    if not lib_bin.is_dir():
        return []
    out = [str(f) for f in lib_bin.glob("*.dll")]
    print(f"  conda 运行时 DLL: {len(out)} 个 (来自 {lib_bin})")
    return out

def build():
    step("打包")
    old = BASE / f"{EXE_NAME}.exe"
    if old.exists():
        try: old.unlink()
        except PermissionError:
            print("请先关闭旧EXE"); return

    # 收集 Python DLLs/ 下所有 .dll + .pyd (stdlib 扩展模块)
    dll_dir = Path(sysconfig.get_config_var("BINDIR") or sys.prefix) / "DLLs"
    dll_binaries = []
    if dll_dir.is_dir():
        for f in dll_dir.iterdir():
            if f.suffix.lower() in (".dll", ".pyd"):
                dll_binaries.append(str(f))
    print(f"  收集 {len(dll_binaries)} 个系统 DLL")

    # conda 环境额外需要的运行时 DLL(修复 _ssl 等 DLL load failed)
    conda_dlls = collect_conda_dlls()

    binaries = dll_binaries + conda_dlls

    cmd = [sys.executable,"-m","PyInstaller",
        "--onefile","--windowed","--name",EXE_NAME,
        "--distpath",str(BASE),"--workpath",str(BASE/"build"),
        "--collect-all","pythonnet","--collect-all","webview","--collect-all","ssl",
        "--hidden-import","clr",
        "--add-data",f"{BASE/'app_manager.html'};.",
        "--add-data",f"{BASE/'marked.min.js'};."]
    for dll in binaries:
        cmd += ["--add-binary", f"{dll};."]
    cmd.append(str(BASE/"app.py"))

    # 精灵图
    subprocess.run([sys.executable, str(BASE/"old_tool"/"make_sprite.py")], check=True)
    # 打包
    subprocess.run(cmd, check=True)

    # 清理
    shutil.rmtree(BASE/"build", ignore_errors=True)
    for s in BASE.glob("*.spec"): s.unlink()
    for pyc in BASE.rglob("__pycache__"): shutil.rmtree(pyc, ignore_errors=True)

    exe = BASE / f"{EXE_NAME}.exe"
    print(f"\n  \u2713 完成: {exe.name} ({exe.stat().st_size/1024/1024:.0f}MB)")
    print("  双击运行")

if __name__ == "__main__":
    step("依赖检查")
    check_deps()
    build()
