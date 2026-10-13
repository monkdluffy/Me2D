import os
import sys
import shutil
import subprocess
from typing import Optional

try:
    from config import APP_NAME, APP_VERSION
except ImportError:
    APP_NAME = "Canido Doc Manager"
    APP_VERSION = "1.0.0"


def find_inno_setup_compiler() -> Optional[str]:
    """Detects Inno Setup Compiler (ISCC.exe) on the host machine."""
    # 1. Search in PATH
    iscc = shutil.which("ISCC.exe") or shutil.which("iscc")
    if iscc and os.path.exists(iscc):
        return iscc

    # 2. Search common default installation paths on Windows
    candidate_paths = [
        r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        r"C:\Program Files\Inno Setup 6\ISCC.exe",
        r"C:\Program Files (x86)\Inno Setup 5\ISCC.exe",
        r"C:\Program Files\Inno Setup 5\ISCC.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe")
    ]
    for path in candidate_paths:
        if os.path.exists(path):
            return path
    return None


def build_application():
    print("=" * 60)
    print(f"   Building {APP_NAME} (v{APP_VERSION})")
    print("   Mode: Fast-Launch OneDir + Standalone Inno Setup")
    print("=" * 60)

    project_root = os.path.dirname(os.path.abspath(__file__))
    assets_dir = os.path.join(project_root, "assets")
    is_win = sys.platform == "win32"
    data_sep = ";" if is_win else ":"

    # 1. Construct flet pack command
    # Menggunakan --onedir (-D) agar aplikasi tidak mengekstrak 300MB+ ke %TEMP% setiap kali dibuka.
    # Startup aplikasi menjadi INSTAN (< 1 detik).
    cmd_pack = [
        "flet", "pack", os.path.join(project_root, "app.py"),
        "--onedir",
        "--yes",
        "--name", APP_NAME,
        "--icon", os.path.join(assets_dir, "logo.ico"),
        "--product-name", APP_NAME,
        "--product-version", APP_VERSION,
        "--file-description", f"{APP_NAME} Document Generator & Manager",
        "--company-name", "Canido",
        "--copyright", "Copyright © 2026 Canido",
        "--add-data", f"{assets_dir}{data_sep}assets",
        "--hidden-import", "docxtpl",
        "--hidden-import", "docx",
        "--hidden-import", "fitz",
        "--hidden-import", "pymupdf",
        "--hidden-import", "pypdf",
        "--hidden-import", "pandas",
        "--hidden-import", "openpyxl",
        "--hidden-import", "core.updater"
    ]

    if is_win:
        cmd_pack.extend([
            "--hidden-import", "win32com",
            "--hidden-import", "win32com.client"
        ])

    print("\n[LANGKAH 1/2] Mengompilasi aplikasi dengan Flet (OneDir Mode)...")
    res_pack = subprocess.run(cmd_pack, cwd=project_root)

    if res_pack.returncode != 0:
        print("\n[ERROR] Proses compile flet pack gagal.")
        sys.exit(1)

    app_dir = os.path.join(project_root, "dist", APP_NAME)
    exe_name = f"{APP_NAME}.exe" if is_win else APP_NAME
    main_exe = os.path.join(app_dir, exe_name)

    print(f"\n[SUKSES] Bundle aplikasi OneDir berhasil dibuat:")
    print(f" -> Folder: {app_dir}")
    print(f" -> Executable: {main_exe}")

    # 2. Compile Inno Setup Installer (jika di Windows)
    if is_win:
        print("\n[LANGKAH 2/2] Memeriksa compiler installer (Inno Setup)...")
        iscc_path = find_inno_setup_compiler()
        iss_script = os.path.join(project_root, "installer.iss")

        if iscc_path and os.path.exists(iss_script):
            print(f" -> Ditemukan Inno Setup: {iscc_path}")
            print(f" -> Mengompilasi installer: {iss_script}")

            res_inno = subprocess.run([iscc_path, iss_script], cwd=project_root)
            if res_inno.returncode == 0:
                installer_output = os.path.join(project_root, "dist_installer", "Install_Canido_Doc_Manager.exe")
                print("\n" + "=" * 60)
                print("   SEMUA PROSES BUILD BERHASIL DISELESAIKAN!")
                print("=" * 60)
                if os.path.exists(installer_output):
                    size_mb = os.path.getsize(installer_output) / (1024 * 1024)
                    print(f"Installer siap didistribusikan:")
                    print(f" -> File: {installer_output} ({size_mb:.1f} MB)")
                    print(f" -> Keunggulan: Loading instan (<1 detik) saat user membuka aplikasi!")
                else:
                    print(f"Installer tersimpan di folder dist_installer/")
            else:
                print("\n[PERINGATAN] Kompilasi Inno Setup gagal. Anda dapat membuka installer.iss langsung lewat Inno Setup GUI.")
        else:
            print("\n[INFO] Inno Setup compiler (ISCC.exe) tidak ditemukan secara otomatis.")
            print("Anda bisa membuka file 'installer.iss' dengan aplikasi Inno Setup untuk membuat file installer setup.")
    else:
        print("\n[INFO] Untuk macOS/Linux, silakan gunakan paket app bundle atau format distribusi masing-masing.")


if __name__ == "__main__":
    build_application()
