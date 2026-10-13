import os
import sys
import shutil
import subprocess

def build_application():
    print("==================================================")
    print("   Building SMAP AutoDoc Studio Executable        ")
    print("          Using flet pack (PyInstaller)           ")
    print("==================================================")

    project_root = os.path.dirname(os.path.abspath(__file__))
    assets_dir = os.path.join(project_root, "assets")

    is_win = sys.platform == "win32"
    data_sep = ";" if is_win else ":"

    # Command for Flet pack
    cmd_pack = [
        "flet", "pack", os.path.join(project_root, "app.py"),
        "--name", "Canido Doc Manager",
        "--icon", os.path.join(assets_dir, "logo.ico"),
        "--product-name", "Canido Doc Manager",
        "--product-version", "1.0.0",
        "--company-name", "Canido",
        "--copyright", "Copyright © 2026 Canido",
        "--add-data", f"{assets_dir}{data_sep}assets",
        "--hidden-import", "docxtpl",
        "--hidden-import", "docx",
        "--hidden-import", "fitz",
        "--hidden-import", "pymupdf",
        "--hidden-import", "pypdf",
        "--hidden-import", "pandas",
        "--hidden-import", "openpyxl"
    ]

    if is_win:
        cmd_pack.extend([
            "--hidden-import", "win32com",
            "--hidden-import", "win32com.client"
        ])

    print("\n[BUILD] Running flet pack...")
    res_pack = subprocess.run(cmd_pack, cwd=project_root)

    if res_pack.returncode == 0:
        print("\n[SUCCESS] Build Completed Successfully!")
        output_exe = os.path.join(project_root, "dist", "Canido Doc Manager.exe" if is_win else "Canido Doc Manager")
        print(f"Executable output: {output_exe}")
        print("Anda sekarang bisa menjalankan Inno Setup (installer.iss) untuk membuat installer.")
    else:
        print("\n[ERROR] Build encountered issues.")
        sys.exit(1)

if __name__ == "__main__":
    build_application()
