import flet as ft
import app
import os
import sys

def get_assets_dir():
    try:
        if hasattr(sys, "_MEIPASS"):
            return os.path.join(sys._MEIPASS, "assets")
    except Exception:
        pass
    return os.path.abspath("assets")

def main(page: ft.Page):
    app.SMAPAutoDocApp(page)

if __name__ == "__main__":
    ft.app(target=main, assets_dir=get_assets_dir())
