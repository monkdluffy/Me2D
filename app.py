import os
import sys
import shutil
import threading
import subprocess
import logging
from typing import Dict, List, Any, Optional
import flet as ft

from config import (
    APP_NAME, APP_VERSION, SMAP_VARIABLE_CATEGORIES, ALL_SMAP_VARIABLES,
    DEFAULT_VARIABLE_VALUES, IMAGE_SLOTS, find_libreoffice_path
)
from core.doc_engine import DocumentEngine
from core.pdf_engine import PDFEngine
from core.image_engine import ImageEngine
from core.profile_manager import ProfileManager
from core.db_manager import DatabaseManager
from core.batch_worker import BatchProcessor
from core.updater import check_for_updates, download_update_file, launch_installer_and_exit

# Setup Logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
logger = logging.getLogger("app")


class SMAPAutoDocApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = f"{APP_NAME} v{APP_VERSION}"
        self.page.window.width = 1280
        self.page.window.height = 850
        self.page.window.min_width = 1024
        self.page.window.min_height = 700
        
        # Load User Settings early so we can apply theme
        self.user_settings_path = os.path.join(os.getcwd(), 'user_settings.json')
        self.load_user_settings()
        
        self.page.theme_mode = ft.ThemeMode.DARK if self.app_theme_mode == "Dark" else ft.ThemeMode.LIGHT
        self.page.padding = 0
        # Window icon requires absolute path for Flet desktop runner
        def get_asset_path(filename):
            if hasattr(sys, "_MEIPASS"):
                p1 = os.path.join(sys._MEIPASS, "assets", filename)
                if os.path.exists(p1):
                    return p1
            if getattr(sys, "frozen", False):
                exe_dir = os.path.dirname(sys.executable)
                p2 = os.path.join(exe_dir, "assets", filename)
                if os.path.exists(p2):
                    return p2
            return os.path.abspath(os.path.join("assets", filename))
            
        if os.path.exists(get_asset_path("title_logo.ico")):
            self.page.window.icon = get_asset_path("title_logo.ico")
        elif os.path.exists(get_asset_path("title_logo.png")):
            self.page.window.icon = get_asset_path("title_logo.png")
        else:
            self.page.window.icon = get_asset_path("logo.png")
            
        # State variables for dynamic fields
        self.serti_count = 1

        self.sbu_count = 1
        
        # Form State
        self.form_fields: Dict[str, ft.TextField] = {}
        self.find_replace_fields: List[Dict[str, Any]] = []
        self.image_fields: Dict[str, Dict[str, Any]] = {}
        self.selected_single_template: str = ""
        self.batch_templates: List[str] = []
        self.batch_data_records: List[Dict[str, Any]] = []
        self.batch_output_dir: str = os.path.abspath("output_documents")
        self.batch_processor = BatchProcessor(max_workers=4)
        self.libreoffice_path: str = find_libreoffice_path()
        
        # Database Manager
        self.db = DatabaseManager()
        
        # Build File Pickers
        self.setup_file_pickers()

        # Initialize Image Slots State
        for tag_name, slot in IMAGE_SLOTS.items():
            self.image_fields[tag_name] = {
                "path_field": ft.TextField(label=f"File {slot['label']}", value="", expand=True, read_only=True,),
                "h_field": ft.TextField(label="Height (mm)", value=str(slot["default_height_mm"]), width=90, on_focus=self._on_focus_select_all),
                "preview_img": ft.Image(src="", width=80, height=50, fit="contain", visible=False)
            }

        # Build UI Structure
        self.build_ui()

        # Load Default Sample Data on Start
        self.load_sample_data(show_snack=False)
        
    def load_user_settings(self):
        self.removebg_api_key = "ASgAWQTBoJEHjKHY4B6JGRpb" # Default provided by user
        self.app_theme_mode = "Light"
        self.default_output_dir = os.path.abspath("output_documents")
        self.company_name = "Canido Sertifikasi"
        self.company_address = ""
        self.auto_save = True

        try:
            if os.path.exists(self.user_settings_path):
                import json
                with open(self.user_settings_path, 'r') as f:
                    data = json.load(f)
                    if "removebg_api_key" in data:
                        self.removebg_api_key = data["removebg_api_key"]
                    if "app_theme_mode" in data:
                        self.app_theme_mode = data["app_theme_mode"]
                    if "default_output_dir" in data:
                        self.default_output_dir = data["default_output_dir"]
                    if "company_name" in data:
                        self.company_name = data["company_name"]
                    if "company_address" in data:
                        self.company_address = data["company_address"]
                    if "auto_save" in data:
                        self.auto_save = data["auto_save"]
        except Exception as e:
            logger.error(f"Failed to load user settings: {e}")

    def save_user_settings(self):
        try:
            import json
            data = {
                "removebg_api_key": getattr(self, 'txt_settings_removebg_key', ft.TextField(value=self.removebg_api_key)).value,
                "app_theme_mode": self.app_theme_mode,
                "default_output_dir": self.default_output_dir,
                "company_name": self.company_name,
                "company_address": self.company_address,
                "auto_save": self.auto_save
            }
            with open(self.user_settings_path, 'w') as f:
                json.dump(data, f)
        except Exception as e:
            logger.error(f"Failed to save user settings: {e}")

    def setup_file_pickers(self):
        """Initializes Flet FilePicker instances and overlays."""
        # File Pickers
        self.picker_json_save = ft.FilePicker()
        self.picker_json_load = ft.FilePicker()
        self.picker_excel_export = ft.FilePicker()
        self.picker_excel_import = ft.FilePicker()
        
        self.picker_batch_templates = ft.FilePicker()
        self.picker_batch_folder = ft.FilePicker()
        self.picker_batch_outdir = ft.FilePicker()
        self.picker_libreoffice = ft.FilePicker()

        # PDF Toolkit Pickers
        self.picker_toolkit_docx = ft.FilePicker()
        self.picker_toolkit_watermark_pdf = ft.FilePicker()
        self.picker_toolkit_stamp_pdf = ft.FilePicker()
        self.picker_toolkit_stamp_img = ft.FilePicker()
        self.picker_toolkit_merge_pdfs = ft.FilePicker()
        self.picker_toolkit_merge_indir = ft.FilePicker()
        self.picker_toolkit_merge_outdir = ft.FilePicker()
        self.picker_toolkit_split_pdf = ft.FilePicker()
        self.picker_toolkit_split_outdir = ft.FilePicker()

        # Image Toolkit Pickers
        self.picker_bg_remover_in = ft.FilePicker()
        self.picker_bg_remover_outdir = ft.FilePicker()

        # Pickers in Flet 0.8x don't need to be appended to page.overlay
        # (This prevents 'Unknown control: FilePicker' error on the client)

        # Dynamic Image Slot Pickers
        self.image_pickers: Dict[str, ft.FilePicker] = {}
        for tag in IMAGE_SLOTS.keys():
            picker = ft.FilePicker()
            self.image_pickers[tag] = picker

    def show_snackbar(self, message: str, color: str = "blue"):
        """Displays a snackbar notification."""
        snack = ft.SnackBar(
            content=ft.Text(message, color=ft.Colors.WHITE, weight=ft.FontWeight.BOLD),
            bgcolor=ft.Colors.GREEN_700 if color == "green" else (ft.Colors.RED_700 if color == "red" else ft.Colors.BLUE_700)
        )
        try:
            self.page.open(snack)
        except AttributeError:
            self.page.snack_bar = snack
            snack.open = True
            self.page.update()
    def build_ui(self):
        """Constructs the primary responsive layout with NavigationRail and View Containers."""
        # Navigation Rail
        self.nav_rail = ft.NavigationRail(
            selected_index=0,
            label_type=ft.NavigationRailLabelType.ALL,
            min_width=100,
            min_extended_width=200,
            bgcolor=ft.Colors.TRANSPARENT,
            leading=ft.Container(
                content=ft.Image(src="logo.png", width=120, height=120, fit="contain"),
                padding=10
            ),
            destinations=[
                ft.NavigationRailDestination(icon=ft.Icons.AUTO_AWESOME, selected_icon=ft.Icons.AUTO_AWESOME, label="SMAP Automation"),
                ft.NavigationRailDestination(icon=ft.Icons.CHECK_CIRCLE, selected_icon=ft.Icons.CHECK_CIRCLE, label="PANCEK Automation"),
                ft.NavigationRailDestination(icon=ft.Icons.PICTURE_IN_PICTURE, selected_icon=ft.Icons.PICTURE_IN_PICTURE, label="PDF Toolkit"),
                ft.NavigationRailDestination(icon=ft.Icons.IMAGE, selected_icon=ft.Icons.IMAGE, label="Image Toolkit"),
                ft.NavigationRailDestination(icon=ft.Icons.SETTINGS, selected_icon=ft.Icons.SETTINGS, label="Settings"),
                ft.NavigationRailDestination(icon=ft.Icons.HELP, selected_icon=ft.Icons.HELP, label="Help"),
            ],
            on_change=self.on_nav_change,
            expand=True
        )

        self.exit_btn = ft.Container(
            content=ft.Column([
                ft.Icon(ft.Icons.EXIT_TO_APP, color=ft.Colors.RED_700, size=24),
                ft.Text("Exit App", size=12, color=ft.Colors.RED_700, weight=ft.FontWeight.BOLD)
            ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=2),
            on_click=self.on_exit_click,
            tooltip="Exit Application",
            ink=True,
            padding=20,
            alignment=ft.alignment.Alignment(0, 0)
        )

        self.sidebar_container = ft.Container(
            content=ft.Column([
                self.nav_rail,
                self.exit_btn
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            bgcolor=ft.Colors.BLUE_50
        )

        # Auto-Updater Controls & State
        self.txt_update_status = ft.Text(f"Versi saat ini: v{APP_VERSION} (Up-to-date)", size=13, color=ft.Colors.GREY_700)
        self.btn_check_update = ft.ElevatedButton(
            "Periksa Pembaruan Sekarang",
            icon=ft.Icons.REFRESH,
            style=ft.ButtonStyle(bgcolor=ft.Colors.BLUE_700, color=ft.Colors.WHITE),
            on_click=self.on_manual_check_update_clicked
        )

        # Build View Containers
        self.input_data_view = self.build_input_data_view()
        self.batch_view = self.build_batch_view()
        self.presets_view = self.build_presets_view()
        self.history_view = self.build_history_view()
        self.pdf_toolkit_view = self.build_pdf_toolkit_view()
        self.image_toolkit_view = self.build_image_toolkit_view()
        self.settings_view = self.build_settings_view()
        self.help_view = self.build_help_view()
        
        # Build SMAP Automation View (Tabs)
        self.smap_tab_bar = ft.TabBar(
            tabs=[
                ft.Tab(label="Input Data", icon=ft.Icons.EDIT_DOCUMENT),
                ft.Tab(label="Process Documents", icon=ft.Icons.DASHBOARD_CUSTOMIZE),
                ft.Tab(label="Presets & Data", icon=ft.Icons.FOLDER_SHARED),
                ft.Tab(label="Database History", icon=ft.Icons.HISTORY)
            ]
        )
        self.smap_tab_view = ft.TabBarView(
            expand=True,
            controls=[
                self.input_data_view,
                self.batch_view,
                self.presets_view,
                self.history_view
            ]
        )
        self.smap_automation_view = ft.Tabs(
            length=4,
            expand=True,
            content=ft.Column(
                expand=True,
                controls=[
                    self.smap_tab_bar,
                    ft.Divider(height=1),
                    self.smap_tab_view
                ]
            )
        )
        self.smap_tabs = self.smap_automation_view

        # Build PANCEK Placeholder
        self.pancek_automation_view = ft.Container(
            content=ft.Column([
                ft.Icon(ft.Icons.CONSTRUCTION, size=80, color=ft.Colors.ORANGE_500),
                ft.Text("PANCEK Automation", size=28, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
                ft.Text("Halaman sedang dalam pengembangan. Silakan siapkan data Anda terlebih dahulu.", size=16, color=ft.Colors.GREY_700)
            ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=20),
            expand=True,
            alignment=ft.alignment.Alignment(0, 0)
        )

        # View Switcher Container
        self.content_area = ft.Container(
            content=self.smap_automation_view,
            expand=True,
            padding=15,
            bgcolor=ft.Colors.BLUE_50
        )

        # Main Layout
        self.page.add(
            ft.Row([
                self.sidebar_container,
                ft.VerticalDivider(width=1),
                self.content_area
            ], expand=True, spacing=0)
        )

        # Check for auto updates asynchronously on startup
        threading.Thread(target=self._check_update_on_startup, daemon=True).start()

    def on_nav_change(self, e):
        idx = e.control.selected_index
        views = [
            self.smap_automation_view,
            self.pancek_automation_view,
            self.pdf_toolkit_view,
            self.image_toolkit_view,
            self.settings_view,
            self.help_view
        ]
        bg_colors = [
            ft.Colors.BLUE_50,
            ft.Colors.ORANGE_50,
            ft.Colors.BLUE_GREY_50,
            ft.Colors.INDIGO_50,
            ft.Colors.TEAL_50,
            ft.Colors.DEEP_PURPLE_50
        ]
        self.content_area.content = views[idx]
        self.content_area.bgcolor = bg_colors[idx]
        self.sidebar_container.bgcolor = bg_colors[idx]
        self.page.update()

    async def on_exit_click(self, e):
        try:
            if hasattr(self.page.window, 'destroy'):
                res = self.page.window.destroy()
                if hasattr(res, '__await__'):
                    await res
            elif hasattr(self.page, 'window_destroy'):
                res = self.page.window_destroy()
                if hasattr(res, '__await__'):
                    await res
        except Exception:
            pass

    # ==========================================
    # 1. INPUT DATA VIEW
    # ==========================================
    def build_input_data_view(self) -> ft.Control:

        # Variable Accordion/Tabs Matrix
        tab_labels = []
        tab_contents = []
        categories = list(SMAP_VARIABLE_CATEGORIES.keys())
        for cat_idx, cat_name in enumerate(categories):
            vars_dict = SMAP_VARIABLE_CATEGORIES[cat_name]
            controls_list = []
            for var_name, meta in vars_dict.items():
                if meta.get("type") == "dropdown" or var_name == "jenis_usaha":
                    options_list = meta.get("options", [
                        "Jasa Pekerjaan Kontruksi",
                        "Jasa Konsultansi Konstruksi"
                    ])
                    default_val = meta.get("default") or options_list[0]
                    dd = ft.Dropdown(
                        label=meta["label"],
                        value=default_val,
                        options=[ft.dropdown.Option(opt) for opt in options_list],
                        expand=True
                    )
                    self.form_fields[var_name] = dd
                    controls_list.append(dd)
                else:
                    tf = ft.TextField(
                        label=meta['label'],
                        value=meta['default'],
                        expand=True,
                        on_focus=self._on_focus_select_all
                    )
                    self.form_fields[var_name] = tf
                    controls_list.append(tf)

            # Responsive Grid Layout for Tab
            grid_columns = []
            i = 0
            while i < len(controls_list):
                c1 = controls_list[i]
                c2 = controls_list[i+1] if i + 1 < len(controls_list) else None
                
                # Check if c2 is the start of a special group. If so, c1 must be placed alone
                # so that c2 can start at the beginning of the next row.
                if c2 and c2.label in ["Sertifikat Standar #1", "SBU / UMKU #1", "FKAP"]:
                    grid_columns.append(ft.Row([c1]))
                    i += 1
                    continue
                    
                row_items = [c1]
                if c2:
                    row_items.append(c2)
                    
                if row_items[0].label == "Sertifikat Standar #1":
                    grid_columns.append(ft.Divider(height=1, color=ft.Colors.BLUE_200))
                    self.serti_container = ft.Column([ft.Row(row_items)], spacing=10)
                    grid_columns.append(self.serti_container)
                    btn_add_serti = ft.ElevatedButton("Tambah Sertifikat Standar", icon=ft.Icons.ADD, on_click=self.add_serti_standar)
                    grid_columns.append(ft.Row([btn_add_serti]))
                    grid_columns.append(ft.Divider(height=1, color=ft.Colors.BLUE_200))
                elif row_items[0].label == "SBU / UMKU #1":
                    grid_columns.append(ft.Divider(height=1, color=ft.Colors.BLUE_200))
                    self.sbu_container = ft.Column([ft.Row(row_items)], spacing=10)
                    grid_columns.append(self.sbu_container)
                    btn_add_sbu = ft.ElevatedButton("Tambah SBU / UMKU", icon=ft.Icons.ADD, on_click=self.add_sbu_umku)
                    grid_columns.append(ft.Row([btn_add_sbu]))
                    grid_columns.append(ft.Divider(height=1, color=ft.Colors.BLUE_200))
                else:
                    grid_columns.append(ft.Row(row_items))
                    
                i += len(row_items)

            if cat_name == "Data Badan Usaha":
                # Batch Find & Replace Controls (Dynamic Multiple Form 1 & Form 2 pairs)
                self.find_replace_fields = []

                # Default #1 row (Default count is 1 every time app starts or resets)
                self.txt_batch_find_word = ft.TextField(
                    label="Form 1: Kata / Kalimat yang Dicari #1",
                    hint_text="Contoh: Jasa Konstruksi, PT ABC, atau kata/kalimat yang ingin diganti...",
                    expand=True,
                    on_focus=self._on_focus_select_all
                )
                self.txt_batch_replace_word = ft.TextField(
                    label="Form 2: Kata / Kalimat Pengganti #1",
                    hint_text="Kata atau kalimat baru pengganti...",
                    expand=True,
                    on_focus=self._on_focus_select_all
                )
                row_fr_1 = ft.Row(spacing=15, vertical_alignment=ft.CrossAxisAlignment.CENTER)
                field_dict_1 = {
                    "find": self.txt_batch_find_word,
                    "replace": self.txt_batch_replace_word,
                    "row": row_fr_1
                }
                btn_del_1 = ft.IconButton(
                    icon=ft.Icons.CLEAR,
                    icon_color=ft.Colors.GREY_400,
                    tooltip="Kosongkan baris pencarian",
                    on_click=lambda _: self.remove_find_replace_row(field_dict_1)
                )
                field_dict_1["btn_del"] = btn_del_1
                row_fr_1.controls = [self.txt_batch_find_word, self.txt_batch_replace_word, btn_del_1]

                self.find_replace_fields.append(field_dict_1)
                self.find_replace_container = ft.Column([row_fr_1], spacing=10)

                btn_add_find_replace = ft.ElevatedButton(
                    "Tambah Pencarian / Penggantian Kata",
                    icon=ft.Icons.ADD,
                    on_click=self.add_find_replace_row
                )

                self.chk_batch_case_sensitive = ft.Checkbox(
                    label="Case Sensitive (Pencarian Peka Huruf Besar & Kecil)",
                    value=True
                )

                find_replace_bu_card = ft.Card(
                    content=ft.Container(
                        content=ft.Column([
                            ft.Row([
                                ft.Icon(ft.Icons.FIND_REPLACE, color=ft.Colors.BLUE_800, size=24),
                                ft.Text("Pencarian & Penggantian Kata / Kalimat (Batch Find & Replace):", weight=ft.FontWeight.BOLD, size=15, color=ft.Colors.BLUE_900),
                            ]),
                            ft.Text(
                                "Setelah memeriksa kelengkapan Data Badan Usaha di atas, tentukan kata/kalimat yang ingin dicari (Form 1) dan diganti (Form 2). Aturan ini akan otomatis dijalankan saat proses penggantian dokumen dieksekusi di tab Process Documents.",
                                size=12,
                                color=ft.Colors.GREY_700
                            ),
                            self.find_replace_container,
                            ft.Row([btn_add_find_replace]),
                            ft.Row([
                                self.chk_batch_case_sensitive,
                                ft.Text("(Default: Case Sensitive aktif. Opsi konfirmasi sensitivitas juga muncul saat tombol Start Batch Replacement diklik)", size=12, italic=True, color=ft.Colors.GREY_600)
                            ])
                        ], spacing=12),
                        padding=15
                    )
                )

                grid_columns.append(ft.Divider(height=2, color=ft.Colors.BLUE_200))
                grid_columns.append(find_replace_bu_card)
                grid_columns.append(ft.Divider(height=1, color=ft.Colors.BLUE_100))

            # Add Next Button
            next_idx = cat_idx + 1
            if next_idx < len(categories):
                next_cat_name = categories[next_idx]
                btn_next = ft.ElevatedButton(
                    f"Selanjutnya: {next_cat_name} ➔",
                    on_click=lambda e, i=next_idx: self.switch_form_tab(i),
                    style=ft.ButtonStyle(color=ft.Colors.WHITE, bgcolor=ft.Colors.BLUE_700)
                )
            else:
                btn_next = ft.ElevatedButton(
                    "Selanjutnya: Image Assets ➔",
                    on_click=lambda e, i=len(categories): self.switch_form_tab(i),
                    style=ft.ButtonStyle(color=ft.Colors.WHITE, bgcolor=ft.Colors.BLUE_700)
                )

            grid_columns.append(ft.Row([ft.Container(expand=True), btn_next]))

            tab_labels.append(ft.Tab(label=cat_name))
            inner_col = ft.Column(grid_columns, spacing=20)
            col = ft.Column(
                [ft.Container(content=inner_col, padding=ft.padding.Padding(left=0, top=0, right=15, bottom=0))],
                scroll=ft.ScrollMode.AUTO
            )
                
            tab_contents.append(
                ft.Container(
                    content=col,
                    padding=ft.padding.Padding(left=15, top=25, right=5, bottom=15)
                )
            )

        # Image Assets Tab
        img_rows = []
        for tag_name, slot in IMAGE_SLOTS.items():
            fields = self.image_fields[tag_name]
            async def _img_click(e, t=tag_name, lbl=slot['label']):
                await self.on_image_click(e, t, f"Select {lbl}")

            btn_pick = ft.ElevatedButton(
                "Browse Image",
                icon=ft.Icons.IMAGE,
                on_click=_img_click
            )
            img_rows.append(
                ft.Row([
                    ft.Text(slot['label'], weight=ft.FontWeight.BOLD, width=200),
                    fields["path_field"],
                    fields["h_field"],
                    btn_pick,
                    fields["preview_img"]
                ], alignment=ft.MainAxisAlignment.CENTER)
            )

        btn_next_process = ft.ElevatedButton(
            "Simpan Data & Lanjut ke Process Documents ➔",
            on_click=lambda e: self.nav_to_process_documents(),
            style=ft.ButtonStyle(color=ft.Colors.WHITE, bgcolor=ft.Colors.GREEN_700)
        )
        img_rows.append(ft.Row([ft.Container(expand=True), btn_next_process]))

        tab_labels.append(ft.Tab(label="Image Assets (Logo/Stamp/TTD)"))
        inner_img_col = ft.Column(img_rows, spacing=15)
        scroll_col = ft.Column(
            [ft.Container(content=inner_img_col, padding=ft.padding.Padding(left=0, top=0, right=15, bottom=0))],
            scroll=ft.ScrollMode.AUTO
        )

        tab_contents.append(
            ft.Container(
                content=scroll_col,
                padding=ft.padding.Padding(left=15, top=25, right=5, bottom=15)
            )
        )

        self.form_tabs = ft.Tabs(
            length=len(tab_labels),
            expand=True,
            content=ft.Column(
                expand=True,
                controls=[
                    ft.TabBar(tabs=tab_labels),
                    ft.TabBarView(controls=tab_contents, expand=True)
                ]
            )
        )

        return ft.Column([
            ft.Row([
                ft.Text("Input Data Badan Usaha", size=20, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
                ft.Container(expand=True),
                ft.OutlinedButton("Reset Form", icon=ft.Icons.REFRESH, on_click=lambda _: self.reset_form_data()),
            ]),
            self.form_tabs
        ], expand=True, spacing=10)

    def switch_form_tab(self, index: int):
        self.form_tabs.selected_index = index
        self.page.update()

    def ensure_dynamic_field(self, key: str):
        import re
        if m := re.match(r'^(?:tgl_)?serti_standar_(\d+)$', key):
            idx = int(m.group(1))
            while self.serti_count < idx:
                self.add_serti_standar(None, update=False)
        elif m := re.match(r'^(?:tgl_)?sbu_umku_(\d+)$', key):
            idx = int(m.group(1))
            while self.sbu_count < idx:
                self.add_sbu_umku(None, update=False)

    def add_serti_standar(self, e, update=True):
        self.serti_count += 1
        idx = self.serti_count
        
        tf1 = ft.TextField(label=f"Sertifikat Standar #{idx}", value="", expand=True, on_focus=self._on_focus_select_all)
        tf2 = ft.TextField(label=f"Tgl Sertifikat Standar #{idx}", value="", expand=True, on_focus=self._on_focus_select_all)
        
        self.form_fields[f"serti_standar_{idx}"] = tf1
        self.form_fields[f"tgl_serti_standar_{idx}"] = tf2
        
        # Append to the dedicated container
        if hasattr(self, 'serti_container') and self.serti_container is not None:
            self.serti_container.controls.append(ft.Row([tf1, tf2]))
            
        if update:
            self.page.update()

    def add_sbu_umku(self, e, update=True):
        self.sbu_count += 1
        idx = self.sbu_count
        
        tf1 = ft.TextField(label=f"SBU / UMKU #{idx}", value="", expand=True, on_focus=self._on_focus_select_all)
        tf2 = ft.TextField(label=f"Tgl SBU / UMKU #{idx}", value="", expand=True, on_focus=self._on_focus_select_all)
        
        self.form_fields[f"sbu_umku_{idx}"] = tf1
        self.form_fields[f"tgl_sbu_umku_{idx}"] = tf2
        
        # Append to the dedicated container
        if hasattr(self, 'sbu_container') and self.sbu_container is not None:
            self.sbu_container.controls.append(ft.Row([tf1, tf2]))
            
        if update:
            self.page.update()

    def add_find_replace_row(self, e=None, update=True):
        idx = len(self.find_replace_fields) + 1
        tf_find = ft.TextField(
            label=f"Form 1: Kata / Kalimat yang Dicari #{idx}",
            hint_text="Kata atau kalimat yang ingin dicari...",
            expand=True,
            on_focus=self._on_focus_select_all
        )
        tf_replace = ft.TextField(
            label=f"Form 2: Kata / Kalimat Pengganti #{idx}",
            hint_text="Kata atau kalimat baru pengganti...",
            expand=True,
            on_focus=self._on_focus_select_all
        )

        row_item = ft.Row(spacing=15, vertical_alignment=ft.CrossAxisAlignment.CENTER)
        field_dict = {
            "find": tf_find,
            "replace": tf_replace,
            "row": row_item
        }

        btn_del = ft.IconButton(
            icon=ft.Icons.DELETE_OUTLINE,
            icon_color=ft.Colors.RED_600,
            tooltip=f"Hapus baris pencarian #{idx}",
            on_click=lambda _: self.remove_find_replace_row(field_dict)
        )
        field_dict["btn_del"] = btn_del

        row_item.controls = [tf_find, tf_replace, btn_del]

        self.find_replace_fields.append(field_dict)
        if hasattr(self, 'find_replace_container') and self.find_replace_container is not None:
            self.find_replace_container.controls.append(row_item)

        self._refresh_find_replace_ui(update=update)

    def remove_find_replace_row(self, field_dict):
        if len(self.find_replace_fields) > 1:
            if field_dict in self.find_replace_fields:
                self.find_replace_fields.remove(field_dict)
            if hasattr(self, 'find_replace_container') and self.find_replace_container is not None:
                if field_dict.get("row") in self.find_replace_container.controls:
                    self.find_replace_container.controls.remove(field_dict["row"])
            # Update legacy pointers to active first row
            if self.find_replace_fields:
                self.txt_batch_find_word = self.find_replace_fields[0]["find"]
                self.txt_batch_replace_word = self.find_replace_fields[0]["replace"]
        else:
            # If only 1 row remains, clear its values instead of deleting the row
            field_dict["find"].value = ""
            field_dict["replace"].value = ""

        self._refresh_find_replace_ui(update=True)

    def _refresh_find_replace_ui(self, update=True):
        # Renumber labels and update action icons/tooltips so all rows stay symmetrical
        for i, item in enumerate(self.find_replace_fields, 1):
            item["find"].label = f"Form 1: Kata / Kalimat yang Dicari #{i}"
            item["replace"].label = f"Form 2: Kata / Kalimat Pengganti #{i}"
            if "btn_del" in item and item["btn_del"]:
                if len(self.find_replace_fields) == 1:
                    item["btn_del"].icon = ft.Icons.CLEAR
                    item["btn_del"].icon_color = ft.Colors.GREY_400
                    item["btn_del"].tooltip = "Kosongkan baris pencarian"
                else:
                    item["btn_del"].icon = ft.Icons.DELETE_OUTLINE
                    item["btn_del"].icon_color = ft.Colors.RED_600
                    item["btn_del"].tooltip = f"Hapus baris pencarian #{i}"

        if update:
            if hasattr(self, 'find_replace_container') and self.find_replace_container is not None:
                try:
                    self.find_replace_container.update()
                except Exception:
                    if hasattr(self, 'page'):
                        self.page.update()
            elif hasattr(self, 'page'):
                self.page.update()

    def get_all_find_replace_pairs(self) -> List[Dict[str, str]]:
        pairs = []
        if hasattr(self, 'find_replace_fields') and self.find_replace_fields:
            for item in self.find_replace_fields:
                f_text = item["find"].value.strip() if item["find"].value else ""
                r_text = item["replace"].value.strip() if item["replace"].value else ""
                if f_text:
                    pairs.append({"find_text": f_text, "replace_text": r_text})
        elif hasattr(self, 'txt_batch_find_word') and self.txt_batch_find_word.value:
            f_text = self.txt_batch_find_word.value.strip()
            r_text = self.txt_batch_replace_word.value.strip() if self.txt_batch_replace_word.value else ""
            if f_text:
                pairs.append({"find_text": f_text, "replace_text": r_text})
        return pairs

    def nav_to_process_documents(self):
        # Programmatically change to SMAP Automation > Process Documents tab
        if hasattr(self, 'radio_data_source'):
            self.radio_data_source.value = "form"
            
        self.nav_rail.selected_index = 0
        self.smap_tabs.selected_index = 1
        self.content_area.content = self.smap_automation_view
        self.page.update()
        self.show_snackbar("Data siap untuk diproses! Anda telah berada di halaman Process Documents.", color="green")

    def switch_to_input_data_bu(self):
        # Programmatically change to SMAP Automation > Input Data > Data Badan Usaha tab
        self.nav_rail.selected_index = 0
        self.smap_tabs.selected_index = 0
        if hasattr(self, 'form_tabs') and self.form_tabs:
            self.form_tabs.selected_index = 0
        self.content_area.content = self.smap_automation_view
        self.page.update()

    async def on_single_docx_click(self, e):
        files = await self.picker_single_docx.pick_files(allowed_extensions=["docx"])
        if files:
            self.selected_single_template = files[0].path
            self.txt_single_template.value = files[0].path
            self.page.update()

    async def on_image_click(self, e, tag_name: str, dialog_title: str):
        files = await self.image_pickers[tag_name].pick_files(dialog_title=dialog_title, file_type=ft.FilePickerFileType.IMAGE)
        if files and tag_name in self.image_fields:
            img_path = files[0].path
            fields = self.image_fields[tag_name]
            fields["path_field"].value = img_path
            fields["preview_img"].src = img_path
            fields["preview_img"].visible = True
            self.page.update()

    def get_current_form_context(self) -> Dict[str, str]:
        context = {}
        for var_name, tf in self.form_fields.items():
            context[var_name] = str(tf.value).strip() if tf.value else ""
        if "jenis_usaha" in context and ("jenis_usaha_bu" not in context or not context["jenis_usaha_bu"]):
            context["jenis_usaha_bu"] = context["jenis_usaha"]
        elif "jenis_usaha_bu" in context and ("jenis_usaha" not in context or not context["jenis_usaha"]):
            context["jenis_usaha"] = context["jenis_usaha_bu"]
        return context

    def _set_form_field_value(self, key: str, val: Any):
        target_key = key
        if target_key not in self.form_fields:
            if target_key == "jenis_usaha_bu" and "jenis_usaha" in self.form_fields:
                target_key = "jenis_usaha"
            elif target_key == "jenis_usaha" and "jenis_usaha_bu" in self.form_fields:
                target_key = "jenis_usaha_bu"
            else:
                return

        ctrl = self.form_fields[target_key]
        if isinstance(ctrl, ft.Dropdown):
            val_str = str(val).strip()
            matched = None
            for opt in ctrl.options:
                if opt.key == val_str:
                    matched = opt.key
                    break
            if not matched:
                val_lower = val_str.lower()
                if "konsultan" in val_lower:
                    matched = "Jasa Konsultansi Konstruksi"
                elif "pekerjaan" in val_lower or "kontruksi" in val_lower:
                    matched = "Jasa Pekerjaan Kontruksi"
                elif ctrl.options:
                    matched = ctrl.options[0].key
            ctrl.value = matched
        else:
            ctrl.value = val

    def get_current_image_map(self) -> Dict[str, Dict[str, Any]]:
        img_map = {}
        for tag_name, fields in self.image_fields.items():
            p = fields["path_field"].value
            if p and os.path.exists(p):
                h = int(fields["h_field"].value) if fields["h_field"].value.isdigit() else 25
                img_map[tag_name] = {"path": p, "height_mm": h}
        return img_map

    def reset_form_data(self):
        # 1. Purge all internal state
        self.form_fields.clear()
        self.serti_count = 1
        self.sbu_count = 1
        
        # 2. Reset image fields to default empty state
        from config import IMAGE_SLOTS
        for tag_name, slot in IMAGE_SLOTS.items():
            if tag_name in self.image_fields:
                fields = self.image_fields[tag_name]
                fields["path_field"].value = ""
                fields["h_field"].value = str(slot["default_height_mm"])
                fields["preview_img"].src = ""
                fields["preview_img"].visible = False
        
        # 3. Re-build the entire input data view from scratch!
        self.input_data_view = self.build_input_data_view()
        
        # 4. Mount the new UI back to the active tab (if it's the active view)
        if hasattr(self, 'smap_tab_view'):
            self.smap_tab_view.controls[0] = self.input_data_view
            if self.nav_rail.selected_index == 0 and self.smap_tabs.selected_index == 0:
                try:
                    self.smap_tab_view.update()
                except Exception:
                    pass
            
        self.show_snackbar("Form berhasil di-reset secara total", color="blue")
        if hasattr(self, 'page'):
            self.page.update()

    def on_generate_single_clicked(self, e):
        if not self.selected_single_template or not os.path.exists(self.selected_single_template):
            self.show_snackbar("Please select a valid Word (.docx) template file first!", color="red")
            return

        context = self.get_current_form_context()
        img_map = self.get_current_image_map()

        bu_name = context.get("bu_folder") or context.get("inisial_bu") or "Output"
        bu_folder = "".join(c for c in bu_name if c.isalnum() or c in (" ", "_", "-")).strip().replace(" ", "_")
        
        out_dir = os.path.join(self.batch_output_dir, bu_folder)
        os.makedirs(out_dir, exist_ok=True)
        base_name = os.path.splitext(os.path.basename(self.selected_single_template))[0]
        out_docx = os.path.join(out_dir, f"{base_name}_generated.docx")

        self.show_snackbar("Generating DOCX document...", color="blue")
        
        def run_proc():
            ok, res = DocumentEngine.process_document(self.selected_single_template, out_docx, context, img_map)
            if not ok:
                self.show_snackbar(f"Error: {res}", color="red")
                return

            msg = f"DOCX saved to: {out_docx}"

            if self.chk_single_pdf.value:
                out_pdf = os.path.splitext(out_docx)[0] + ".pdf"
                pdf_ok, pdf_res = PDFEngine.convert_docx_to_pdf(out_docx, out_pdf, self.libreoffice_path)
                if pdf_ok:
                    wm_text = self.txt_single_watermark.value.strip()
                    if wm_text:
                        PDFEngine.apply_watermark(out_pdf, out_pdf, watermark_text=wm_text)
                    msg += f"\nPDF saved to: {out_pdf}"
                else:
                    msg += f"\n(PDF Conversion Failed: {pdf_res})"

            self.show_snackbar("Document Generated Successfully!", color="green")
            # Open output directory
            try:
                if sys.platform == "win32":
                    os.startfile(out_dir)
                else:
                    subprocess.run(["open", out_dir])
            except Exception:
                pass

        threading.Thread(target=run_proc, daemon=True).start()

    # ==========================================
    # 2. BATCH MODE VIEW
    # ==========================================
    def build_batch_view(self) -> ft.Control:
        self.txt_batch_templates_count = ft.Text("No templates selected", weight=ft.FontWeight.BOLD)
        self.txt_batch_outdir = ft.TextField(label="Batch Output Folder", value=self.batch_output_dir, expand=True,)

        btn_select_batch_templates = ft.ElevatedButton(
            "Select Templates (Files)",
            icon=ft.Icons.FILE_OPEN,
            on_click=self.on_batch_templates_click
        )

        btn_select_batch_folder = ft.ElevatedButton(
            "Select Templates (Folder)",
            icon=ft.Icons.FOLDER_OPEN,
            on_click=self.on_batch_folder_click
        )

        btn_select_outdir = ft.IconButton(
            icon=ft.Icons.FOLDER,
            tooltip="Browse Output Directory",
            on_click=self.on_batch_outdir_click
        )

        self.radio_data_source = ft.RadioGroup(
            value="form",
            content=ft.Row([
                ft.Radio(value="form", label="Gunakan Data Input Manual Form (1 Record)"),
                ft.Radio(value="batch", label="Gunakan Data Import Excel (0 Records)")
            ])
        )

        config_card = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Row([btn_select_batch_templates, btn_select_batch_folder, self.txt_batch_templates_count], spacing=15),
                    ft.Row([self.txt_batch_outdir, btn_select_outdir]),
                    ft.Divider(height=1),
                    ft.Text("Sumber Data (Data Source):", weight=ft.FontWeight.BOLD),
                    self.radio_data_source
                ], spacing=15),
                padding=15
            )
        )

        # Batch Options & Console
        self.txt_batch_wm = ft.TextField(label="Batch PDF Watermark", value="", width=200,)

        self.batch_progress_bar = ft.ProgressBar(value=0.0, expand=True)
        self.txt_batch_status = ft.Text("Status: Idle", weight=ft.FontWeight.BOLD)
        self.batch_log_list = ft.ListView(expand=True, spacing=5, auto_scroll=True)

        btn_run_batch_word = ft.ElevatedButton(
            "Start Batch Replacement (Word)",
            style=ft.ButtonStyle(bgcolor=ft.Colors.BLUE_700, color=ft.Colors.WHITE),
            icon=ft.Icons.FILE_COPY,
            height=45,
            on_click=lambda e: self.on_run_batch_clicked(e, convert_to_pdf=False)
        )

        btn_run_batch_pdf = ft.ElevatedButton(
            "Start Batch Replacement (PDF)",
            style=ft.ButtonStyle(bgcolor=ft.Colors.RED_700, color=ft.Colors.WHITE),
            icon=ft.Icons.PICTURE_AS_PDF,
            height=45,
            on_click=lambda e: self.on_run_batch_clicked(e, convert_to_pdf=True)
        )

        btn_cancel_batch = ft.OutlinedButton(
            "Cancel",
            icon=ft.Icons.CANCEL,
            on_click=lambda _: self.batch_processor.cancel()
        )

        btn_open_outdir = ft.ElevatedButton(
            "Open Output Folder",
            icon=ft.Icons.FOLDER_SPECIAL,
            on_click=self.on_open_outdir_clicked
        )

        console_card = ft.Card(
            expand=True,
            content=ft.Container(
                expand=True,
                content=ft.Column([
                    ft.Row([self.txt_batch_wm, ft.Container(expand=True), btn_cancel_batch, btn_run_batch_word, btn_run_batch_pdf]),
                    ft.Row([self.batch_progress_bar]),
                    ft.Row([self.txt_batch_status, ft.Container(expand=True), btn_open_outdir]),
                    ft.Container(
                        content=self.batch_log_list,
                        border=ft.Border.all(1, ft.Colors.OUTLINE),
                        border_radius=5,
                        expand=True,
                        padding=10
                    )
                ], spacing=12, expand=True),
                padding=15
            )
        )

        return ft.Column([
            ft.Text("Document Processing Engine", size=20, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
            config_card,
            console_card
        ], expand=True, spacing=15)

    async def on_batch_templates_click(self, e):
        files = await self.picker_batch_templates.pick_files(
            dialog_title="Select Word Templates",
            allow_multiple=True, 
            allowed_extensions=["docx"]
        )
        if files:
            self.batch_templates = [f.path for f in files]
            msg = f"Terpilih {len(self.batch_templates)} template (.docx)"
            if hasattr(self, 'txt_batch_templates_count') and self.txt_batch_templates_count:
                self.txt_batch_templates_count.value = msg
            self.page.update()

    async def on_batch_folder_click(self, e):
        path = await self.picker_batch_folder.get_directory_path(
            dialog_title="Select Folder containing Word Templates"
        )
        if path:
            import glob
            # Filter out lock files starting with ~$
            files = [f for f in glob.glob(os.path.join(path, "*.docx")) if not os.path.basename(f).startswith("~$")]
            self.batch_templates = files
            msg = f"Folder: {len(self.batch_templates)} template (.docx) ditemukan"
            if hasattr(self, 'txt_batch_templates_count') and self.txt_batch_templates_count:
                self.txt_batch_templates_count.value = msg
            self.page.update()

    async def on_batch_outdir_click(self, e):
        path = await self.picker_batch_outdir.get_directory_path(dialog_title="Select Output Directory")
        if path:
            self.batch_output_dir = path
            self.txt_batch_outdir.value = path
            self.page.update()

    def on_run_batch_clicked(self, e, convert_to_pdf=False):
        if not self.batch_templates:
            self.show_snackbar("Silakan pilih file atau folder template Word (.docx) terlebih dahulu!", color="red")
            return

        out_dir = ""
        if hasattr(self, 'txt_batch_outdir') and self.txt_batch_outdir and self.txt_batch_outdir.value:
            out_dir = self.txt_batch_outdir.value.strip()
        if not out_dir:
            out_dir = getattr(self, 'batch_output_dir', os.path.abspath("output_documents"))
            if hasattr(self, 'txt_batch_outdir') and self.txt_batch_outdir:
                self.txt_batch_outdir.value = out_dir

        pairs = self.get_all_find_replace_pairs()
        init_cs = self.chk_batch_case_sensitive.value if hasattr(self, 'chk_batch_case_sensitive') else True

        # RadioGroup for user choice of Case Sensitive or Non-Case Sensitive
        case_radio = ft.RadioGroup(
            value="cs" if init_cs else "ci",
            content=ft.Column([
                ft.Radio(value="cs", label="Case Sensitive (Peka Huruf Besar & Kecil) - [Default]"),
                ft.Radio(value="ci", label="Non-Case Sensitive (Abaikan Huruf Besar & Kecil)")
            ], spacing=10)
        )

        mode_name = "Word (.docx)" if not convert_to_pdf else "PDF (.pdf)"
        dialog_items = [
            ft.Text(f"Target Format: {mode_name}", weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
            ft.Text(f"Jumlah File / Dokumen: {len(self.batch_templates)} template terpilih"),
            ft.Text(f"Folder Output: {out_dir}", size=12, color=ft.Colors.GREY_700),
            ft.Divider(height=1),
        ]

        if pairs:
            pairs_rows = []
            for p_idx, p in enumerate(pairs, 1):
                pairs_rows.append(
                    ft.Text(f"• [Aturan #{p_idx}] Cari: \"{p['find_text']}\" ➔ Ganti: \"{p['replace_text']}\"", weight=ft.FontWeight.BOLD)
                )

            dialog_items.extend([
                ft.Text(f"Penggantian Kata / Kalimat Aktif ({len(pairs)} aturan):", weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_800),
                ft.Container(
                    content=ft.Column(pairs_rows, spacing=5),
                    padding=10,
                    bgcolor=ft.Colors.BLUE_50,
                    border_radius=5
                ),
                ft.Text("Pilih Metode Sensitivitas Huruf (Case Sensitivity):", weight=ft.FontWeight.BOLD),
                case_radio
            ])
        else:
            dialog_items.extend([
                ft.Text("Pencarian kata khusus (Form 1) tidak diisi.", italic=True, color=ft.Colors.GREY_700),
                ft.Text("Proses batch akan mereplace semua variabel data template & image."),
                ft.Text("Opsi Sensitivitas Huruf (jika ada pencarian kata):", weight=ft.FontWeight.BOLD, size=12),
                case_radio
            ])

        def close_dlg(ev):
            batch_dialog.open = False
            try:
                self.page.close(batch_dialog)
            except Exception:
                pass
            self.page.update()

        def start_proc(ev):
            is_cs = (case_radio.value == "cs")
            if hasattr(self, 'chk_batch_case_sensitive'):
                self.chk_batch_case_sensitive.value = is_cs
            close_dlg(ev)
            self._start_batch_execution(convert_to_pdf=convert_to_pdf, case_sensitive=is_cs)

        batch_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Row([
                ft.Icon(ft.Icons.SETTINGS_SUGGEST, color=ft.Colors.BLUE_700),
                ft.Text("Konfirmasi Opsi Batch Replacement", weight=ft.FontWeight.BOLD)
            ]),
            content=ft.Container(
                content=ft.Column(dialog_items, spacing=10, tight=True),
                width=520
            ),
            actions=[
                ft.OutlinedButton("Batal", icon=ft.Icons.CANCEL, on_click=close_dlg),
                ft.ElevatedButton(
                    "Mulai Proses Batch ➔",
                    icon=ft.Icons.PLAY_ARROW,
                    style=ft.ButtonStyle(
                        bgcolor=ft.Colors.BLUE_700 if not convert_to_pdf else ft.Colors.RED_700,
                        color=ft.Colors.WHITE
                    ),
                    on_click=start_proc
                )
            ],
            actions_alignment=ft.MainAxisAlignment.END
        )

        self.page.overlay.append(batch_dialog)
        batch_dialog.open = True
        self.page.update()

    def _start_batch_execution(self, convert_to_pdf=False, case_sensitive=True):
        # Automatically switch to Process Documents tab to display progress & logs live
        self.nav_rail.selected_index = 0
        self.smap_tabs.selected_index = 1
        self.content_area.content = self.smap_automation_view

        if hasattr(self, 'radio_data_source') and self.radio_data_source.value == "batch" and self.batch_data_records:
            records = self.batch_data_records
        else:
            records = [self.get_current_form_context()]
            
        img_map = self.get_current_image_map()

        pairs = self.get_all_find_replace_pairs()
        custom_find_replace = None
        if pairs:
            custom_find_replace = {
                "find_text": pairs[0]["find_text"],
                "replace_text": pairs[0]["replace_text"],
                "case_sensitive": case_sensitive,
                "pairs": [
                    {
                        "find_text": p["find_text"],
                        "replace_text": p["replace_text"],
                        "case_sensitive": case_sensitive
                    }
                    for p in pairs
                ]
            }

        self.batch_log_list.controls.clear()
        self.batch_progress_bar.value = 0.0
        self.txt_batch_status.value = "Status: Running..."
        self.page.update()

        self.batch_is_running = True
        async def watchdog(*args):
            import asyncio
            while getattr(self, 'batch_is_running', False):
                try:
                    self.page.update()
                except Exception:
                    pass
                await asyncio.sleep(1)
        self.page.run_task(watchdog)

        def log_cb(msg):
            self.batch_log_list.controls.append(ft.Text(msg, size=12))
            self.batch_log_list.update()

        def progress_cb(current, total, filename, status):
            self.batch_progress_bar.value = current / total if total > 0 else 0
            self.txt_batch_status.value = f"Progress: {current}/{total} ({int(self.batch_progress_bar.value*100)}%) - {filename}"
            self.batch_progress_bar.update()
            self.txt_batch_status.update()

        def complete_cb(summary):
            self.batch_is_running = False
            self.txt_batch_status.value = f"Status: Finished in {summary['elapsed_seconds']}s ({summary['successful']} succeeded, {summary['errors']} failed)"
            self.txt_batch_status.update()
            self.show_snackbar("Batch Execution Finished!", color="green")

        def run_thread():
            out_target = self.txt_batch_outdir.value if (hasattr(self, 'txt_batch_outdir') and self.txt_batch_outdir and self.txt_batch_outdir.value) else getattr(self, 'batch_output_dir', os.path.abspath("output_documents"))
            self.batch_processor.process_batch(
                template_paths=self.batch_templates,
                data_records=records,
                output_dir=out_target,
                convert_to_pdf=convert_to_pdf,
                watermark_text=self.txt_batch_wm.value,
                image_map=img_map,
                libreoffice_path=self.libreoffice_path,
                custom_find_replace=custom_find_replace,
                on_progress=progress_cb,
                on_log=log_cb,
                on_complete=complete_cb
            )

        threading.Thread(target=run_thread, daemon=True).start()

    def on_open_outdir_clicked(self, e):
        d = self.txt_batch_outdir.value
        if os.path.exists(d):
            try:
                if sys.platform == "win32":
                    os.startfile(d)
                else:
                    subprocess.run(["open", d])
            except Exception as ex:
                self.show_snackbar(f"Failed to open folder: {ex}", color="red")

    # ==========================================
    # 3. PRESETS & PROFILES VIEW
    # ==========================================
    def build_presets_view(self) -> ft.Control:
        btn_load_sample = ft.ElevatedButton(
            "Load Standard SMAP Preset",
            icon=ft.Icons.SETTINGS_BACKUP_RESTORE,
            on_click=lambda _: self.load_sample_data()
        )

        btn_save_json = ft.ElevatedButton(
            "Export Profile to JSON",
            icon=ft.Icons.SAVE_ALT,
            on_click=self.on_json_save_click
        )

        btn_load_json = ft.ElevatedButton(
            "Import Profile from JSON",
            icon=ft.Icons.FOLDER_OPEN,
            on_click=self.on_json_load_click
        )
        
        btn_save_history = ft.ElevatedButton(
            "Save Current Form to Database History",
            icon=ft.Icons.SAVE,
            on_click=self.on_save_history_click,
            style=ft.ButtonStyle(color=ft.Colors.WHITE, bgcolor=ft.Colors.BLUE_700)
        )

        btn_export_excel = ft.ElevatedButton(
            "Export Single/Batch Data to Excel (.xlsx)",
            icon=ft.Icons.GRID_ON,
            on_click=self.on_excel_export_click
        )

        btn_import_excel = ft.ElevatedButton(
            "Import Single/Batch Data from Excel (.xlsx)",
            icon=ft.Icons.FILE_DOWNLOAD,
            on_click=self.on_excel_import_click
        )

        return ft.Column([
            ft.Text("Data Profiles, Presets & File Exchange", size=20, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
            ft.Card(
                content=ft.Container(
                    content=ft.Column([
                        ft.Text("Database History Actions", weight=ft.FontWeight.BOLD, size=16),
                        ft.Row([btn_save_history]),
                        ft.Divider(),
                        ft.Text("Quick Profile Actions", weight=ft.FontWeight.BOLD, size=16),
                        ft.Row([btn_load_sample, btn_save_json, btn_load_json]),
                        ft.Divider(),
                        ft.Text("Excel Batch Data Exchange", weight=ft.FontWeight.BOLD, size=16),
                        ft.Row([btn_export_excel, btn_import_excel])
                    ], spacing=15),
                    padding=20
                )
            )
        ], spacing=20)

    def _on_focus_select_all(self, e):
        val_len = len(e.control.value or "")
        e.control.selection = ft.TextSelection(0, val_len)
        e.control.update()

    def on_save_history_click(self, e):
        # Determine the name of the profile based on the 'bu' field
        nama_bu = self.form_fields.get('bu').value if 'bu' in self.form_fields else "Unknown BU"
        if not nama_bu or not nama_bu.strip():
            nama_bu = "Unnamed Profile"
            
        data = {
            "variables": self.get_current_form_context(),
            "images": {k: {"path": v["path_field"].value, "height_mm": int(v["h_field"].value or 25)} for k, v in self.image_fields.items()}
        }
        
        profile_id = self.db.save_profile(nama_bu, data)
        if profile_id:
            self.show_snackbar(f"Saved to Database History as '{nama_bu}'", color="green")
            # Also refresh the history list
            self.refresh_history_list(None)
        else:
            self.show_snackbar("Failed to save to Database History.", color="red")

    def load_sample_data(self, show_snack: bool = True):
        sample = ProfileManager.get_sample_smap_profile()
        self.reset_form_data()
        for k, val in sample["variables"].items():
            self.ensure_dynamic_field(k)
            self._set_form_field_value(k, val)

        for k, val in sample["images"].items():
            if k in self.image_fields:
                self.image_fields[k]["path_field"].value = val["path"]
                self.image_fields[k]["h_field"].value = str(val.get("height_mm", 25))
                try:
                    self.image_fields[k]["path_field"].update()
                    self.image_fields[k]["h_field"].update()
                except Exception:
                    pass

        if show_snack:
            self.show_snackbar("Loaded Standard SMAP Corporate Preset!", color="green")
        self.page.update()

    async def on_json_save_click(self, e):
        path = await self.picker_json_save.save_file(dialog_title="Save Profile as JSON", file_type=ft.FilePickerFileType.CUSTOM, allowed_extensions=["json"])
        if path:
            if not path.lower().endswith(".json"):
                path += ".json"
            data = {
                "variables": self.get_current_form_context(),
                "images": {k: {"path": v["path_field"].value, "height_mm": int(v["h_field"].value or 25)} for k, v in self.image_fields.items()}
            }
            ok, msg = ProfileManager.save_profile_json(data, path)
            if ok:
                self.show_snackbar(f"Saved profile to {path}", color="green")
            else:
                self.show_snackbar(f"Error: {msg}", color="red")

    async def on_json_load_click(self, e):
        files = await self.picker_json_load.pick_files(dialog_title="Load Profile from JSON", file_type=ft.FilePickerFileType.CUSTOM, allowed_extensions=["json"])
        if files:
            ok, profile, msg = ProfileManager.load_profile_json(files[0].path)
            if ok and profile:
                self.reset_form_data()
                for k, val in profile["variables"].items():
                    self.ensure_dynamic_field(k)
                    self._set_form_field_value(k, val)
                self.show_snackbar("Loaded JSON profile!", color="green")
                self.page.update()
            else:
                self.show_snackbar(f"Error: {msg}", color="red")

    async def on_excel_export_click(self, e):
        path = await self.picker_excel_export.save_file(dialog_title="Export Data as Excel", file_name="Template_Batch_Data.xlsx", file_type=ft.FilePickerFileType.CUSTOM, allowed_extensions=["xlsx"])
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            records = [self.get_current_form_context()]
            ok, msg = ProfileManager.export_profile_excel(records, path)
            if ok:
                self.show_snackbar(f"Exported Excel to {path}", color="green")
            else:
                self.show_snackbar(f"Error: {msg}", color="red")

    async def on_excel_import_click(self, e):
        files = await self.picker_excel_import.pick_files(dialog_title="Import Data from Excel", file_type=ft.FilePickerFileType.CUSTOM, allowed_extensions=["xlsx"])
        if files:
            ok, records, msg = ProfileManager.import_profile_excel(files[0].path)
            if ok:
                self.batch_data_records = records
                if hasattr(self, 'radio_data_source'):
                    self.radio_data_source.value = "batch"
                    self.radio_data_source.content.controls[1].label = f"Gunakan Data Import Excel ({len(records)} Records)"
                    try:
                        self.radio_data_source.update()
                    except Exception:
                        pass
                if records:
                    self.reset_form_data()
                    # Populate form with first record
                    for k, val in records[0].items():
                        self.ensure_dynamic_field(k)
                        self._set_form_field_value(k, val)
                self.show_snackbar(msg, color="green")
                self.page.update()
            else:
                self.show_snackbar(f"Error: {msg}", color="red")

    # ==========================================
    # 4. DATABASE HISTORY VIEW
    # ==========================================
    def build_history_view(self) -> ft.Control:
        self.txt_history_search = ft.TextField(label="Search by Nama Badan Usaha (BU) or Date", expand=True)
        self.history_list_view = ft.ListView(expand=True, spacing=10, padding=10)

        btn_search = ft.ElevatedButton("Search", icon=ft.Icons.SEARCH, on_click=self.on_history_search_click)
        btn_refresh = ft.OutlinedButton("Refresh", icon=ft.Icons.REFRESH, on_click=self.refresh_history_list)

        search_row = ft.Row([self.txt_history_search, btn_search, btn_refresh])

        # Initial Load
        self.refresh_history_list(None)

        return ft.Column([
            ft.Text("Database History", size=20, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
            ft.Card(
                content=ft.Container(
                    content=ft.Column([
                        search_row,
                        ft.Divider(),
                        ft.Container(
                            content=self.history_list_view,
                            border=ft.Border.all(1, ft.Colors.OUTLINE),
                            border_radius=5,
                            expand=True,
                            height=600
                        )
                    ]),
                    padding=20,
                    expand=True
                ),
                expand=True
            )
        ], expand=True, spacing=20)

    def refresh_history_list(self, e):
        profiles = self.db.get_all_profiles()
        self._populate_history_list(profiles)

    def on_history_search_click(self, e):
        keyword = self.txt_history_search.value.strip()
        if keyword:
            profiles = self.db.search_profiles(keyword)
        else:
            profiles = self.db.get_all_profiles()
        self._populate_history_list(profiles)

    def _populate_history_list(self, profiles: List[Dict[str, Any]]):
        self.history_list_view.controls.clear()
        
        if not profiles:
            self.history_list_view.controls.append(ft.Text("No historical data found.", italic=True))
            
        for prof in profiles:
            btn_view = ft.ElevatedButton("View", icon=ft.Icons.VISIBILITY, on_click=lambda e, p=prof: self.on_history_view_click(p))
            btn_load = ft.ElevatedButton("Load", icon=ft.Icons.DOWNLOAD, on_click=lambda e, p=prof: self.on_history_load_click(p))
            btn_delete = ft.IconButton(icon=ft.Icons.DELETE, icon_color=ft.Colors.RED_700, tooltip="Delete Profile", on_click=lambda e, pid=prof["id"]: self.on_history_delete_click(pid))
            
            row = ft.Row([
                ft.Column([
                    ft.Text(prof["nama_bu"], weight=ft.FontWeight.BOLD, size=16),
                    ft.Text(f"Saved: {prof['timestamp']}", size=12, color=ft.Colors.GREY_700)
                ], expand=True),
                ft.Row([btn_view, btn_load, btn_delete], spacing=5)
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
            
            self.history_list_view.controls.append(
                ft.Container(
                    content=row,
                    padding=10,
                    border=ft.Border.all(1, ft.Colors.GREY_300),
                    border_radius=5,
                    bgcolor=ft.Colors.GREY_50
                )
            )
        if hasattr(self, 'page') and self.page:
            self.page.update()

    def on_history_view_click(self, profile: Dict[str, Any]):
        data = profile["data_json"]
        
        view_content = ft.Column(scroll=ft.ScrollMode.AUTO, expand=True, spacing=15)
        
        if "variables" in data and data["variables"]:
            view_content.controls.append(ft.Text("Variables:", weight=ft.FontWeight.BOLD, size=16))
            var_grid = ft.Column(spacing=5)
            for k, v in data["variables"].items():
                if v: 
                    var_grid.controls.append(
                        ft.Row([
                            ft.Text(f"{k}:", weight=ft.FontWeight.W_500, width=150, color=ft.Colors.BLUE_900),
                            ft.Text(str(v), expand=True)
                        ])
                    )
            view_content.controls.append(
                ft.Container(content=var_grid, padding=10, border=ft.Border.all(1, ft.Colors.BLUE_100), border_radius=5, bgcolor=ft.Colors.BLUE_50)
            )
            
        if "images" in data and data["images"]:
            view_content.controls.append(ft.Text("Image Assets:", weight=ft.FontWeight.BOLD, size=16))
            img_grid = ft.Column(spacing=5)
            for k, v in data["images"].items():
                if v.get("path"):
                    img_grid.controls.append(
                        ft.Row([
                            ft.Text(f"{k}:", weight=ft.FontWeight.W_500, width=150, color=ft.Colors.GREEN_900),
                            ft.Text(f"{v['path']} (Height: {v.get('height_mm', 25)}mm)", expand=True)
                        ])
                    )
            if img_grid.controls:
                view_content.controls.append(
                    ft.Container(content=img_grid, padding=10, border=ft.Border.all(1, ft.Colors.GREEN_100), border_radius=5, bgcolor=ft.Colors.GREEN_50)
                )

        def close_dialog(e):
            dialog.open = False
            self.page.update()

        dialog = ft.AlertDialog(
            title=ft.Text(f"Profile: {profile['nama_bu']}"),
            content=ft.Container(content=view_content, width=700, height=500),
            actions=[ft.TextButton("Close", on_click=close_dialog)],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        
        self.page.overlay.append(dialog)
        dialog.open = True
        self.page.update()
        if hasattr(self, 'page') and self.page:
            self.page.update()

    def on_history_load_click(self, profile: Dict[str, Any]):
        data = profile["data_json"]
        self.reset_form_data()
        
        if "variables" in data:
            for k, val in data["variables"].items():
                self.ensure_dynamic_field(k)
                self._set_form_field_value(k, val)
                    
        if "images" in data:
            for k, val in data["images"].items():
                if k in self.image_fields:
                    self.image_fields[k]["path_field"].value = val["path"]
                    self.image_fields[k]["h_field"].value = str(val.get("height_mm", 25))

        self.show_snackbar(f"Loaded history profile: {profile['nama_bu']}", color="green")
        
        # Navigate to Input Data Tab
        self.nav_rail.selected_index = 0
        self.content_area.content = self.input_data_view
        self.page.update()

    def on_history_delete_click(self, profile_id: int):
        ok = self.db.delete_profile(profile_id)
        if ok:
            self.show_snackbar("Profile deleted from history.", color="blue")
            self.refresh_history_list(None)
        else:
            self.show_snackbar("Failed to delete profile.", color="red")

    # ==========================================
    # 5. PDF TOOLKIT VIEW
    # ==========================================
    def build_pdf_toolkit_view(self) -> ft.Control:
        # Converter Section
        self.txt_tk_docx = ft.TextField(label="DOCX File to Convert", expand=True,)
        btn_convert = ft.ElevatedButton("Convert to PDF", icon=ft.Icons.PICTURE_AS_PDF, on_click=self.on_tk_convert_clicked)

        card_converter = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("1. DOCX to PDF Converter", weight=ft.FontWeight.BOLD),
                    ft.Row([
                        self.txt_tk_docx,
                        ft.IconButton(ft.Icons.FOLDER, on_click=self.on_toolkit_docx_click),
                        btn_convert
                    ])
                ]),
                padding=15
            )
        )

        # Watermark Section
        self.txt_tk_wm_pdf = ft.TextField(label="PDF File to Watermark", expand=True,)
        self.txt_tk_wm_text = ft.TextField(label="Watermark Text", value="SALINAN RESMI", width=200,)
        btn_wm = ft.ElevatedButton("Apply Watermark", icon=ft.Icons.WATER_DROP, on_click=self.on_tk_watermark_clicked)

        card_wm = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("2. PDF Watermarking Utility", weight=ft.FontWeight.BOLD),
                    ft.Row([
                        self.txt_tk_wm_pdf,
                        ft.IconButton(ft.Icons.FOLDER, on_click=self.on_toolkit_watermark_pdf_click),
                        self.txt_tk_wm_text,
                        btn_wm
                    ])
                ]),
                padding=15
            )
        )

        # Merge Section
        self.col_tk_merge_list = ft.Column(spacing=5, height=150, scroll=ft.ScrollMode.AUTO)
        self.selected_merge_pdfs: List[str] = []
        self._render_merge_list()
        
        self.txt_tk_merge_outdir = ft.TextField(label="Output Folder (Optional)", expand=True, read_only=True)
        self.txt_tk_merge_filename = ft.TextField(label="Output Filename", value="merged_consolidated.pdf", width=200)
        btn_merge = ft.ElevatedButton("Merge PDFs", icon=ft.Icons.MERGE_TYPE, on_click=self.on_tk_merge_clicked)

        card_merge = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("3. PDF Batch Merge Utility", weight=ft.FontWeight.BOLD),
                    ft.Row([
                        ft.ElevatedButton("Select PDFs", icon=ft.Icons.PICTURE_AS_PDF, on_click=self.on_toolkit_merge_pdfs_click),
                        ft.ElevatedButton("Add Folder", icon=ft.Icons.FOLDER_OPEN, on_click=self.on_toolkit_merge_indir_click),
                        ft.ElevatedButton("Clear List", icon=ft.Icons.CLEAR_ALL, on_click=self.on_tk_merge_clear_list),
                    ]),
                    ft.Text("You can append more PDFs and reorder them below:", italic=True, size=12),
                    ft.Container(
                        content=self.col_tk_merge_list,
                        border_radius=5,
                        padding=10,
                    ),
                    ft.Row([
                        self.txt_tk_merge_outdir,
                        ft.IconButton(ft.Icons.FOLDER, on_click=self.on_toolkit_merge_outdir_click),
                        self.txt_tk_merge_filename,
                        btn_merge
                    ])
                ]),
                padding=15
            )
        )

        # Split Section
        self.txt_tk_split_pdf = ft.TextField(label="Select PDF to Split", expand=True, read_only=True)
        self.txt_tk_split_outdir = ft.TextField(label="Output Folder", expand=True, read_only=True)
        def on_split_mode_change(e):
            self.txt_tk_split_pages.disabled = (self.radio_split_mode.value == "all")
            self.page.update()
            
        self.radio_split_mode = ft.RadioGroup(content=ft.Row([
            ft.Radio(value="all", label="Split All Pages"),
            ft.Radio(value="selected", label="Split Selected Pages")
        ]), value="all", on_change=on_split_mode_change)
        
        self.txt_tk_split_pages = ft.TextField(label="Pages (e.g. 1, 3, 5-7)", disabled=True)
        btn_split = ft.ElevatedButton("Split PDF", icon=ft.Icons.CALL_SPLIT, on_click=self.on_tk_split_clicked)
        
        card_split = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("4. PDF Split Utility", weight=ft.FontWeight.BOLD),
                    ft.Row([
                        self.txt_tk_split_pdf,
                        ft.IconButton(ft.Icons.FILE_OPEN, on_click=self.on_toolkit_split_pdf_click)
                    ]),
                    self.radio_split_mode,
                    self.txt_tk_split_pages,
                    ft.Row([
                        self.txt_tk_split_outdir,
                        ft.IconButton(ft.Icons.FOLDER, on_click=self.on_toolkit_split_outdir_click),
                        btn_split
                    ])
                ]),
                padding=15
            )
        )

        return ft.Column([
            ft.Container(
                content=ft.Column([
                    ft.Text("Advanced PDF Publishing Toolkit", size=20, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
                    card_converter,
                    card_wm,
                    card_merge,
                    card_split
                ], spacing=15),
                padding=ft.Padding.only(right=15)
            )
        ], scroll=ft.ScrollMode.AUTO, expand=True)

    async def on_toolkit_docx_click(self, e):
        files = await self.picker_toolkit_docx.pick_files(allowed_extensions=["docx"])
        if files:
            self.txt_tk_docx.value = files[0].path
            self.page.update()

    def on_tk_convert_clicked(self, e):
        f = self.txt_tk_docx.value
        if f and os.path.exists(f):
            out_pdf = os.path.splitext(f)[0] + ".pdf"
            ok, msg = PDFEngine.convert_docx_to_pdf(f, out_pdf, self.libreoffice_path)
            if ok:
                self.show_snackbar(f"Converted to PDF: {out_pdf}", color="green")
            else:
                self.show_snackbar(f"Conversion failed: {msg}", color="red")

    async def on_toolkit_watermark_pdf_click(self, e):
        files = await self.picker_toolkit_watermark_pdf.pick_files(allowed_extensions=["pdf"])
        if files:
            self.txt_tk_wm_pdf.value = files[0].path
            self.page.update()

    def on_tk_watermark_clicked(self, e):
        f = self.txt_tk_wm_pdf.value
        wm = self.txt_tk_wm_text.value.strip()
        if f and os.path.exists(f) and wm:
            out_pdf = os.path.splitext(f)[0] + "_watermarked.pdf"
            ok, msg = PDFEngine.apply_watermark(f, out_pdf, watermark_text=wm)
            if ok:
                self.show_snackbar(f"Watermarked PDF saved: {out_pdf}", color="green")
            else:
                self.show_snackbar(f"Watermark failed: {msg}", color="red")

    async def on_toolkit_stamp_pdf_click(self, e):
        pass

    async def on_toolkit_stamp_img_click(self, e):
        pass

    async def on_toolkit_merge_outdir_click(self, e):
        path = await self.picker_toolkit_merge_outdir.get_directory_path(dialog_title="Select Output Directory")
        if path:
            self.txt_tk_merge_outdir.value = path
            self.page.update()

    async def on_toolkit_merge_indir_click(self, e):
        path = await self.picker_toolkit_merge_indir.get_directory_path(dialog_title="Select Folder with PDFs")
        if path:
            import glob
            pdfs = glob.glob(os.path.join(path, "*.pdf"))
            for p in sorted(pdfs):
                if p not in self.selected_merge_pdfs:
                    self.selected_merge_pdfs.append(p)
            self._render_merge_list()
            self.page.update()
            
    def on_tk_merge_clear_list(self, e):
        self.selected_merge_pdfs.clear()
        self._render_merge_list()
        self.page.update()

    async def on_toolkit_merge_pdfs_click(self, e):
        files = await self.picker_toolkit_merge_pdfs.pick_files(allowed_extensions=["pdf"], allow_multiple=True)
        if files:
            for f in files:
                if f.path not in self.selected_merge_pdfs:
                    self.selected_merge_pdfs.append(f.path)
            self._render_merge_list()
            self.page.update()

    def on_tk_merge_clicked(self, e):
        if len(self.selected_merge_pdfs) < 2:
            self.show_snackbar("Please select at least 2 PDF files to merge!", color="red")
            return
            
        out_dir = self.txt_tk_merge_outdir.value
        if not out_dir:
            out_dir = os.path.dirname(self.selected_merge_pdfs[0])
            
        out_filename = self.txt_tk_merge_filename.value.strip()
        if not out_filename:
            out_filename = "merged_consolidated.pdf"
        if not out_filename.lower().endswith(".pdf"):
            out_filename += ".pdf"
            
        out_pdf = os.path.join(out_dir, out_filename)
        ok, msg = PDFEngine.merge_pdfs(self.selected_merge_pdfs, out_pdf)
        if ok:
            def open_folder(e):
                import platform, subprocess
                if platform.system() == "Windows":
                    os.startfile(out_dir)
                elif platform.system() == "Darwin":
                    subprocess.Popen(["open", out_dir])
                else:
                    subprocess.Popen(["xdg-open", out_dir])
                dlg.open = False
                self.page.update()
            
            def close_dlg(e):
                dlg.open = False
                self.page.update()

            dlg = ft.AlertDialog(
                title=ft.Text("Merge Successful!"),
                content=ft.Text(f"Merged PDF saved to:\n{out_pdf}"),
                actions=[
                    ft.TextButton("Open Folder", on_click=open_folder),
                    ft.TextButton("OK", on_click=close_dlg)
                ],
                actions_alignment=ft.MainAxisAlignment.END,
            )
            self.page.dialog = dlg
            dlg.open = True
            self.page.update()
        else:
            self.show_snackbar(f"Merge failed: {msg}", color="red")

    def _render_merge_list(self):
        self.col_tk_merge_list.controls.clear()
        if not self.selected_merge_pdfs:
            self.col_tk_merge_list.controls.append(ft.Text("No PDFs selected.", italic=True, color=ft.Colors.GREY_500))
        else:
            for i, pdf_path in enumerate(self.selected_merge_pdfs):
                filename = os.path.basename(pdf_path)
                
                def create_callbacks(idx):
                    def move_up(e):
                        if idx > 0:
                            self.selected_merge_pdfs[idx - 1], self.selected_merge_pdfs[idx] = self.selected_merge_pdfs[idx], self.selected_merge_pdfs[idx - 1]
                            self._render_merge_list()
                            self.page.update()
                            
                    def move_down(e):
                        if idx < len(self.selected_merge_pdfs) - 1:
                            self.selected_merge_pdfs[idx + 1], self.selected_merge_pdfs[idx] = self.selected_merge_pdfs[idx], self.selected_merge_pdfs[idx + 1]
                            self._render_merge_list()
                            self.page.update()
                            
                    def remove_item(e):
                        self.selected_merge_pdfs.pop(idx)
                        self._render_merge_list()
                        self.page.update()
                    return move_up, move_down, remove_item
                
                m_up, m_down, r_item = create_callbacks(i)
                
                info = PDFEngine.get_pdf_info(pdf_path)
                info_text = f"Size: {info['size_kb']} KB | Pages: {info['pages']} | Modified: {info['modified_date']}"

                row = ft.Row([
                    ft.Text(f"{i+1}.", width=20, weight=ft.FontWeight.BOLD),
                    ft.Column([
                        ft.Text(filename, tooltip=pdf_path, weight=ft.FontWeight.BOLD),
                        ft.Text(info_text, size=11, color=ft.Colors.GREY_600)
                    ], expand=True, spacing=1),
                    ft.IconButton(ft.Icons.ARROW_UPWARD, on_click=m_up, disabled=(i == 0), icon_size=18, tooltip="Move Up", icon_color=ft.Colors.BLUE_500),
                    ft.IconButton(ft.Icons.ARROW_DOWNWARD, on_click=m_down, disabled=(i == len(self.selected_merge_pdfs) - 1), icon_size=18, tooltip="Move Down", icon_color=ft.Colors.BLUE_500),
                    ft.IconButton(ft.Icons.DELETE, on_click=r_item, icon_color=ft.Colors.RED_500, icon_size=18, tooltip="Remove")
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
                
                self.col_tk_merge_list.controls.append(row)

    async def on_toolkit_split_pdf_click(self, e):
        files = await self.picker_toolkit_split_pdf.pick_files(allowed_extensions=["pdf"])
        if files:
            self.txt_tk_split_pdf.value = files[0].path
            if not self.txt_tk_split_outdir.value:
                self.txt_tk_split_outdir.value = os.path.dirname(files[0].path)
            self.page.update()
            
    async def on_toolkit_split_outdir_click(self, e):
        path = await self.picker_toolkit_split_outdir.get_directory_path(dialog_title="Select Output Directory")
        if path:
            self.txt_tk_split_outdir.value = path
            self.page.update()
            
    def on_tk_split_clicked(self, e):
        pdf_path = self.txt_tk_split_pdf.value
        out_dir = self.txt_tk_split_outdir.value
        mode = self.radio_split_mode.value
        pages_str = self.txt_tk_split_pages.value.strip()
        
        if not pdf_path or not os.path.exists(pdf_path):
            self.show_snackbar("Please select a valid PDF to split.", color="red")
            return
            
        selected_pages = []
        if mode == "selected":
            if not pages_str:
                self.show_snackbar("Please enter pages to split (e.g. 1, 3, 5-7).", color="red")
                return
            try:
                # Parse pages string like "1, 3, 5-7"
                parts = [p.strip() for p in pages_str.split(',')]
                for part in parts:
                    if '-' in part:
                        start, end = map(int, part.split('-'))
                        selected_pages.extend(range(start - 1, end))
                    else:
                        selected_pages.append(int(part) - 1)
            except Exception as parse_err:
                self.show_snackbar(f"Invalid page format: {parse_err}", color="red")
                return
                
        ok, msg = PDFEngine.split_pdf(pdf_path, out_dir, mode, selected_pages)
        if ok:
            self.show_snackbar(msg, color="green")
        else:
            self.show_snackbar(f"Split failed: {msg}", color="red")

    # ==========================================
    # 5. IMAGE TOOLKIT VIEW
    # ==========================================
    def build_image_toolkit_view(self) -> ft.Control:
        self.txt_bg_in = ft.TextField(label="Select Image (Pas Foto, Tanda Tangan, Stempel)", expand=True, read_only=True)
        self.txt_bg_outdir = ft.TextField(label="Output Folder", expand=True, read_only=True)
        
        btn_bg_remove = ft.ElevatedButton("Remove Background", icon=ft.Icons.AUTO_FIX_HIGH, on_click=self.on_bg_remove_clicked, style=ft.ButtonStyle(bgcolor=ft.Colors.BLUE_700, color=ft.Colors.WHITE))
        self.btn_bg_save = ft.ElevatedButton("Save Result 300 DPI", icon=ft.Icons.SAVE, on_click=self.on_bg_save_clicked, style=ft.ButtonStyle(bgcolor=ft.Colors.GREEN_700, color=ft.Colors.WHITE), disabled=True)
        
        self.txt_bg_status = ft.Text("Ready", italic=True, size=12, color=ft.Colors.GREY_700)
        
        self.img_bg_before = ft.Image(src="", width=250, height=250, fit="contain", visible=False)
        self.img_bg_after = ft.Image(src="", width=250, height=250, fit="contain", visible=False)
        
        preview_row = ft.Row([
            ft.Column([ft.Text("Original", weight=ft.FontWeight.BOLD), self.img_bg_before], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            ft.Column([ft.Text("Result", weight=ft.FontWeight.BOLD), self.img_bg_after], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER)
        ], alignment=ft.MainAxisAlignment.CENTER, visible=False)
        self.bg_preview_row = preview_row
        
        self.processed_bg_base64 = None
        
        card_bg = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("AI Background Remover", weight=ft.FontWeight.BOLD),
                    ft.Text("Use this tool to remove the background of passport photos, signatures, or stamps. The result is saved as a high-quality 300 DPI PNG.", size=12, color=ft.Colors.GREY_600),
                    ft.Row([
                        self.txt_bg_in,
                        ft.IconButton(ft.Icons.IMAGE_SEARCH, on_click=self.on_bg_remover_in_click)
                    ]),
                    ft.Row([
                        self.txt_bg_outdir,
                        ft.IconButton(ft.Icons.FOLDER, on_click=self.on_bg_remover_outdir_click)
                    ]),
                    ft.Row([btn_bg_remove, self.btn_bg_save]),
                    self.txt_bg_status,
                    self.bg_preview_row
                ], spacing=15),
                padding=15
            )
        )
        
        return ft.Column([
            ft.Container(
                content=ft.Column([
                    ft.Text("Image Toolkit", size=20, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
                    card_bg
                ], spacing=15),
                padding=ft.Padding.only(right=15)
            )
        ], scroll=ft.ScrollMode.AUTO, expand=True)

    async def on_bg_remover_in_click(self, e):
        files = await self.picker_bg_remover_in.pick_files(allowed_extensions=["png", "jpg", "jpeg", "webp"])
        if files:
            self.txt_bg_in.value = files[0].path
            if not self.txt_bg_outdir.value:
                self.txt_bg_outdir.value = os.path.dirname(files[0].path)
                
            self.img_bg_before.src = files[0].path
            self.img_bg_before.visible = True
            self.img_bg_after.visible = False
            self.bg_preview_row.visible = True
            
            self.processed_bg_base64 = None
            self.btn_bg_save.disabled = True
            
            self.page.update()
            
    async def on_bg_remover_outdir_click(self, e):
        path = await self.picker_bg_remover_outdir.get_directory_path(dialog_title="Select Output Directory")
        if path:
            self.txt_bg_outdir.value = path
            self.page.update()
            
    async def on_bg_remove_clicked(self, e):
        in_path = self.txt_bg_in.value
        
        if not in_path or not os.path.exists(in_path):
            self.show_snackbar("Please select a valid input image.", color="red")
            return
            
        self.txt_bg_status.value = "Processing image using AI... This may take a moment."
        self.txt_bg_status.color = ft.Colors.BLUE_700
        self.btn_bg_save.disabled = True
        self.page.update()
        
        import asyncio
        import tempfile
        import time
        import base64
        
        api_key = self.removebg_api_key
        if api_key and len(api_key) > 5:
            ok, msg = await asyncio.to_thread(ImageEngine.process_background_removebg_api, in_path, api_key)
        else:
            ok, msg = await asyncio.to_thread(ImageEngine.process_background_base64, in_path)
            
        if ok:
            self.processed_bg_base64 = msg
            
            # Save base64 to a temporary file for preview to ensure compatibility
            temp_dir = tempfile.gettempdir()
            temp_preview = os.path.join(temp_dir, f"smap_bg_preview_{int(time.time())}.png")
            with open(temp_preview, "wb") as fh:
                fh.write(base64.b64decode(msg))
                
            self.img_bg_after.src = temp_preview
            self.img_bg_after.visible = True
            self.btn_bg_save.disabled = False
            
            self.txt_bg_status.value = "Background removed! Preview is ready. Click Save to export."
            self.txt_bg_status.color = ft.Colors.GREEN_700
            self.show_snackbar("Background removed successfully! Please save the result.", color="green")
        else:
            self.txt_bg_status.value = f"Failed: {msg}"
            self.txt_bg_status.color = ft.Colors.RED_700
            self.show_snackbar("Failed to remove background.", color="red")
            
        self.page.update()

    def on_bg_save_clicked(self, e):
        if not getattr(self, 'processed_bg_base64', None):
            self.show_snackbar("No processed image available. Please remove background first.", color="red")
            return
            
        out_dir = self.txt_bg_outdir.value
        in_path = self.txt_bg_in.value
        if not out_dir or not os.path.exists(out_dir):
            self.show_snackbar("Please select a valid output directory.", color="red")
            return
            
        base_name = os.path.splitext(os.path.basename(in_path))[0]
        out_path = os.path.join(out_dir, f"{base_name}_nobg.png")
        
        ok, msg = ImageEngine.save_base64_to_image(self.processed_bg_base64, out_path)
        
        if ok:
            self.txt_bg_status.value = f"Success! Saved to {out_path}"
            self.txt_bg_status.color = ft.Colors.GREEN_700
            self.show_snackbar("Image saved successfully!", color="green")
        else:
            self.txt_bg_status.value = f"Failed to save: {msg}"
            self.txt_bg_status.color = ft.Colors.RED_700
            self.show_snackbar("Failed to save image.", color="red")
            
        self.page.update()


    # ==========================================
    # 5. SETTINGS VIEW
    # ==========================================
    def build_settings_view(self) -> ft.Control:
        def on_theme_change(e):
            self.app_theme_mode = dropdown_theme.value
            self.page.theme_mode = ft.ThemeMode.DARK if self.app_theme_mode == "Dark" else ft.ThemeMode.LIGHT
            self.save_user_settings()
            self.page.update()

        def on_autosave_change(e):
            self.auto_save = switch_autosave.value
            self.save_user_settings()
            
        def on_company_change(e):
            self.company_name = txt_company_name.value
            self.company_address = txt_company_address.value
            self.save_user_settings()

        async def on_default_outdir_click(e):
            dir_path = await self.picker_batch_outdir.get_directory_path(dialog_title="Select Default Output Folder")
            if dir_path:
                self.default_output_dir = dir_path
                txt_default_outdir.value = dir_path
                self.save_user_settings()
                self.page.update()

        dropdown_theme = ft.Dropdown(
            label="App Theme",
            options=[ft.dropdown.Option("Light"), ft.dropdown.Option("Dark")],
            value=self.app_theme_mode,
            on_select=on_theme_change,
            width=200
        )
        
        switch_autosave = ft.Switch(
            label="Auto-Save Form Data",
            value=self.auto_save,
            on_change=on_autosave_change
        )
        
        txt_default_outdir = ft.TextField(
            label="Default Output Folder",
            value=self.default_output_dir,
            expand=True,
            read_only=True
        )
        
        txt_company_name = ft.TextField(
            label="Global Company Name",
            value=self.company_name,
            expand=True,
            on_change=on_company_change
        )
        
        txt_company_address = ft.TextField(
            label="Global Company Address",
            value=self.company_address,
            expand=True,
            on_change=on_company_change,
            multiline=True
        )

        card_general = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("General Settings", weight=ft.FontWeight.BOLD),
                    ft.Row([dropdown_theme, switch_autosave]),
                    ft.Row([
                        txt_default_outdir,
                        ft.IconButton(ft.Icons.FOLDER, on_click=on_default_outdir_click)
                    ]),
                ]),
                padding=15
            )
        )
        
        card_company = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("Global Company Data", weight=ft.FontWeight.BOLD),
                    txt_company_name,
                    txt_company_address
                ]),
                padding=15
            )
        )

        self.txt_settings_lo_path = ft.TextField(
            label="LibreOffice Binary Executable Path",
            value=self.libreoffice_path or "Not Found (COM Word Automation Fallback)",
            expand=True,
        )

        btn_detect_lo = ft.ElevatedButton(
            "Autodetect LibreOffice",
            icon=ft.Icons.SEARCH,
            on_click=self.on_detect_lo_clicked
        )

        btn_browse_lo = ft.IconButton(
            icon=ft.Icons.FOLDER,
            on_click=self.on_libreoffice_click
        )
        
        self.txt_settings_removebg_key = ft.TextField(
            label="Remove.bg API Key (Optional, speeds up background removal)",
            value=self.removebg_api_key,
            password=True,
            can_reveal_password=True,
            expand=True,
            on_change=self.on_removebg_api_key_changed
        )

        card_lo = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("LibreOffice & PDF Conversion Engine Settings", weight=ft.FontWeight.BOLD),
                    ft.Row([self.txt_settings_lo_path, btn_browse_lo, btn_detect_lo])
                ]),
                padding=15
            )
        )
        
        card_api = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Text("API Settings", weight=ft.FontWeight.BOLD),
                    ft.Row([self.txt_settings_removebg_key])
                ]),
                padding=15
            )
        )

        card_updater = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Icon(ft.Icons.SYSTEM_UPDATE_ALT, color=ft.Colors.BLUE_800, size=24),
                        ft.Text("Pembaruan Aplikasi (Auto-Updater)", weight=ft.FontWeight.BOLD, size=15, color=ft.Colors.BLUE_900),
                    ]),
                    ft.Text(
                        f"Versi saat ini: v{APP_VERSION} | Repository GitHub: monkdluffy/Me2D",
                        size=12,
                        color=ft.Colors.GREY_700
                    ),
                    ft.Row([
                        self.btn_check_update,
                        self.txt_update_status
                    ], spacing=15, vertical_alignment=ft.CrossAxisAlignment.CENTER)
                ], spacing=10),
                padding=15
            )
        )

        return ft.Column([
            ft.Text("Application Settings & Configuration", size=20, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
            card_general,
            card_company,
            card_lo,
            card_api,
            card_updater
        ], spacing=15)

    def build_help_view(self) -> ft.Control:
        card_help_updater = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Icon(ft.Icons.SYSTEM_UPDATE_ALT, color=ft.Colors.BLUE_800, size=24),
                        ft.Text("Status Pembaruan Sistem (Auto-Updater)", weight=ft.FontWeight.BOLD, size=15, color=ft.Colors.BLUE_900),
                    ]),
                    ft.Text(
                        "Aplikasi ini terhubung langsung dengan repository GitHub (monkdluffy/Me2D). Saat pembaruan baru dirilis, aplikasi akan otomatis memunculkan notifikasi pembaruan.",
                        size=12,
                        color=ft.Colors.GREY_700
                    ),
                    ft.Row([
                        ft.ElevatedButton(
                            "Periksa Pembaruan Sekarang",
                            icon=ft.Icons.REFRESH,
                            style=ft.ButtonStyle(bgcolor=ft.Colors.BLUE_700, color=ft.Colors.WHITE),
                            on_click=self.on_manual_check_update_clicked
                        ),
                        self.txt_update_status
                    ], spacing=15, vertical_alignment=ft.CrossAxisAlignment.CENTER)
                ], spacing=10),
                padding=15
            )
        )

        card_guide = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Icon(ft.Icons.MENU_BOOK, color=ft.Colors.BLUE_800, size=24),
                        ft.Text("Panduan Singkat Penggunaan Canido Doc Manager", weight=ft.FontWeight.BOLD, size=15, color=ft.Colors.BLUE_900)
                    ]),
                    ft.Text("1. Input Data: Lengkapi data perusahaan, personil, logbook, dan rekan bisnis pada tab 'Input Data'."),
                    ft.Text("2. Batch Find & Replace: Tentukan kata/kalimat khusus yang ingin dicari dan diganti pada form di Data Badan Usaha."),
                    ft.Text("3. Process Documents: Pilih folder atau file template Word (.docx), lalu klik tombol 'Start Batch Replacement' untuk memproses Word atau PDF secara otomatis."),
                    ft.Text("4. PDF Toolkit: Gunakan menu konversi, watermark, merge, dan split PDF bila diperlukan."),
                    ft.Text("5. Image Toolkit: Gunakan AI Background Removal untuk menghapus background stempel atau tanda tangan menjadi transparan 300 DPI.")
                ], spacing=8),
                padding=15
            )
        )

        card_about = ft.Card(
            content=ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Icon(ft.Icons.INFO_OUTLINE, color=ft.Colors.BLUE_800, size=24),
                        ft.Text("Tentang Aplikasi & Lisensi", weight=ft.FontWeight.BOLD, size=15, color=ft.Colors.BLUE_900)
                    ]),
                    ft.Text(f"{APP_NAME} v{APP_VERSION} © 2026 Canido Sertifikasi.", weight=ft.FontWeight.BOLD),
                    ft.Text("Dibuat untuk otomatisasi dokumen SMAP, audit, dan penerbitan dokumen sertifikasi."),
                    ft.Text("Official Repository: https://github.com/monkdluffy/Me2D", italic=True, color=ft.Colors.BLUE_800)
                ], spacing=6),
                padding=15
            )
        )

        return ft.Column([
            ft.Text("Pusat Bantuan & Pembaruan Sistem (Help Center)", size=20, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900),
            card_help_updater,
            card_guide,
            card_about
        ], spacing=15, scroll=ft.ScrollMode.ADAPTIVE)

    def _check_update_on_startup(self):
        import time
        time.sleep(2)
        try:
            update_info = check_for_updates(APP_VERSION)
            if update_info:
                self.show_update_dialog(update_info, is_manual=False)
        except Exception as ex:
            logger.debug(f"Startup update check: {ex}")

    def on_manual_check_update_clicked(self, e):
        self.txt_update_status.value = "Sedang memeriksa pembaruan di GitHub..."
        self.btn_check_update.disabled = True
        self.page.update()

        def run_check():
            try:
                update_info = check_for_updates(APP_VERSION)
                self.btn_check_update.disabled = False
                if update_info:
                    self.txt_update_status.value = f"Pembaruan tersedia: {update_info['tag_name']}"
                    self.page.update()
                    self.show_update_dialog(update_info, is_manual=True)
                else:
                    self.txt_update_status.value = f"Aplikasi sudah menggunakan versi terbaru (v{APP_VERSION})."
                    self.show_snackbar(f"Aplikasi Canido Doc Manager sudah versi terbaru (v{APP_VERSION}).", color="green")
                    self.page.update()
            except Exception as ex:
                self.btn_check_update.disabled = False
                self.txt_update_status.value = "Tidak dapat memeriksa pembaruan (offline / koneksi terputus)."
                self.page.update()

        threading.Thread(target=run_check, daemon=True).start()

    def show_update_dialog(self, update_info: dict, is_manual: bool = False):
        new_tag = update_info.get("tag_name", "Versi Baru")
        release_name = update_info.get("name") or new_tag
        body_notes = update_info.get("body") or "Pembaruan fitur & perbaikan bug."
        asset_name = update_info.get("asset_name") or "File Pembaruan"
        asset_size = update_info.get("asset_size", 0)
        size_str = f" (~{asset_size / (1024*1024):.1f} MB)" if asset_size > 0 else ""
        is_direct = update_info.get("is_direct_download", False)

        prog_bar = ft.ProgressBar(value=0.0, visible=False)
        lbl_status = ft.Text("Klik 'Perbarui Sekarang' untuk mengunduh dan memasang pembaruan.", size=12, color=ft.Colors.GREY_700)

        btn_close = ft.OutlinedButton("Nanti Saja")
        btn_action = ft.ElevatedButton(
            "Perbarui Sekarang" if is_direct else "Buka Halaman Download ➔",
            icon=ft.Icons.DOWNLOAD if is_direct else ft.Icons.OPEN_IN_NEW,
            style=ft.ButtonStyle(bgcolor=ft.Colors.GREEN_700, color=ft.Colors.WHITE)
        )

        def close_dlg(ev=None):
            dlg.open = False
            try:
                self.page.close(dlg)
            except Exception:
                pass
            self.page.update()

        btn_close.on_click = close_dlg

        def do_update(ev):
            if not is_direct:
                close_dlg()
                try:
                    self.page.launch_url(update_info.get("html_url", ""))
                except Exception:
                    pass
                return

            btn_action.disabled = True
            btn_close.disabled = True
            prog_bar.visible = True
            lbl_status.value = f"Mengunduh {asset_name}... 0%"
            self.page.update()

            def run_dl():
                import tempfile
                dest_dir = tempfile.gettempdir()
                dest_path = os.path.join(dest_dir, asset_name)

                def on_progress(fraction, downloaded, total):
                    pct = int(fraction * 100)
                    prog_bar.value = fraction
                    lbl_status.value = f"Mengunduh {asset_name}... {pct}% ({downloaded // (1024*1024)}/{total // (1024*1024)} MB)"
                    self.page.update()

                ok = download_update_file(update_info.get("download_url", ""), dest_path, progress_callback=on_progress)
                if ok and os.path.exists(dest_path):
                    lbl_status.value = "Unduhan selesai! Menjalankan installer..."
                    self.page.update()
                    import time
                    time.sleep(1)
                    close_dlg()
                    launch_installer_and_exit(dest_path)
                else:
                    btn_action.disabled = False
                    btn_close.disabled = False
                    prog_bar.visible = False
                    lbl_status.value = "Gagal mengunduh otomatis. Silakan buka halaman web rilis."
                    self.page.update()

            threading.Thread(target=run_dl, daemon=True).start()

        btn_action.on_click = do_update

        dlg = ft.AlertDialog(
            title=ft.Row([
                ft.Icon(ft.Icons.SYSTEM_UPDATE_ALT, color=ft.Colors.GREEN_700, size=28),
                ft.Text(f"Pembaruan Tersedia: {release_name}", weight=ft.FontWeight.BOLD, size=16, color=ft.Colors.BLUE_900),
            ]),
            content=ft.Container(
                width=500,
                content=ft.Column([
                    ft.Row([
                        ft.Container(
                            content=ft.Text(f"Versi Anda: v{APP_VERSION}", size=12, weight=ft.FontWeight.BOLD, color=ft.Colors.GREY_800),
                            padding=ft.padding.symmetric(horizontal=10, vertical=5),
                            bgcolor=ft.Colors.GREY_200,
                            border_radius=5
                        ),
                        ft.Icon(ft.Icons.ARROW_FORWARD, size=16, color=ft.Colors.BLUE_700),
                        ft.Container(
                            content=ft.Text(f"Versi Baru: {new_tag}", size=12, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
                            padding=ft.padding.symmetric(horizontal=10, vertical=5),
                            bgcolor=ft.Colors.GREEN_700,
                            border_radius=5
                        ),
                    ], spacing=10),
                    ft.Text("Catatan Pembaruan (Release Notes):", weight=ft.FontWeight.BOLD, size=13),
                    ft.Container(
                        content=ft.Column([
                            ft.Text(body_notes, size=12, color=ft.Colors.GREY_800),
                        ], scroll=ft.ScrollMode.ADAPTIVE),
                        max_height=140,
                        padding=10,
                        bgcolor=ft.Colors.BLUE_50,
                        border_radius=5,
                        border=ft.Border.all(1, ft.Colors.BLUE_100)
                    ),
                    ft.Text(f"Target Installer: {asset_name}{size_str}", size=12, italic=True, color=ft.Colors.GREY_700),
                    prog_bar,
                    lbl_status
                ], spacing=10),
            ),
            actions=[
                btn_close,
                btn_action
            ]
        )

        self.page.overlay.append(dlg)
        dlg.open = True
        self.page.update()
        
    def on_removebg_api_key_changed(self, e):
        self.removebg_api_key = self.txt_settings_removebg_key.value
        self.save_user_settings()

    def on_detect_lo_clicked(self, e):
        p = find_libreoffice_path()
        if p:
            self.libreoffice_path = p
            self.txt_settings_lo_path.value = p
            self.show_snackbar(f"LibreOffice detected: {p}", color="green")
        else:
            self.show_snackbar("LibreOffice not found. Will use MS Word COM automation on Windows.", color="blue")
        self.page.update()

    async def on_libreoffice_click(self, e):
        files = await self.picker_libreoffice.pick_files(dialog_title="Select LibreOffice soffice Executable")
        if files:
            self.libreoffice_path = files[0].path
            self.txt_settings_lo_path.value = files[0].path
            self.page.update()


def main(page: ft.Page):
    SMAPAutoDocApp(page)

def get_assets_dir():
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        return os.path.join(sys._MEIPASS, "assets")
    except Exception:
        return os.path.abspath("assets")

if __name__ == "__main__":
    ft.app(target=main, assets_dir=get_assets_dir())
