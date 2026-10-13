import os
import sys
from typing import Dict, List, Any

APP_NAME = "Canido Doc Manager"
APP_VERSION = "1.0.0"

# Canonical SMAP 78 Variables categorized
SMAP_VARIABLE_CATEGORIES: Dict[str, Dict[str, Dict[str, str]]] = {
    "Data Badan Usaha": {
        "inisial_bu": {"label": "Inisial (BU) Badan Usaha (e.g. PT cbu)", "default": ""},
        "bu_caps": {"label": "Nama BU (HURUF BESAR)", "default": ""},
        "bu": {"label": "Nama BU (Standard)", "default": ""},
        "alamat_bu": {"label": "Alamat Lengkap BU", "default": ""},
        "kab_kota": {"label": "Kabupaten / Kota", "default": ""},

        "website": {"label": "Website BU", "default": ""},
        "email_bu": {"label": "Email Resmi BU", "default": ""},
        "telepon_bu": {"label": "Telepon BU", "default": ""},
        "jenis_usaha": {"label": "Jenis Usaha", "default": "Jasa Pekerjaan Kontruksi", "type": "dropdown", "options": ["Jasa Pekerjaan Kontruksi", "Jasa Konsultansi Konstruksi"]},
        "akta_pendirian": {"label": "No. Akta Pendirian", "default": ""},
        "tgl_akta_pendirian": {"label": "Tgl Akta Pendirian", "default": ""},
        "akta_perubahan_akhir": {"label": "No. Akta Perubahan Terakhir", "default": ""},
        "tgl_akta_perubahan_akhir": {"label": "Tgl Akta Perubahan Terakhir", "default": ""},
        "npwp_bu": {"label": "NPWP Badan Usaha", "default": ""},
        "tgl_npwp_bu": {"label": "Tgl Terbit NPWP", "default": ""},
        "nib_bu": {"label": "NIB (Nomor Induk Berusaha)", "default": ""},
        "tgl_nib_bu": {"label": "Tgl Terbit NIB", "default": ""},
        "no_kumham": {"label": "No. SK Kemenkumham", "default": ""},
        "serti_standar_1": {"label": "Sertifikat Standar #1", "default": ""},
        "tgl_serti_standar_1": {"label": "Tgl Sertifikat Standar #1", "default": ""},
        "sbu_umku_1": {"label": "SBU / UMKU #1", "default": ""},
        "tgl_sbu_umku_1": {"label": "Tgl SBU / UMKU #1", "default": ""},
        "tgl_berlaku": {"label": "Tgl Berlaku Dokumen", "default": ""},
        "tgl_berlaku_pendek": {"label": "Tgl Berlaku Format Pendek", "default": ""},
        "tgl_sekarang": {"label": "Tgl Dokumen / TTD", "default": ""},
        "kbli": {"label": "Kode KBLI", "default": ""},
        "bu_folder": {"label": "Nama Folder BU", "default": ""}
    },
    "Personil & Staff SMAP": {
        "direktur": {"label": "Direktur", "default": ""},
        "fkap": {"label": "FKAP", "default": ""},
        "ins_fkap": {"label": "Inisial Div. FKAP", "default": ""},
        "manager": {"label": "Div. HRD/Manager", "default": ""},
        "ins_mgr": {"label": "Inisial Div. HRD/Manager", "default": ""},
        "keuangan": {"label": "Div. Keuangan", "default": ""},
        "ins_keu": {"label": "Inisial Div. Keuangan", "default": ""},
        "administrasi": {"label": "Div. Administrasi", "default": ""},
        "ins_adm": {"label": "Inisial Div. Administrasi", "default": ""}
    },
    "Logbook & Data Audit": {
        "tgl_logbook": {"label": "Tgl Logbook", "default": ""},
        "tgl_lapor_suap": {"label": "Tgl Lapor Suap", "default": ""},
        "tgl_analisa_pelaporan": {"label": "Tgl Analisa Pelaporan", "default": ""},
        "no_surat_penunjukan": {"label": "No Surat Penunjukan", "default": ""},
        "tgl_surat_penunjukan": {"label": "Tgl Surat Penunjukan", "default": ""},
        "th_program_audit_th4": {"label": "Tahun Prog. Audit (th4)", "default": ""},
        "tgl_program_audit_th4": {"label": "Tgl Program Audit Th4", "default": ""},
        "tgl_jadwal_rencana_audit": {"label": "Tgl Jadwal Rencana Audit", "default": ""},
        "no_undangan_audit": {"label": "No Undangan Audit", "default": ""},
        "tgl1_undangan_audit": {"label": "Tgl1 Undangan Audit", "default": ""},
        "tgl2_undangan_audit": {"label": "Tgl2 Undangan Audit", "default": ""},
        "hari_undangan_audit": {"label": "Hari Undangan Audit", "default": ""},
        "tgl_uji_layak_personil": {"label": "Tgl Uji Layak Personil", "default": ""},
        "tgl_uji_layak_rekan": {"label": "Tgl Uji Layak Rekan", "default": ""},
        "tgl_daftar_periksa": {"label": "Tgl Daftar Periksa", "default": ""},
        "tgl_daftar_periksa1": {"label": "Tgl Daftar Periksa1", "default": ""},
        "tgl_daftar_periksa2": {"label": "Tgl Daftar Periksa2", "default": ""},
        "no_lap_tidak_sesuai": {"label": "No Lap Tidak Sesuai", "default": ""},
        "tgl_selesai1": {"label": "Tgl Selesai1", "default": ""},
        "tgl_selesai2": {"label": "Tgl Selesai2", "default": ""},
        "no_ba_pemusnahan": {"label": "No BA Pemusnahan", "default": ""},
        "tgl_ba_pemusnahan": {"label": "Tgl BA Pemusnahan", "default": ""},
        "no_terima_distribusi": {"label": "No Terima Distribusi", "default": ""},
        "tgl_terima_distribusi": {"label": "Tgl Terima Distribusi", "default": ""},
        "haritgl_terbilang_ba_pemusnahan": {"label": "HariTgl Terbilang BA Pemusnahan", "default": ""}
    },
    "Rekan Bisnis": {
        "nama_rekan": {"label": "Nama Penanggung Jawab", "default": ""},
        "bu_dinas_rekan": {"label": "Badan Usaha (BU) Rekan/Dinas", "default": ""},
        "alamat_rekan": {"label": "Alamat BU Rekan/Dinas", "default": ""},
        "hub_rekan": {"label": "Hub. Dengan Rekan/Dinas", "default": ""},
        "nik_nip": {"label": "NIK/NIP Penanggung Jawab", "default": ""},
        "jabatan_rekan": {"label": "Jabatan Penanggung Jawab", "default": ""},
        "tlp_rekan": {"label": "No. Tlp Rekan/Dinas", "default": ""}
    }
}

# Flatten 77 variable keys for quick validation & lookup
ALL_SMAP_VARIABLES: List[str] = []
DEFAULT_VARIABLE_VALUES: Dict[str, str] = {}

for category, vars_dict in SMAP_VARIABLE_CATEGORIES.items():
    for var_name, meta in vars_dict.items():
        ALL_SMAP_VARIABLES.append(var_name)
        DEFAULT_VARIABLE_VALUES[var_name] = meta["default"]

# Backward compatibility alias for jenis_usaha_bu
if "jenis_usaha" in ALL_SMAP_VARIABLES and "jenis_usaha_bu" not in ALL_SMAP_VARIABLES:
    ALL_SMAP_VARIABLES.append("jenis_usaha_bu")
    DEFAULT_VARIABLE_VALUES["jenis_usaha_bu"] = DEFAULT_VARIABLE_VALUES.get("jenis_usaha", "Jasa Pekerjaan Kontruksi")

# Dynamic Image Slot Configuration
IMAGE_SLOTS: Dict[str, Dict[str, Any]] = {
    "logo_bu": {"label": "Logo Badan Usaha (BU) Utama", "tag": "logo_bu", "default_width_mm": 40, "default_height_mm": 38},
    "logo_bu_header": {"label": "Logo Badan Usaha (BU) Header", "tag": "logo_bu_header", "default_width_mm": 40, "default_height_mm": 21},
    "stempel_bu": {"label": "Stempel/CAP BU", "tag": "stempel_bu", "default_width_mm": 35, "default_height_mm": 30},
    "ttd_direktur": {"label": "TTD Direktur", "tag": "ttd_direktur", "default_width_mm": 40, "default_height_mm": 23},
    "ttd_fkap": {"label": "TTD FKAP", "tag": "ttd_fkap", "default_width_mm": 40, "default_height_mm": 23},
    "ttd_manager": {"label": "TTD Manager", "tag": "ttd_manager", "default_width_mm": 35, "default_height_mm": 23},
    "ttd_adm": {"label": "TTD Administrasi", "tag": "ttd_adm", "default_width_mm": 35, "default_height_mm": 23},
    "ttd_keu": {"label": "TTD Keuangan", "tag": "ttd_keu", "default_width_mm": 35, "default_height_mm": 23},
    "logo_rekan": {"label": "Logo BU Rekan/Dinas", "tag": "logo_rekan", "default_width_mm": 40, "default_height_mm": 20},
    "ttd_rekan": {"label": "TTD PJ Rekan/Dinas", "tag": "ttd_rekan", "default_width_mm": 40, "default_height_mm": 23},
}

# Cross-platform LibreOffice Binary Search Paths
LIBREOFFICE_PATHS_WIN = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\LibreOffice\program\soffice.exe"),
]

LIBREOFFICE_PATHS_MAC = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice.bin",
    "/usr/local/bin/soffice",
    "/opt/homebrew/bin/soffice",
]

def find_libreoffice_path() -> str:
    """Detect installed LibreOffice binary path across Windows & macOS."""
    if sys.platform == "win32":
        for path in LIBREOFFICE_PATHS_WIN:
            if os.path.exists(path):
                return path
    elif sys.platform == "darwin":
        for path in LIBREOFFICE_PATHS_MAC:
            if os.path.exists(path):
                return path
    # System PATH fallback
    import shutil
    soffice_in_path = shutil.which("soffice")
    if soffice_in_path:
        return soffice_in_path
    return ""
