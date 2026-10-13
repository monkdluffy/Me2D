#!/bin/bash
# Canido Doc Manager - macOS Build Script
# 
# PENTING: Script ini HANYA BISA DIJALANKAN DI MESIN/LAPTOP MAC (macOS).
# Tidak bisa dijalankan dari Windows.

echo "========================================================"
echo "   Canido Doc Manager - macOS Build Preparation"
echo "========================================================"

# 1. Pastikan python dan pip terinstall
echo "[1/4] Memeriksa instalasi Python..."
if ! command -v python3 &> /dev/null
then
    echo "ERROR: python3 tidak ditemukan. Harap install Python 3 terlebih dahulu."
    exit 1
fi

# 2. Install dependensi
echo "[2/4] Menginstal dependensi Python (Flet)..."
python3 -m pip install -r requirements.txt

# 3. Proses Build / Pack
echo "[3/4] Mulai proses kompilasi aplikasi untuk macOS..."
flet pack app.py --name "Canido Doc Manager" --icon "assets/logo.png" --add-data "assets:assets"

# 4. Selesai
echo "[4/4] Selesai! Anda akan menemukan file aplikasi Mac (.app) di folder 'dist'."
echo "========================================================"
echo "Untuk mendistribusikan ke teman yang menggunakan Mac,"
echo "Anda bisa menaruh file .app tersebut di dalam sebuah folder,"
echo "lalu copy folder 'Master SMAP' dan 'Template_Batch_Data.xlsx'"
echo "ke folder yang sama, dan jadikan file .ZIP untuk dibagikan."
echo "========================================================"
