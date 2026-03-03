@echo off
chcp 65001 >nul
echo ============================================
echo   Karsilastirici - Windows Kurulum Scripti
echo ============================================
echo.

REM Python kontrol
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [HATA] Python bulunamadi!
    echo Lutfen https://www.python.org/downloads/ adresinden Python yukleyin.
    echo Yuklerken "Add Python to PATH" kutucugunu isaretlemeyi unutmayin!
    pause
    exit /b 1
)

echo [1/3] Sanal ortam olusturuluyor...
python -m venv .venv
if %errorlevel% neq 0 (
    echo [HATA] Sanal ortam olusturulamadi!
    pause
    exit /b 1
)

echo [2/3] Kutuphaneler yukleniyor...
.venv\Scripts\pip install pandas openpyxl xlrd customtkinter pyinstaller
if %errorlevel% neq 0 (
    echo [HATA] Kutuphaneler yuklenemedi!
    pause
    exit /b 1
)

echo [3/3] EXE dosyasi olusturuluyor...
.venv\Scripts\pyinstaller --noconfirm --onefile --windowed --name "Karsilastirici" --add-data "engine.py;." app.py
if %errorlevel% neq 0 (
    echo [HATA] EXE olusturulamadi!
    pause
    exit /b 1
)

echo.
echo ============================================
echo   BASARILI!
echo   EXE dosyasi: dist\Karsilastirici.exe
echo ============================================
echo.
echo Bu dosyayi istediginiz yere kopyalayabilirsiniz.
echo Calistirmak icin cift tiklayin.
echo.
pause
