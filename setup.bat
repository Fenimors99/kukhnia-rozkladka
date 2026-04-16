@echo off
chcp 65001 >nul
echo.
echo ============================================
echo  Встановлення залежностей для Генератора
echo ============================================
echo.

:: ── 1. Перевірка Python ───────────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo [!] Python не знайдено.
    echo.
    echo     Встановіть Python з https://python.org/downloads/
    echo     На першому екрані обов'язково поставте галочку:
    echo       "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo [OK] %PYVER%
echo.

:: ── 2. GTK (потрібен для WeasyPrint / PDF) ────────────────────────────────────
echo [*] Перевірка GTK...

python -c "from weasyprint import HTML" >nul 2>&1
if not errorlevel 1 (
    echo [OK] WeasyPrint + GTK вже встановлені, пропускаємо.
    goto :pip
)

echo [*] Завантаження GTK Runtime з GitHub...
echo     Це може зайняти хвилину...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference='Stop';" ^
    "try {" ^
    "  $rel = Invoke-RestMethod 'https://api.github.com/repos/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases/latest';" ^
    "  $asset = $rel.assets | Where-Object { $_.name -like '*installer.exe' } | Select-Object -First 1;" ^
    "  if (-not $asset) { throw 'Не знайдено інсталятор GTK' }" ^
    "  Write-Host ('  -> ' + $asset.name);" ^
    "  Invoke-WebRequest $asset.browser_download_url -OutFile 'gtk_setup.exe' -UseBasicParsing;" ^
    "  Write-Host '  -> Запускаємо інсталятор (тихий режим)...';" ^
    "  Start-Process 'gtk_setup.exe' -ArgumentList '/S' -Wait;" ^
    "  Remove-Item 'gtk_setup.exe';" ^
    "  Write-Host '  -> GTK встановлено.'" ^
    "} catch { Write-Host ('ПОМИЛКА: ' + $_.Exception.Message); exit 1 }"

if errorlevel 1 (
    echo.
    echo [!] Не вдалося встановити GTK автоматично.
    echo     Завантажте вручну:
    echo     https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases
    echo     Файл: gtk3-runtime-*-installer.exe
    echo     Після встановлення запустіть setup.bat знову.
    echo.
    pause
    exit /b 1
)

echo.

:: ── 3. Python-пакети ──────────────────────────────────────────────────────────
:pip
echo [*] Встановлення Python-пакетів...
pip install --upgrade openpyxl num2words weasyprint
if errorlevel 1 (
    echo.
    echo [!] Помилка під час pip install.
    echo     Перевірте підключення до інтернету і спробуйте ще раз.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Готово! Запускайте програму через run.bat
echo ============================================
echo.
pause
