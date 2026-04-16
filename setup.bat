@echo off
chcp 65001 >nul
echo.
echo ============================================
echo  Vstanovlennya zalezhnostei dlya Generatora
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [!] Python ne znaideno.
    echo.
    echo     Vstanovit Python z https://python.org/downloads/
    echo     Na pershomu ekrani obovyazkovo postavte galochku:
    echo       "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo [OK] %PYVER%
echo.

echo [*] Perevirka WeasyPrint...
python -c "from weasyprint import HTML" >nul 2>&1
if not errorlevel 1 (
    echo [OK] WeasyPrint + GTK vzhe vstanovleni, propuskaemo.
    goto :pip
)

echo [*] Zavantazhennya GTK Runtime z GitHub...
echo     Tse mozhe zainyaty khvylinu...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference='Stop';" ^
    "try {" ^
    "  $rel = Invoke-RestMethod 'https://api.github.com/repos/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases/latest';" ^
    "  $asset = $rel.assets | Where-Object { $_.name -like '*.exe' } | Select-Object -First 1;" ^
    "  if (-not $asset) { throw 'GTK installer not found' }" ^
    "  Write-Host ('  -> ' + $asset.name);" ^
    "  Invoke-WebRequest $asset.browser_download_url -OutFile 'gtk_setup.exe' -UseBasicParsing;" ^
    "  Write-Host '  -> Running installer (silent)...';" ^
    "  Start-Process 'gtk_setup.exe' -ArgumentList '/S' -Wait;" ^
    "  Remove-Item 'gtk_setup.exe';" ^
    "  Write-Host '  -> GTK installed.'" ^
    "} catch { Write-Host ('ERROR: ' + $_.Exception.Message); exit 1 }"

if errorlevel 1 (
    echo.
    echo [!] Ne vdalosia vstanovyty GTK avtomatychno.
    echo     Zavantazhte vruchnu:
    echo     https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases
    echo     Fail: gtk3-runtime-*-win64.exe
    echo     Pislia vstanovlennya zapustit setup.bat znovu.
    echo.
    pause
    exit /b 1
)

echo.

:pip
echo [*] Vstanovlennya Python-paketiv...
pip install --upgrade openpyxl num2words weasyprint
if errorlevel 1 (
    echo.
    echo [!] Pomylka pid chas pip install.
    echo     Perevirte pidklyuchennya do internetu i sprobuite shche raz.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Gotovo! Zapuskajte programu cherez run.bat
echo ============================================
echo.
pause
