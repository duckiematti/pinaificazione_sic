@echo off
chcp 65001 > nul
title Pianificazione Corsi 2026
color 0B

echo.
echo ╔═══════════════════════════════════════════════════════════════════════╗
echo ║                                                                       ║
echo ║       🚀 PIANIFICAZIONE CORSI                                         ║
echo ║                                                                       ║
echo ╚═══════════════════════════════════════════════════════════════════════╝
echo.

REM ===== VERIFICA PYTHON =====
python --version >nul 2>&1
if %errorlevel% neq 0 (
    color 0C
    echo.
    echo ❌ Python non trovato!
    echo.
    echo 💡 Esegui prima: installa_dipendenze.bat
    echo.
    pause
    exit /b 1
)

REM ===== VERIFICA DIPENDENZE =====
echo 🔍 Verifica dipendenze...
python -c "import openpyxl, reportlab" >nul 2>&1
if %errorlevel% neq 0 (
    color 0E
    echo.
    echo ⚠️  Dipendenze mancanti!
    echo.
    echo 💡 SOLUZIONE:
    echo    1. Chiudi questa finestra
    echo    2. Esegui: installa_dipendenze.bat
    echo    3. Poi riprova
    echo.
    pause
    exit /b 1
)
echo ✅ Dipendenze OK
echo.

REM ===== VERIFICA/CREA FILE EXCEL =====
if not exist "Pianificazione_Corsi_2026.xlsx" (
    echo 📝 Prima esecuzione: creazione file Excel...
    echo.
    python crea_pianificazione_smart.py
    if %errorlevel% neq 0 (
        color 0C
        echo.
        echo ❌ Errore nella creazione del file Excel
        echo.
        pause
        exit /b 1
    )
    echo.
)

REM ===== AVVIO SISTEMA =====
color 0A
echo ╔═══════════════════════════════════════════════════════════════════════╗
echo ║                                                                       ║
echo ║       ✅ SISTEMA PRONTO                                               ║
echo ║                                                                       ║
echo ╚═══════════════════════════════════════════════════════════════════════╝
echo.
echo 🌐 Avvio server web...
echo.
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo.
echo    � Apertura browser tra 2 secondi...
echo.
echo    🌐 URL: http://localhost:8765
echo.
echo    ⏹️  Per fermare: Premi CTRL+C
echo.
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo.

REM ===== AVVIA SERVER IN BACKGROUND =====
start /B python server.py

REM ===== ATTENDI CHE IL SERVER SIA PRONTO =====
timeout /t 2 /nobreak > nul

REM ===== APRI BROWSER =====
start http://localhost:8765

REM ===== MANTIENI FINESTRA APERTA =====
echo.
echo ✅ Server attivo
echo.
echo 💡 NON chiudere questa finestra!
echo    Il server continuerà a funzionare finché non la chiudi.
echo.
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo.

REM ===== ATTENDI CHIUSURA =====
pause
