@echo off
chcp 65001 >nul

echo (Opzionale) Si consiglia di eseguire questo script come Amministratore per evitare problemi di permessi.

REM Script semplificato: verifica Python/pip e installa dipendenze

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python non trovato.
    winget --version >nul 2>&1
    if %errorlevel% equ 0 (
        echo winget trovato: provo a installare Python con winget...
        winget install --id Python.Python.3 -e --accept-package-agreements --accept-source-agreements
        if %errorlevel% neq 0 (
            echo Errore durante l'installazione di Python tramite winget.
            pause >nul
            exit /b 1
        )
        echo Python installato (o in installazione). Chiudi e riapri la finestra del terminale, poi riesegui questo script.
        pause >nul
        exit /b 0
    ) else (
        echo winget non disponibile. Installa Python manualmente e assicurati di selezionare "Add Python to PATH".
        pause >nul
        exit /b 1
    )
)

python -m pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo pip non trovato: provo a installare pip con ensurepip...
    python -m ensurepip --upgrade >nul 2>&1 || (
        echo Impossibile installare pip automaticamente. Installa pip manualmente.
        pause >nul
        exit /b 1
    )
)

echo Aggiorno pip...
python -m pip install --upgrade pip

echo Installazione delle librerie principali suggerite dall'utente...
python -m pip install "openpyxl>=3.0.0" "reportlab>=3.6.0"
if %errorlevel% neq 0 (
    echo Errore durante l'installazione dei pacchetti principali.
    echo Prova a eseguire manualmente: python -m pip install "openpyxl>=3.0.0" "reportlab>=3.6.0"
    pause >nul
    exit /b 1
)

if exist requirements.txt (
    echo Trovato requirements.txt: lo rimuovo (le dipendenze sono integrate nello script)...
    del /f /q requirements.txt >nul 2>&1
    if %errorlevel% equ 0 (
        echo requirements.txt rimosso.
    ) else (
        echo Impossibile rimuovere requirements.txt (controlla permessi).
    )
)

echo Installazione completata.
pause >nul
