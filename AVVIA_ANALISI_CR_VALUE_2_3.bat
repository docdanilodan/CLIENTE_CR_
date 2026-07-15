@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
  echo Python non trovato. Installare Python 3.11 o 3.12 e selezionare Add Python to PATH.
  pause
  exit /b 1
)

py -m pip show streamlit >nul 2>nul
if errorlevel 1 (
  echo Installazione componenti necessari...
  py -m pip install -r requirements.txt
  if errorlevel 1 (
    echo Installazione non riuscita.
    pause
    exit /b 1
  )
)

echo Avvio Analisi CR VALUE 2.3...
py -m streamlit run Analisi_CR_VALUE_2_3_STABILE.py
pause
