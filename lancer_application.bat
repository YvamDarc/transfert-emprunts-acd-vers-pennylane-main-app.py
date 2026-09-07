@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
  echo Python n'est pas installe ou la commande "py" n'est pas disponible.
  echo Installez Python 3.11 ou une version plus recente, puis relancez ce fichier.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Creation de l'environnement Python...
  py -m venv .venv
)

echo Installation ou verification des dependances...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo L'installation des dependances a echoue.
  pause
  exit /b 1
)

echo Ouverture de l'application Streamlit...
".venv\Scripts\python.exe" -m streamlit run app.py
pause

