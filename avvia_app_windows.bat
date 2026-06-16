@echo off
chcp 65001 >nul

REM --- Si sposta SEMPRE nella cartella di questo file .bat ---
REM Cosi' app.py, .venv e .streamlit\secrets.toml vengono sempre trovati,
REM qualunque sia la cartella da cui parte Windows.
cd /d "%~dp0"

REM --- Attiva l'ambiente virtuale ---
call ".venv\Scripts\activate.bat"

REM --- Avvia l'app e apre il browser sulla pagina corretta ---
REM headless=false fa aprire il browser da solo su http://localhost:8501
streamlit run app.py --server.port 8501 --server.headless false

pause