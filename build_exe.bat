@echo off
REM ============================================================
REM  Compila CampusCentinela.exe en Windows.
REM  Requiere tener Python instalado (python.org).
REM  Doble clic en este archivo.
REM ============================================================
echo Instalando dependencias...
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller
echo.
echo Compilando ejecutable...
pyinstaller --onefile --windowed --name CampusCentinela --clean campus_centinela_gui.py
echo.
echo ============================================================
echo  Listo. El ejecutable quedo en:  dist\CampusCentinela.exe
echo ============================================================
pause
