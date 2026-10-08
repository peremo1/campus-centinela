@echo off
REM ============================================================
REM  Compila CampusCentinela.exe en Windows (interfaz Flet).
REM  Requiere tener Python instalado (python.org).
REM  Doble clic en este archivo.
REM ============================================================
echo Instalando dependencias...
python -m pip install --upgrade pip
pip install -r requirements.txt -r requirements-gui.txt pyinstaller
echo.
echo Compilando ejecutable (puede tardar un par de minutos)...
flet pack campus_centinela_flet.py --name CampusCentinela --product-name "Campus Centinela"
echo.
echo ============================================================
echo  Listo. El ejecutable quedo en:  dist\CampusCentinela.exe
echo ============================================================
pause
