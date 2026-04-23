@echo off
echo Instalando dependencias necesarias...
pip install pandas openpyxl pyinstaller
echo Generando archivo ejecutable (.exe)...
pyinstaller --onefile --name ConciliadorFinanciero main.py
echo.
echo Proceso terminado. El archivo se encuentra en la carpeta 'dist'
pause
