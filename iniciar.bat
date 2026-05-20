@echo off
echo Instalando dependencias...
pip install flask pdfplumber openpyxl --quiet
echo.
echo Iniciando servidor BICA PDF a Excel...
echo Abriendo navegador en http://localhost:5050
echo.
echo Para cerrar: presiona Ctrl+C o cierra esta ventana.
echo.
python app.py
pause
