@echo off
echo ===================================================
echo    CREANDO EJECUTABLE DEL GESTOR DE CSV
echo ===================================================
echo.
echo Instalando PyInstaller...
pip install pyinstaller

echo.
echo Generando el ejecutable...
python -m PyInstaller csv_manager_exe.spec

echo.
echo ===================================================
echo PROCESO COMPLETADO!
echo.
echo El ejecutable se encuentra en la carpeta "dist"
echo con el nombre "Gestor de CSV.exe"
echo ===================================================
echo.

pause