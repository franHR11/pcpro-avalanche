@echo off
echo Instalando dependencias...
python -m pip install pyinstaller
python -m pip install ttkbootstrap

echo Creando ejecutable...
REM Usamos la especificación definida en envioemail.spec
python -m PyInstaller envioemail.spec

echo Limpiando archivos temporales...
rmdir /s /q build
del /q *.spec

echo Moviendo ejecutable...
move dist\EnvioEmails.exe .
rmdir /s /q dist

echo Proceso completado!
pause
