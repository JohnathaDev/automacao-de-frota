@echo off
chcp 65001 > nul
where pythonw >nul 2>&1
if %errorlevel% equ 0 (
    start "" pythonw automacao_frota_gui.py
) else (
    start "" python automacao_frota_gui.py
)
exit
