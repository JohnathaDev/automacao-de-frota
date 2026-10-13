@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo Iniciando Automação de Frota (Linha de Comando)...
python automacao_frota_ce.py
pause
