@echo off
title MIC DIGITAL CONSULTING - Validador de Estados Contables (RT 54 / CPCECABA)
cd /d "C:\Users\mcres\validador"
echo ===============================================================================
echo   MIC DIGITAL CONSULTING - VALIDADOR PRE-LEGALIZACION DE EECC (RT 54 / CPCECABA)
echo ===============================================================================
echo Iniciando servidor web en http://127.0.0.1:8054 ...
start "" http://127.0.0.1:8054
python app.py
pause
