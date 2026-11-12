@echo off
rem Inicia o ouvinte do hub da Dominio (modo definido em config.json -> "simulacao")
cd /d "%~dp0"
python ouvinte.py %*
