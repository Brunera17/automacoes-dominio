@echo off
rem Duplo clique: abre a janela de instalacao/configuracao do ouvinte.
start "" powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0instalar.ps1"
