@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Nextcloud Teams Helfer
echo.
echo Starte lokalen Nextcloud-Helfer ...
echo (Fenster offen lassen, solange du in der App browsen willst)
echo.
where node >nul 2>&1
if errorlevel 1 (
  echo FEHLER: Node.js wurde nicht gefunden.
  echo Bitte Node.js installieren: https://nodejs.org/
  echo.
  pause
  exit /b 1
)
if not exist ".env.local" (
  echo FEHLER: .env.local fehlt in diesem Ordner.
  echo Vorlage kopieren: copy .env.example .env.local
  echo Dann Benutzer und App-Passwort eintragen.
  echo.
  pause
  exit /b 1
)
node proxy-server.mjs
echo.
echo Helfer beendet.
pause
