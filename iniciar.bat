@echo off
echo.
echo  ========================================
echo   BACCO - Dashboard Gerencial
echo  ========================================
echo.
cd /d "%~dp0"

IF NOT EXIST .env (
  echo  Arquivo .env nao encontrado.
  echo  Copie .env.example para .env e preencha suas credenciais.
  echo.
  pause
  exit /b 1
)

start http://localhost:3001
echo  Servidor iniciado em http://localhost:3001
echo  Mantenha esta janela aberta.
echo.
node server.js
pause
