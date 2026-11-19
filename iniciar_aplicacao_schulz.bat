@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Automacao de Certificados CMM - Schulz S.A.

echo ============================================================
echo   AUTOMACAO DE CERTIFICADOS CMM - SCHULZ S.A. (AV 05)
echo   Grupo 41 - Centro Universitario SENAI/SC
echo ============================================================
echo.

:: 1. Checagem de dependencias: Python 3.10+ no PATH
echo [1/4] Verificando instalacao do Python 3.10+...
set "PYTHON_CMD="

python --version >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=python"
) else (
    py --version >nul 2>&1
    if %errorlevel%==0 (
        set "PYTHON_CMD=py"
    )
)

if "%PYTHON_CMD%"=="" (
    echo [ERRO] Python nao encontrado no PATH do sistema.
    echo Instale o Python 3.10 ou superior em https://www.python.org/downloads/
    echo Certifique-se de marcar a opcao "Add Python to PATH" na instalacao.
    pause
    exit /b 1
)

for /f "tokens=2 delims= " %%v in ('%PYTHON_CMD% --version 2^>^&1') do set "PY_VER=%%v"
for /f "tokens=1,2 delims=." %%a in ("%PY_VER%") do (
    set "PY_MAJOR=%%a"
    set "PY_MINOR=%%b"
)

if %PY_MAJOR% LSS 3 (
    echo [ERRO] Versao do Python detectada: %PY_VER%. E necessario Python 3.10 ou superior.
    pause
    exit /b 1
)
if %PY_MAJOR% EQU 3 if %PY_MINOR% LSS 10 (
    echo [ERRO] Versao do Python detectada: %PY_VER%. E necessario Python 3.10 ou superior.
    pause
    exit /b 1
)
echo [OK] Python %PY_VER% detectado.
echo.

:: 2. Criacao e ativacao do ambiente virtual (venv)
echo [2/4] Verificando ambiente virtual (venv)...
if not exist "venv\Scripts\activate.bat" (
    echo Criando ambiente virtual isolado...
    %PYTHON_CMD% -m venv venv
    if %errorlevel% neq 0 (
        echo [ERRO] Falha ao criar o ambiente virtual.
        pause
        exit /b 1
    )
)
call venv\Scripts\activate.bat
echo [OK] Ambiente virtual ativado.
echo.

:: 3. Instalacao automatizada das bibliotecas do requirements.txt
echo [3/4] Instalando/atualizando dependencias do requirements.txt...
if not exist "requirements.txt" (
    echo [ERRO] Arquivo requirements.txt nao encontrado na raiz do projeto.
    pause
    exit /b 1
)
python -m pip install --upgrade pip --quiet
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERRO] Falha ao instalar as dependencias. Verifique sua conexao com a internet.
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas com sucesso.
echo.

:: 4. Inicializacao do servico e abertura automatica no navegador
echo [4/4] Iniciando a aplicacao Streamlit...
if not exist "app.py" (
    echo [ERRO] Arquivo app.py nao encontrado na raiz do projeto.
    pause
    exit /b 1
)

echo Servidor iniciando em http://localhost:8501
echo Pressione CTRL+C nesta janela para encerrar a aplicacao.
echo.
python -m streamlit run app.py --server.port 8501

pause