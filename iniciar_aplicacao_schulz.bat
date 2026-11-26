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
echo [1/5] Verificando instalacao do Python 3.10+...
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
echo [OK] Python %PY_VER% detectado (SQLite3 nativo disponivel).
echo.

:: 2. Criacao e ativacao do ambiente virtual (venv)
echo [2/5] Verificando ambiente virtual (venv)...
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
echo [3/5] Instalando/atualizando dependencias do requirements.txt...
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

:: 4. Verificacao do config.txt e estrutura de pastas (certificados, historico e dados)
echo [4/5] Verificando config.txt e estrutura de diretorios...
set "PASTA_CERTIFICADOS=certificados"
set "PASTA_DADOS=dados"

if not exist "config.txt" (
    echo Criando arquivo padrao config.txt...
    (
        echo # Configuracao de diretorios do Gerador de Certificados CMM - Schulz
        echo # Use caminhos relativos ^(ex: certificados^) ou absolutos ^(ex: C:\Schulz\Certificados^)
        echo PASTA_CERTIFICADOS=certificados
        echo PASTA_DADOS=dados
    ) > "config.txt"
)

for /f "usebackq eol=# tokens=1,* delims==" %%a in ("config.txt") do (
    if /i "%%a"=="PASTA_CERTIFICADOS" if not "%%b"=="" set "PASTA_CERTIFICADOS=%%b"
    if /i "%%a"=="PASTA_DADOS" if not "%%b"=="" set "PASTA_DADOS=%%b"
)

if not exist "%PASTA_CERTIFICADOS%" mkdir "%PASTA_CERTIFICADOS%"
if not exist "%PASTA_CERTIFICADOS%\historico" mkdir "%PASTA_CERTIFICADOS%\historico"
if not exist "%PASTA_DADOS%" mkdir "%PASTA_DADOS%"

echo [OK] Pasta de certificados : %PASTA_CERTIFICADOS%
echo [OK] Historico de revisoes : %PASTA_CERTIFICADOS%\historico
echo [OK] Pasta de dados (DB/Excel): %PASTA_DADOS%
echo.

:: 5. Inicializacao do servico e abertura automatica no navegador
echo [5/5] Iniciando a aplicacao Streamlit...
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