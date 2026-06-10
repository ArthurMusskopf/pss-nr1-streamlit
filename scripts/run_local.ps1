$ErrorActionPreference = "Stop"

Write-Host "== PS&S NR-1 | Execucao local ==" -ForegroundColor Cyan

# Preferir Python Launcher do Windows; evita pegar Python do Git Bash/MSYS/WSL por engano.
$pythonCmd = $null
if (Get-Command py -ErrorAction SilentlyContinue) {
  $pythonCmd = "py"
  $pythonArgs = @("-3")
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
  $pythonCmd = "python"
  $pythonArgs = @()
} else {
  throw "Python nao encontrado. Instale com: winget install -e --id Python.Python.3.12"
}

Write-Host "Usando Python:" -ForegroundColor Yellow
& $pythonCmd @pythonArgs --version

# Se existe um .venv quebrado, remove antes de recriar.
if ((Test-Path ".venv") -and (-not (Test-Path ".venv\Scripts\Activate.ps1"))) {
  Write-Host "Ambiente virtual .venv quebrado/incompleto encontrado. Removendo..." -ForegroundColor Yellow
  Remove-Item -Recurse -Force ".venv"
}

if (-not (Test-Path ".venv")) {
  Write-Host "Criando ambiente virtual .venv..." -ForegroundColor Cyan
  & $pythonCmd @pythonArgs -m venv .venv
}

if (-not (Test-Path ".venv\Scripts\Activate.ps1")) {
  throw "O .venv foi criado sem Scripts\Activate.ps1. Provavelmente o comando 'python' aponta para uma distribuicao nao-Windows. Rode: where.exe python ; where.exe py"
}

Write-Host "Ativando ambiente virtual..." -ForegroundColor Cyan
. .\.venv\Scripts\Activate.ps1

Write-Host "Atualizando pip..." -ForegroundColor Cyan
python -m pip install --upgrade pip

Write-Host "Instalando dependencias..." -ForegroundColor Cyan
python -m pip install -r requirements.txt

Write-Host "Abrindo Streamlit..." -ForegroundColor Green
python -m streamlit run app.py
