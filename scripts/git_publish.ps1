param(
  [string] $RepoName = "pss-nr1-streamlit",
  [switch] $Private = $true
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path ".git")) {
  git init
}

git add .
git commit -m "feat: app Streamlit NR-1 com BigQuery, certificado e matriz" 2>$null
if ($LASTEXITCODE -ne 0) { Write-Host "Nada novo para commitar ou commit já existente." -ForegroundColor Yellow }

$visibility = if ($Private) { "--private" } else { "--public" }

try {
  gh repo view $RepoName | Out-Null
  Write-Host "Repositório remoto já existe. Continuando push..." -ForegroundColor Yellow
} catch {
  gh repo create $RepoName $visibility --source=. --remote=origin --push
  exit 0
}

git branch -M main
git remote get-url origin 2>$null
if ($LASTEXITCODE -ne 0) {
  gh repo create $RepoName $visibility --source=. --remote=origin
}
git push -u origin main
