param(
  [Parameter(Mandatory=$true)] [string] $ProjectId,
  [string] $DatasetId = "pss_nr1",
  [string] $Location = "US",
  [string] $ServiceAccountName = "streamlit-pss-nr1"
)

$ErrorActionPreference = "Stop"

Write-Host "Configurando projeto GCP: $ProjectId" -ForegroundColor Cyan
gcloud config set project $ProjectId

gcloud services enable bigquery.googleapis.com iam.googleapis.com

Write-Host "Criando dataset BigQuery, se necessário..." -ForegroundColor Cyan
bq --location=$Location mk -d --description "Checklist NR-1 PS&S Advogados" $DatasetId 2>$null
if ($LASTEXITCODE -ne 0) { Write-Host "Dataset pode já existir. Continuando..." -ForegroundColor Yellow }

$schemaTemp = Join-Path $env:TEMP "pss_nr1_schema.sql"
(Get-Content "sql/schema.sql" -Raw).Replace("SEU_PROJECT_ID", $ProjectId).Replace("pss_nr1", $DatasetId) | Set-Content $schemaTemp -Encoding UTF8

Write-Host "Criando tabelas..." -ForegroundColor Cyan
bq query --use_legacy_sql=false --location=$Location (Get-Content $schemaTemp -Raw)

$saEmail = "$ServiceAccountName@$ProjectId.iam.gserviceaccount.com"
Write-Host "Criando service account: $saEmail" -ForegroundColor Cyan
gcloud iam service-accounts create $ServiceAccountName --display-name "Streamlit PS&S NR1" 2>$null
if ($LASTEXITCODE -ne 0) { Write-Host "Service account pode já existir. Continuando..." -ForegroundColor Yellow }

Write-Host "Concedendo permissões mínimas para BigQuery..." -ForegroundColor Cyan
gcloud projects add-iam-policy-binding $ProjectId --member "serviceAccount:$saEmail" --role "roles/bigquery.dataEditor" | Out-Null
gcloud projects add-iam-policy-binding $ProjectId --member "serviceAccount:$saEmail" --role "roles/bigquery.jobUser" | Out-Null

New-Item -ItemType Directory -Force -Path ".secrets" | Out-Null
$keyPath = ".secrets/gcp-sa.json"
if (-not (Test-Path $keyPath)) {
  Write-Host "Gerando chave JSON local em $keyPath" -ForegroundColor Cyan
  gcloud iam service-accounts keys create $keyPath --iam-account $saEmail
} else {
  Write-Host "Chave já existe em $keyPath. Não gerei outra." -ForegroundColor Yellow
}

New-Item -ItemType Directory -Force -Path ".streamlit" | Out-Null
$json = Get-Content $keyPath -Raw | ConvertFrom-Json
$secrets = @"
[bigquery]
project_id = "$ProjectId"
dataset_id = "$DatasetId"
location = "$Location"

[gcp_service_account]
type = "$($json.type)"
project_id = "$($json.project_id)"
private_key_id = "$($json.private_key_id)"
private_key = '''$($json.private_key)'''
client_email = "$($json.client_email)"
client_id = "$($json.client_id)"
auth_uri = "$($json.auth_uri)"
token_uri = "$($json.token_uri)"
auth_provider_x509_cert_url = "$($json.auth_provider_x509_cert_url)"
client_x509_cert_url = "$($json.client_x509_cert_url)"
universe_domain = "$($json.universe_domain)"
"@
$secrets | Set-Content ".streamlit/secrets.toml" -Encoding UTF8

Write-Host "\nPronto. Copie o conteúdo de .streamlit/secrets.toml para o Secrets do Streamlit Cloud quando for publicar." -ForegroundColor Green
