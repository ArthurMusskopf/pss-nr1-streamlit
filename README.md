# PS&S NR-1 Streamlit MVP

Aplicativo Streamlit para checklist de adequação à NR-1 com identidade visual PS&S, salvamento no BigQuery, geração de certificado PDF e matriz de ações em Excel.

## Fluxo

1. Participante preenche Identificação da Empresa e aceita o processamento de dados.
2. O app salva o pré-cadastro na tabela `inicios` do BigQuery quando os secrets estiverem configurados. Sem secrets, salva localmente em `.local_starts`.
3. O questionário é liberado.
4. Ao final, o app calcula score, status e pacote recomendado.
5. Salva a submissão em `submissoes` e as respostas em `respostas`.
6. Disponibiliza download do certificado PDF e da Matriz de Ações em Excel.

## Rodar localmente

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\run_local.ps1
```

## BigQuery

```powershell
$PROJECT_ID = "basedeempresas2025"
.\scripts\gcp_setup.ps1 -ProjectId $PROJECT_ID -DatasetId "pss_nr1" -Location "US"
```

O script cria/atualiza:

- `pss_nr1.inicios`
- `pss_nr1.submissoes`
- `pss_nr1.respostas`

Depois copie o conteúdo de `.streamlit/secrets.toml` para os Secrets do Streamlit Cloud.

## Publicar no GitHub

```powershell
gh auth login
.\scripts\git_publish.ps1 -RepoName "pss-nr1-streamlit"
```

## Customização visual

Substitua os arquivos em `assets/` se precisar atualizar a marca:

- `logo_symbol_white.png`
- `logo_wordmark_white.png`
- `logo_symbol_graphite.png`
- `logo_wordmark_graphite.png`

As cores principais estão no dicionário `BRAND` no arquivo `app.py`.
