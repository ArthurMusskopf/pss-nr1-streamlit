CREATE SCHEMA IF NOT EXISTS `football-data-science.pss_nr1`
OPTIONS(
  location = "US",
  description = "Checklist NR-1 PS&S Advogados"
);

CREATE TABLE IF NOT EXISTS `football-data-science.pss_nr1.inicios` (
  start_id STRING NOT NULL,
  created_at TIMESTAMP NOT NULL,
  created_at_br STRING,
  nome STRING,
  email STRING,
  empresa STRING,
  cnpj STRING,
  cargo STRING,
  telefone STRING,
  aceite_dados BOOL,
  fonte STRING
)
PARTITION BY DATE(created_at)
CLUSTER BY empresa;

CREATE TABLE IF NOT EXISTS `football-data-science.pss_nr1.submissoes` (
  submission_id STRING NOT NULL,
  start_id STRING,
  certificado_id STRING NOT NULL,
  created_at TIMESTAMP NOT NULL,
  created_at_br STRING,
  nome STRING,
  email STRING,
  empresa STRING,
  cnpj STRING,
  cargo STRING,
  telefone STRING,
  workshop_titulo STRING,
  pontuacao FLOAT64,
  pontuacao_maxima FLOAT64,
  percentual FLOAT64,
  status STRING,
  pacote_recomendado STRING,
  respostas_json STRING,
  user_agent STRING,
  fonte STRING
)
PARTITION BY DATE(created_at)
CLUSTER BY empresa, status;

CREATE TABLE IF NOT EXISTS `football-data-science.pss_nr1.respostas` (
  submission_id STRING NOT NULL,
  certificado_id STRING NOT NULL,
  created_at TIMESTAMP NOT NULL,
  nome STRING,
  email STRING,
  empresa STRING,
  numero INT64,
  area STRING,
  requisito STRING,
  pergunta STRING,
  resposta STRING,
  pontuacao FLOAT64,
  observacao STRING,
  acao_sugerida STRING
)
PARTITION BY DATE(created_at)
CLUSTER BY area, resposta;
