# Processo de Construção e Deploy — People Analytics (Diversidade)

Este documento descreve como construir e fazer deploy do serviço de Relatório de Diversidade (FastAPI) no Google Cloud Run usando o Artifact Registry multi‑região US.

## 1. Pré‑requisitos

- gcloud SDK instalado e autenticado:
  - `gcloud version`
  - `gcloud auth login`
- Projeto configurado: `gcloud config set project <PROJECT_ID>`
- APIs habilitadas (uma vez):
  - `gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com`

## 2. Estrutura relevante do projeto

- `api_service/main.py`: aplicação FastAPI (endpoints `/`, `/ui`, `/process`, `/download-model`)
- `api_service/requirements.txt`: dependências (inclui `python-multipart`)
- `api_service/Dockerfile`: imagem para runtime (copia `index.html` e `Diversidade/`)
- `index.html`: interface web servida por `/ui`
- `Diversidade/Caso 1 - Variáveis categóricas/dataset.csv`: modelo baixado por `/download-model`
- `core/`, `python_pipeline/`, `excel_only/`: lógica de análise e geração de planilhas

## 3. Build da Imagem (Artifact Registry US)

Parâmetros recomendados:

- Repositório do Artifact Registry: `diversidade-genero-raca`
- Localização do AR: `us` (multi‑região)
- Região do Cloud Run: `us-central1`
- Nome da imagem: `people-analytics-api`

Criar repositório (uma vez):

```bash
gcloud artifacts repositories create diversidade-genero-raca \
  --repository-format=docker \
  --location=us \
  --description="diversidade_genero_raca"
```

Construir e enviar a imagem (usa `api_service/Dockerfile`):

```bash
gcloud builds submit \
  --tag us-docker.pkg.dev/<PROJECT_ID>/diversidade-genero-raca/people-analytics-api:latest .
```

## 4. Deploy no Cloud Run

Definir região e publicar o serviço:

```bash
gcloud config set run/region us-central1
gcloud run deploy people-analytics-api \
  --image us-docker.pkg.dev/<PROJECT_ID>/diversidade-genero-raca/people-analytics-api:latest \
  --platform managed \
  --allow-unauthenticated
```

Observações:
- A imagem respeita a variável `PORT` injetada pelo Cloud Run (default `8080`).
- O `Dockerfile` copia `index.html` (para `/ui`), a pasta `Diversidade/` e também o `dataset.csv` diretamente para `/app/dataset.csv` (usado por `/download-model`).

## 5. Verificação pós‑deploy

Recupere a URL do serviço:

```bash
gcloud run services describe people-analytics-api --region us-central1 --format='value(status.url)'
```

Testes rápidos:

```bash
# Health
curl -s https://<SERVICE_URL>/

# UI
open https://<SERVICE_URL>/ui

# Modelo CSV
curl -I https://<SERVICE_URL>/download-model

# Processamento (CSV)
curl -X POST https://<SERVICE_URL>/process \
  -F "file=@Diversidade/Caso 1 - Variáveis categóricas/dataset.csv" \
  -o relatorio.xlsx

# Excel-only (relatório baseado em fórmulas Excel)
curl -X POST https://<SERVICE_URL>/process \
  -F "file=@Diversidade/Caso 1 - Variáveis categóricas/dataset.csv" \
  -F "excel_only=true" \
  -o relatorio_excel_only.xlsx
```

Logs e status:

```bash
gcloud run services logs read people-analytics-api --region us-central1 --limit 100
gcloud run services describe people-analytics-api --region us-central1
```

## 6. Script de deploy (opcional)

Arquivo sugerido: `scripts/deploy.sh`

```bash
#!/usr/bin/env bash
set -euo pipefail

PROJECT_ID="<PROJECT_ID>"
REGION="us-central1"
LOCATION="us"               # Artifact Registry multi-região US
REPO="diversidade_genero_raca"
IMAGE_NAME="people-analytics-api"
IMAGE_URI="${LOCATION}-docker.pkg.dev/${PROJECT_ID}/${REPO}/${IMAGE_NAME}:latest"

echo "[1/4] Configurando projeto e APIs..."
gcloud config set project "${PROJECT_ID}"
gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com

echo "[2/4] Garantindo repositório no Artifact Registry (${LOCATION})..."
gcloud artifacts repositories describe "${REPO}" --location="${LOCATION}" >/dev/null 2>&1 || \
  gcloud artifacts repositories create "${REPO}" --repository-format=docker --location="${LOCATION}" --description="diversidade_genero_raca"

echo "[3/4] Build da imagem (Cloud Build)..."
gcloud builds submit --tag "${IMAGE_URI}" -f api_service/Dockerfile .

echo "[4/4] Deploy no Cloud Run (${REGION})..."
gcloud config set run/region "${REGION}"
gcloud run deploy "${IMAGE_NAME}" --image "${IMAGE_URI}" --platform managed --allow-unauthenticated

echo "URL do serviço:"
gcloud run services describe "${IMAGE_NAME}" --region "${REGION}" --format='value(status.url)'
```

Torne executável e use:

```bash
chmod +x scripts/deploy.sh
./scripts/deploy.sh
```

## 7. Troubleshooting rápido

- 404 em `/download-model`: confirme que o caminho existe na imagem (`COPY Diversidade/ ./Diversidade/`).
- Erro de upload: garanta `python-multipart` instalado (já listado em `requirements.txt`).
- Porta incorreta: Cloud Run define `PORT`; a imagem utiliza `${PORT}` automaticamente.
