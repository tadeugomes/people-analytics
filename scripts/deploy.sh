#!/usr/bin/env bash
set -euo pipefail

# === CONFIGURAÇÃO ===
# Se não definir PROJECT_ID, usa o configurado no gcloud
PROJECT_ID="${PROJECT_ID:-<PROJECT_ID>}"
if [[ "$PROJECT_ID" == "<PROJECT_ID>" || -z "$PROJECT_ID" ]]; then
  PROJECT_ID="$(gcloud config get-value project 2>/dev/null)"
fi
if [[ -z "$PROJECT_ID" ]]; then
  echo "[erro] Defina PROJECT_ID (env) ou configure via: gcloud config set project <PROJECT_ID>" >&2
  exit 1
fi

# Região do Cloud Run
REGION="us-central1"

# Localização multi-região do Artifact Registry (US)
LOCATION="us"

# Nome do repositório no Artifact Registry
REPO="diversidade-genero-raca"

# Nome da imagem/serviço
IMAGE_NAME="people-analytics-api"

# URI completo da imagem
IMAGE_URI="${LOCATION}-docker.pkg.dev/${PROJECT_ID}/${REPO}/${IMAGE_NAME}:latest"

echo "[1/4] Configurando projeto e habilitando APIs..."
gcloud config set project "${PROJECT_ID}"
gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com

echo "[2/4] Garantindo repositório no Artifact Registry (${LOCATION})..."
if ! gcloud artifacts repositories describe "${REPO}" --location="${LOCATION}" >/dev/null 2>&1; then
  gcloud artifacts repositories create "${REPO}" \
    --repository-format=docker \
    --location="${LOCATION}" \
    --description="diversidade_genero_raca"
fi

echo "[3/4] Build da imagem com Cloud Build..."
# Gera metadados de build (commit e horário) para rastreabilidade
COMMIT="$(git rev-parse --short HEAD 2>/dev/null || echo unknown)"
BUILD_TIME="$(date -u +%Y-%m-%dT%H:%M:%SZ)"
cat > build_meta.json <<EOF
{ "commit": "${COMMIT}", "build_time": "${BUILD_TIME}" }
EOF

# Usa o Dockerfile da raiz
gcloud builds submit --tag "${IMAGE_URI}" .

# Limpa artefato local
rm -f build_meta.json || true

echo "[4/4] Deploy no Cloud Run (${REGION})..."
gcloud config set run/region "${REGION}"
gcloud run deploy "${IMAGE_NAME}" \
  --image "${IMAGE_URI}" \
  --platform managed \
  --allow-unauthenticated

echo "URL do serviço:"
gcloud run services describe "${IMAGE_NAME}" --region "${REGION}" --format='value(status.url)'
