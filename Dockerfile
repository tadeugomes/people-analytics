FROM python:3.11-slim

WORKDIR /app

COPY api_service/requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy project modules needed by the service
COPY core/ ./core/
COPY python_pipeline/ ./python_pipeline/
COPY excel_only/ ./excel_only/
COPY api_service/main.py ./main.py

# Copy UI and sample dataset used by endpoints (/ui and /download-model)
COPY index.html ./
COPY Diversidade/ ./Diversidade/
COPY build_meta.json ./
COPY ["Diversidade/Caso 1 - Variáveis categóricas/dataset.csv", "./dataset.csv"]

EXPOSE 8080
# Honor Cloud Run PORT env var if present
ENV PORT=8080
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT}"]
