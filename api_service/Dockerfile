FROM python:3.11-slim

WORKDIR /app
COPY api_service/requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy project modules needed by the service
COPY core/ ./core/
COPY python_pipeline/ ./python_pipeline/
COPY excel_only/ ./excel_only/
COPY api_service/main.py ./main.py

EXPOSE 8080
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8080"]

