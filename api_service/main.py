#!/usr/bin/env python3
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse, JSONResponse, HTMLResponse, FileResponse
import tempfile
import os
from typing import Optional
import json

import pandas as pd

from python_pipeline.diversity_analyzer import DiversityAnalyzer

app = FastAPI(title="Relatório de Diversidade (Excel)")


def _get_build_meta():
    commit = "unknown"
    build_time = "unknown"
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        candidates = [
            os.path.join(here, 'build_meta.json'),
            os.path.abspath(os.path.join(here, os.pardir, 'build_meta.json')),
        ]
        for path in candidates:
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    commit = str(data.get('commit', commit))
                    build_time = str(data.get('build_time', build_time))
                    break
    except Exception:
        pass
    return commit, build_time


@app.get("/")
def root():
    commit, build_time = _get_build_meta()
    return {
        "status": "ok",
        "service": "diversity-excel",
        "version": "1.0",
        "commit": commit,
        "build_time": build_time,
    }


@app.get("/ui", response_class=HTMLResponse)
def ui_page():
    """Serve o index.html do projeto (funciona local e no container)."""
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        candidates = [
            os.path.join(here, 'index.html'),                # no container: /app/index.html
            os.path.abspath(os.path.join(here, os.pardir, 'index.html')),  # local: repo raiz
        ]
        for path in candidates:
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    return HTMLResponse(f.read())
        return HTMLResponse(
            "<h1>Erro ao carregar UI</h1><pre>index.html não encontrado</pre>",
            status_code=500,
        )
    except Exception as e:
        return HTMLResponse(f"<h1>Erro ao carregar UI</h1><pre>{e}</pre>", status_code=500)


@app.get("/download-model")
def download_model():
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        repo_root = os.path.abspath(os.path.join(here, os.pardir))
        # 1) Caminho preferido no container: /app/dataset.csv (copiado no Dockerfile)
        direct_path = os.path.join(here, 'dataset.csv')
        if os.path.exists(direct_path):
            return FileResponse(direct_path, media_type='text/csv', filename='modelo_dados_exemplo.csv')
        # Tenta caminho conhecido primeiro
        preferred = os.path.join(repo_root, 'Diversidade', 'Caso 1 - Variáveis categóricas', 'dataset.csv')
        model_path = preferred if os.path.exists(preferred) else None
        # Se não achar (p. ex., diferenças de normalização de unicode), busca por nome do arquivo
        if not model_path:
            for root, _dirs, files in os.walk(os.path.join(repo_root, 'Diversidade')):
                if 'dataset.csv' in files:
                    model_path = os.path.join(root, 'dataset.csv')
                    break
        if not model_path or not os.path.exists(model_path):
            # Fallback: gerar um CSV de exemplo mínimo em memória
            import io
            sample_csv = io.StringIO()
            sample_csv.write("genero,raca,setor,cargo\n")
            sample_csv.write("Feminino,Branca,Vendas,Analista\n")
            sample_csv.write("Masculino,Preta,Operações,Técnico\n")
            sample_csv.write("Feminino,Parda,Financeiro,Coordenadora\n")
            sample_csv.seek(0)
            return StreamingResponse(sample_csv, media_type='text/csv', headers={
                "Content-Disposition": "attachment; filename=modelo_dados_exemplo.csv"
            })
        return FileResponse(model_path, media_type='text/csv', filename='modelo_dados_exemplo.csv')
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


@app.post("/process")
def process(
    file: UploadFile = File(...),
    company_name: Optional[str] = Form(None),
    excel_only: Optional[bool] = Form(False)
):
    try:
        # Save upload to temp csv/xlsx
        suffix = os.path.splitext(file.filename or 'input.csv')[-1].lower()
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as f:
            content = file.file.read()
            f.write(content)
            src_path = f.name

        # Output path
        out_fd, out_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(out_fd)

        # Ensure CSV input for CSV-only pipelines
        is_excel = suffix in ['.xls', '.xlsx', '.xlsm']

        if excel_only:
            # Generate Excel-only report
            from excel_only.create_excel_only_report import main as excel_only_main
            # Patch argv for excel_only script
            import sys
            argv_bak = sys.argv
            try:
                # If user sent Excel, convert to CSV first
                csv_path = src_path
                if is_excel:
                    df = pd.read_excel(src_path)
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tf:
                        df.to_csv(tf.name, index=False)
                        csv_path = tf.name
                sys.argv = ["create_excel_only_report.py", csv_path, out_path]
                excel_only_main()
            finally:
                sys.argv = argv_bak
        else:
            # Use Python pipeline analyzer
            # DiversityAnalyzer expects CSV; convert if Excel
            csv_path = src_path
            if is_excel:
                df = pd.read_excel(src_path)
                with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tf:
                    df.to_csv(tf.name, index=False)
                    csv_path = tf.name
            analyzer = DiversityAnalyzer(csv_path)
            analyzer.run_analysis(out_path)

        # Stream the Excel file
        def iterfile():
            with open(out_path, 'rb') as f:
                yield from f
        filename = os.path.basename(file.filename or 'relatorio.xlsx').rsplit('.', 1)[0] + "_diversidade.xlsx"
        return StreamingResponse(iterfile(), media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                 headers={"Content-Disposition": f"attachment; filename={filename}"})
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})


@app.post("/inject_excel_only")
def inject_excel_only(workbook: UploadFile = File(...)):
    try:
        # Save upload workbook
        suffix = os.path.splitext(workbook.filename or 'wb.xlsm')[-1].lower()
        if suffix not in ['.xlsm', '.xlsx']:
            return JSONResponse(status_code=400, content={"error": "Workbook must be .xlsm or .xlsx"})
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as f:
            content = workbook.file.read()
            f.write(content)
            wb_path = f.name

        # Inject sheets
        from excel_only.inject_excel_only_sheets import main as inject_main
        import sys
        argv_bak = sys.argv
        try:
            sys.argv = ["inject_excel_only_sheets.py", wb_path]
            inject_main()
        finally:
            sys.argv = argv_bak

        # Return modified workbook
        def iterfile():
            with open(wb_path, 'rb') as f:
                yield from f
        filename = os.path.basename(workbook.filename or 'workbook.xlsm').rsplit('.', 1)[0] + "_with_xl_sheets" + suffix
        return StreamingResponse(iterfile(), media_type='application/vnd.ms-excel.sheet.macroEnabled.12' if suffix=='.xlsm' else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                 headers={"Content-Disposition": f"attachment; filename={filename}"})
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})
