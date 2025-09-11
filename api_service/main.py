#!/usr/bin/env python3
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse, JSONResponse, HTMLResponse, FileResponse
import tempfile
import os
from typing import Optional

import pandas as pd

from python_pipeline.diversity_analyzer import DiversityAnalyzer

app = FastAPI(title="Relatório de Diversidade (Excel)")


@app.get("/")
def root():
    return {"status": "ok", "service": "diversity-excel", "version": "1.0"}


@app.get("/ui", response_class=HTMLResponse)
def ui_page():
    # Serve the repository root index.html to be the single UI
    try:
        repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
        index_path = os.path.join(repo_root, 'index.html')
        with open(index_path, 'r', encoding='utf-8') as f:
            return HTMLResponse(f.read())
    except Exception as e:
        return HTMLResponse(f"<h1>Erro ao carregar UI</h1><pre>{e}</pre>", status_code=500)


@app.get("/download-model")
def download_model():
    try:
        repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
        model_path = os.path.join(repo_root, 'Diversidade', 'Caso 1 - Variáveis categóricas', 'dataset.csv')
        if not os.path.exists(model_path):
            return JSONResponse(status_code=404, content={"error": "Modelo não encontrado"})
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
