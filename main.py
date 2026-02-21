import os
import io
import csv
import logging
import sys
from pathlib import Path
from tempfile import NamedTemporaryFile
from fastapi import FastAPI, HTTPException, UploadFile, File, Form, BackgroundTasks
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import numpy as np

# --- Logging Setup ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-5s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    stream=sys.stdout,
)
log = logging.getLogger("server_v2")

app = FastAPI()

# --- CORS Configuration ---
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Constants
CACHE_DIR = Path("/tmp/xlsx_cache")
# Fallback to a local temp folder if /tmp is not writable (e.g. Windows dev)
if not os.path.exists("/tmp"):
    CACHE_DIR = Path("tmp_cache")
CACHE_DIR.mkdir(parents=True, exist_ok=True)

# Master Lists (Case-Insensitive Exact Match Expected)
ALLOWED_HUBS = {
    "aligarhmyntrahub",
    "faizabadmyntrahub",
    "deoriamyntrahub",
    "jaunpurmyntrahub",
    "maumyntrahub",
    "mirzapurmyntrahub"
}

ALLOWED_DCS = {
    "alg", "ayp", "deo", "jnp", "mau", "mrz"
}

# Header Priorities
DC_HEADERS = ["dc", "source dc", "source_dc", "dc_code", "dc code"]
HUB_HEADERS = ["hubname", "hub name", "hub_name", "finalhub", "final hub"]

def cleanup_files(*file_paths: Path):
    """Background task to delete temporary files after the response is sent."""
    for path in file_paths:
        try:
            if path and path.exists():
                path.unlink()
                log.info(f"[CLEANUP] Deleted {path}")
        except Exception as e:
            log.error(f"[CLEANUP] Error deleting {path}: {e}")

@app.get("/")
async def root():
    return {"status": "ready", "message": "Stateless XLSX-to-CSV server running. Visit /test for UI."}

@app.get("/test", response_class=HTMLResponse)
async def test_page():
    return """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLSX/CSV Deterministic Filter Test Bench</title>
    <style>
        :root {
            --primary: #2563eb;
            --bg: #f8fafc;
            --card: #ffffff;
            --text: #1e293b;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: var(--bg);
            color: var(--text);
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            margin: 0;
            padding: 20px;
        }
        .container {
            background: var(--card);
            padding: 2rem;
            border-radius: 1rem;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1);
            width: 100%;
            max-width: 500px;
        }
        h1 {
            font-size: 1.5rem;
            margin-bottom: 1.5rem;
            text-align: center;
            color: var(--primary);
        }
        .field {
            margin-bottom: 1rem;
        }
        label {
            display: block;
            margin-bottom: 0.5rem;
            font-weight: 600;
        }
        input[type="text"], input[type="file"] {
            width: 100%;
            padding: 0.75rem;
            border: 1px solid #cbd5e1;
            border-radius: 0.5rem;
            box-sizing: border-box;
        }
        button {
            width: 100%;
            padding: 0.75rem;
            background-color: var(--primary);
            color: white;
            border: none;
            border-radius: 0.5rem;
            font-size: 1rem;
            font-weight: 600;
            cursor: pointer;
            transition: background 0.2s;
        }
        button:hover {
            background-color: #1d4ed8;
        }
        button:disabled {
            background-color: #94a3b8;
            cursor: not-allowed;
        }
        #status {
            margin-top: 1.5rem;
            padding: 1rem;
            border-radius: 0.5rem;
            display: none;
            white-space: pre-wrap;
            word-break: break-all;
        }
        .error { background-color: #fee2e2; color: #991b1b; border: 1px solid #f87171; }
        .success { background-color: #dcfce7; color: #166534; border: 1px solid #4ade80; }
        .info { background-color: #e0f2fe; color: #075985; border: 1px solid #7dd3fc; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Filter API Test Bench</h1>
        
        <div class="field">
            <label for="bridgeUrl">Server URL</label>
            <input type="text" id="bridgeUrl" readonly value="">
        </div>

        <div class="field">
            <label for="jobId">Job ID (Required)</label>
            <input type="text" id="jobId" value="test-job-12345" required>
        </div>

        <div class="field">
            <label for="fileInput">Select File (.xlsx, .xls, .xlsb, .csv)</label>
            <input type="file" id="fileInput" accept=".xlsx,.xls,.xlsb,.csv" required>
        </div>

        <button id="testBtn">Process File</button>

        <div id="status"></div>
    </div>

    <script>
        document.getElementById('bridgeUrl').value = window.location.origin;
        const testBtn = document.getElementById('testBtn');
        const statusDiv = document.getElementById('status');

        testBtn.addEventListener('click', async () => {
            const bridgeUrl = document.getElementById('bridgeUrl').value.trim();
            const jobId = document.getElementById('jobId').value.trim();
            const fileInput = document.getElementById('fileInput');

            if (!jobId) {
                showStatus('Job ID is required.', 'error');
                return;
            }

            if (!fileInput.files.length) {
                showStatus('Please select a file first.', 'error');
                return;
            }

            const file = fileInput.files[0];
            const formData = new FormData();
            formData.append('job_id', jobId);
            formData.append('file', file);

            const uploadUrl = `${bridgeUrl}/process`;

            statusDiv.style.display = 'none';
            testBtn.disabled = true;
            showStatus('Uploading & Processing...', 'info');

            fetch(uploadUrl, {
                method: 'POST',
                body: formData
            })
            .then(async response => {
                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `filtered_${jobId}.csv`;
                    document.body.appendChild(a);
                    a.click();
                    a.remove();
                    window.URL.revokeObjectURL(url);
                    showStatus('Success! Filtered CSV has been downloaded.', 'success');
                } else {
                    const errorJson = await response.json();
                    showStatus(`Error ${response.status}: ${JSON.stringify(errorJson)}`, 'error');
                }
            })
            .catch(error => {
                showStatus(`Request failed: ${error}`, 'error');
            })
            .finally(() => {
                testBtn.disabled = false;
            });
        });

        function showStatus(message, type) {
            statusDiv.textContent = message;
            statusDiv.className = type;
            statusDiv.style.display = 'block';
        }
    </script>
</body>
</html>
    """

@app.post("/process")
async def process_file(
    background_tasks: BackgroundTasks,
    job_id: str = Form(..., description="Unique Job ID for processing"),
    file: UploadFile = File(..., description="Excel or CSV file from Drive")
):
    log.info(f"[JOB {job_id}] Processing request for file: {file.filename}")
    
    # 1. Validate File Ext
    ext = Path(file.filename).suffix.lower()
    if ext not in [".xlsx", ".xls", ".xlsb", ".csv"]:
        raise HTTPException(status_code=415, detail=f"Unsupported file type: {ext}. Allowed: .xlsx, .xls, .xlsb, .csv")

    input_path = CACHE_DIR / f"{job_id}_input{ext}"
    output_path = CACHE_DIR / f"{job_id}_output.csv"

    # Queue Cleanup
    background_tasks.add_task(cleanup_files, input_path, output_path)

    # 2. Save stream to disk
    try:
        with open(input_path, "wb") as f:
            while chunk := await file.read(8192):
                f.write(chunk)
    except Exception as e:
        log.error(f"[JOB {job_id}] Error saving file: {e}")
        raise HTTPException(status_code=500, detail="Failed to save uploaded file.")

    # 3. Read Data & Select Sheet
    try:
        if ext == ".csv":
            df = pd.read_csv(input_path)
            if df.empty:
                raise HTTPException(status_code=422, detail="CSV file is empty.")
        else:
            # Handle Excel formats
            engine = None
            if ext == ".xls": engine = "xlrd"
            elif ext == ".xlsb": engine = "pyxlsb"
            elif ext == ".xlsx": engine = "openpyxl"
            
            excel_file = pd.ExcelFile(input_path, engine=engine)
            best_sheet = None
            max_cells = -1

            for sheet_name in excel_file.sheet_names:
                df_test = pd.read_excel(excel_file, sheet_name=sheet_name)
                # Count non-null cells
                cell_count = df_test.notna().sum().sum()
                if cell_count > max_cells:
                    max_cells = cell_count
                    best_sheet = sheet_name
                # explicit memory cleanup
                del df_test

            import gc
            gc.collect()

            if best_sheet is None or max_cells == 0:
                raise HTTPException(status_code=422, detail="No data found in any Excel sheet.")

            log.info(f"[JOB {job_id}] Selected sheet '{best_sheet}' with {max_cells} valid cells.")
            df = pd.read_excel(excel_file, sheet_name=best_sheet)
            excel_file.close()

    except HTTPException:
        raise
    except Exception as e:
        log.error(f"[JOB {job_id}] File parsing error: {e}")
        raise HTTPException(status_code=422, detail=f"Could not parse file: str{e}")

    # 4. Header Detection & Priority Resolution
    columns = list(df.columns)
    columns_lower = {str(col).lower().strip(): col for col in columns}

    dc_col_actual = None
    hub_col_actual = None

    # Search for DC Headers (High Priority)
    for dc_h in DC_HEADERS:
        if dc_h in columns_lower:
            dc_col_actual = columns_lower[dc_h]
            break

    # If no DC found, search for Hub Headers (Fallback)
    if not dc_col_actual:
        for hub_h in HUB_HEADERS:
            if hub_h in columns_lower:
                hub_col_actual = columns_lower[hub_h]
                break

    if dc_col_actual:
        log.info(f"[JOB {job_id}] Strategy: DC Priority. Found column '{dc_col_actual}'.")
        # Ensure string type, then trim and lower for matching
        df['__match_col'] = df[dc_col_actual].astype(str).str.strip().str.lower()
        df_filtered = df[df['__match_col'].isin(ALLOWED_DCS)]
        df_filtered = df_filtered.drop(columns=['__match_col'])
    elif hub_col_actual:
        log.info(f"[JOB {job_id}] Strategy: Hub Fallback. Found column '{hub_col_actual}'.")
        # Split by underscore and take the first part to handle "AligarhMYNTRAHUB_ALG" -> "aligarhmyntrahub"
        df['__match_col'] = df[hub_col_actual].astype(str).str.strip().str.lower().apply(lambda x: x.split('_')[0])
        df_filtered = df[df['__match_col'].isin(ALLOWED_HUBS)]
        df_filtered = df_filtered.drop(columns=['__match_col'])
    else:
        raise HTTPException(status_code=400, detail=f"No valid DC or Hub column found. Detected headers: {columns}")

    # 5. Export CSV
    try:
        df_filtered.to_csv(output_path, index=False, encoding='utf-8')
    except Exception as e:
        log.error(f"[JOB {job_id}] CSV generation error: {e}")
        raise HTTPException(status_code=500, detail="Failed to write CSV output.")

    # Explicit memory cleanup
    del df
    del df_filtered
    gc.collect()

    log.info(f"[JOB {job_id}] Processing complete. Ready for download.")

    # 6. Return response
    return FileResponse(
        path=output_path,
        media_type="text/csv",
        filename=f"filtered_{job_id}.csv",
        headers={"X-Job-ID": job_id}
    )

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
