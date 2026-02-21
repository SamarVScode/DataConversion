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
import gc

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
HUB_HEADERS = ["hubname", "hub name", "hub_name", "finalhub", "final hub", "sourcehub", "source hub", "source_hub"]

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
        
        /* Progress Bar Styles */
        .progress-group {
            margin-top: 1.5rem;
            display: none;
        }
        .progress-label {
            font-size: 0.875rem;
            margin-bottom: 0.5rem;
            display: flex;
            justify-content: space-between;
        }
        .progress-bar-container {
            width: 100%;
            height: 10px;
            background: #e2e8f0;
            border-radius: 5px;
            overflow: hidden;
            margin-bottom: 1rem;
        }
        .progress-bar-fill {
            height: 100%;
            background: var(--primary);
            width: 0%;
            transition: width 0.3s;
        }
        .progress-bar-fill.indeterminate {
            width: 100%;
            background: linear-gradient(90deg, #2563eb 25%, #60a5fa 50%, #2563eb 75%);
            background-size: 200% 100%;
            animation: move-bg 1.5s infinite linear;
        }
        @keyframes move-bg {
            0% { background-position: 200% 0; }
            100% { background-position: -200% 0; }
        }
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

        <button id="testBtn">Upload & Process File</button>

        <div class="progress-group" id="progressGroup">
            <div class="progress-label">
                <span>Uploading...</span>
                <span id="uploadPct">0%</span>
            </div>
            <div class="progress-bar-container">
                <div class="progress-bar-fill" id="uploadBar"></div>
            </div>

            <div class="progress-label" id="convertLabel" style="display:none">
                <span>Processing & Filtering (Please Wait)...</span>
            </div>
            <div class="progress-bar-container" id="convertBarContainer" style="display:none">
                <div class="progress-bar-fill indeterminate"></div>
            </div>
        </div>

        <div id="status"></div>
    </div>

    <script>
        document.getElementById('bridgeUrl').value = window.location.origin;
        const testBtn = document.getElementById('testBtn');
        const statusDiv = document.getElementById('status');
        const progressGroup = document.getElementById('progressGroup');
        const uploadBar = document.getElementById('uploadBar');
        const uploadPct = document.getElementById('uploadPct');
        const convertLabel = document.getElementById('convertLabel');
        const convertBarContainer = document.getElementById('convertBarContainer');

        testBtn.addEventListener('click', () => {
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

            // Reset UI
            statusDiv.style.display = 'none';
            progressGroup.style.display = 'block';
            convertLabel.style.display = 'none';
            convertBarContainer.style.display = 'none';
            uploadBar.style.width = '0%';
            uploadPct.textContent = '0%';
            testBtn.disabled = true;

            const xhr = new XMLHttpRequest();
            xhr.open('POST', uploadUrl, true);
            xhr.responseType = 'blob'; // Expect binary blob for download success
            
            xhr.upload.onprogress = (e) => {
                if (e.lengthComputable) {
                    const percent = Math.round((e.loaded / e.total) * 100);
                    uploadBar.style.width = percent + '%';
                    uploadPct.textContent = percent + '%';
                    
                    if (percent === 100) {
                        setTimeout(() => {
                            convertLabel.style.display = 'flex';
                            convertBarContainer.style.display = 'block';
                            addLog('Upload complete. Parsing sheets...');
                        }, 200);
                    }
                }
            };
            
            xhr.onloadstart = () => {
                statusDiv.style.display = 'block';
                statusDiv.className = 'info';
                statusDiv.innerHTML = '<strong>Server Logs:</strong><br>';
                addLog('Connecting to server...');
                addLog('Upload started. Check your Python terminal for live processing logs once upload finishes!');
            };

            xhr.onload = async () => {
                progressGroup.style.display = 'none';
                if (xhr.status >= 200 && xhr.status < 300) {
                    addLog('✅ Success! CSV generated. Initiating download...');
                    const blob = xhr.response;
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
                    // Blob might be JSON error dict, read it out
                    const reader = new FileReader();
                    reader.onload = () => {
                        let errorMsg = reader.result;
                        try {
                            const parsed = JSON.parse(reader.result);
                            errorMsg = JSON.stringify(parsed);
                        } catch (e) {}
                        showStatus(`Error ${xhr.status}: ${errorMsg}`, 'error');
                    };
                    reader.readAsText(xhr.response);
                }
                testBtn.disabled = false;
            };

            xhr.onerror = () => {
                progressGroup.style.display = 'none';
                showStatus('Network Error during upload or processing.', 'error');
                testBtn.disabled = false;
            };

            xhr.send(formData);
        });

        function addLog(msg) {
            statusDiv.innerHTML += `<div>[${new Date().toLocaleTimeString()}] ${msg}</div>`;
        }

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

    # 3. Unified Streaming Processor
    class UnifiedRowProcessor:
        def __init__(self, out_path, j_id, s_name):
            self.file_handle = open(out_path, "w", encoding="utf-8", newline="")
            self.writer = csv.writer(self.file_handle)
            self.j_id = j_id
            self.s_name = s_name
            self.first_row = True
            self.strategy = None
            self.target_col_idx = None
            self.row_count = 0
            self.match_count = 0
            self.line_buffer = ""
            self.error_headers = []

        def write(self, data):
            """Special helper for xlsx2csv which emits raw CSV string chunks."""
            if isinstance(data, bytes):
                data = data.decode('utf-8', errors='ignore')
            self.line_buffer += data
            if "\n" in self.line_buffer:
                parts = self.line_buffer.split("\n")
                for line in parts[:-1]:
                    if line.strip():
                        try:
                            f_line = io.StringIO(line)
                            row = next(csv.reader(f_line))
                            self.process_row(row)
                        except: pass
                self.line_buffer = parts[-1]

        def process_row(self, row):
            """Core logic to filter a single row (list of values)."""
            self.row_count += 1
            if self.row_count % 10000 == 0:
                log.info(f"[FILTER] ⏳ {self.s_name}: {self.row_count:,} rows processed, {self.match_count:,} matches so far")
                
            if not any(row): return # Skip empty rows
            
            if self.first_row:
                self.first_row = False
                # Cleanup headers: stringify, strip, lowercase
                curr_headers = [str(h).strip() if h is not None else "" for h in row]
                col_map = {h.lower(): i for i, h in enumerate(curr_headers)}
                self.error_headers = curr_headers
                
                # Strategy Detection
                for dc_h in DC_HEADERS:
                    if dc_h in col_map:
                        self.target_col_idx = col_map[dc_h]
                        self.strategy = "dc"
                        log.info(f"[JOB {self.j_id}] Strategy: DC Priority. Found column '{dc_h}'.")
                        break
                
                if not self.strategy:
                    for hub_h in HUB_HEADERS:
                        if hub_h in col_map:
                            self.target_col_idx = col_map[hub_h]
                            self.strategy = "hub"
                            log.info(f"[JOB {self.j_id}] Strategy: Hub Fallback. Found column '{hub_h}'.")
                            break
                
                if not self.strategy:
                    self.strategy = "error"
                    raise ValueError("NO_VALID_HEADERS")
                
                self.writer.writerow(curr_headers)
                return
                
            if self.strategy == "error": return
            
            if self.target_col_idx is not None and len(row) > self.target_col_idx:
                raw_val = row[self.target_col_idx]
                val = str(raw_val).strip().lower() if raw_val is not None else ""
                
                if self.strategy == "dc":
                    if val in ALLOWED_DCS:
                        self.match_count += 1
                        self.writer.writerow(row)
                elif self.strategy == "hub":
                    val_prefix = val.split('_')[0]
                    if val_prefix in ALLOWED_HUBS:
                        self.match_count += 1
                        self.writer.writerow(row)

        def finalize(self):
            # Finalize buffer for xlsx2csv if any
            if self.line_buffer.strip():
                try:
                    f_line = io.StringIO(self.line_buffer)
                    row = next(csv.reader(f_line))
                    self.process_row(row)
                except: pass
            self.line_buffer = ""
            self.file_handle.close()

    try:
        # Format fork
        if ext == ".csv":
            log.info(f"[JOB {job_id}] Processing CSV via streaming reader...")
            best_sheet = "CSV"
            f_processor = UnifiedRowProcessor(output_path, job_id, best_sheet)
            with open(input_path, "r", encoding="utf-8", errors="ignore") as f:
                reader = csv.reader(f)
                for row in reader:
                    f_processor.process_row(row)
            f_processor.finalize()
            log.info(f"[JOB {job_id}] ✅ Complete. Kept {f_processor.match_count:,} rows.")

        elif ext == ".xlsx":
            import xlsx2csv
            log.info(f"[JOB {job_id}] Scanning XLSX structure...")
            excel_file = pd.ExcelFile(input_path, engine="openpyxl")
            sheet_names = excel_file.sheet_names
            
            # Fast-path sheet selection
            best_sheet = None
            RAW_SHEET_CANDIDATES = {"raw", "raw data", "raw_data", "row data", "row_data"}
            for name in sheet_names:
                if name.strip().lower() in RAW_SHEET_CANDIDATES:
                    best_sheet = name
                    log.info(f"[JOB {job_id}] Fast-path: Selected '{best_sheet}'.")
                    break
            
            if not best_sheet:
                log.info(f"[JOB {job_id}] Scoring sheets by row count...")
                max_rows = -1
                for name in sheet_names:
                    # Optimized: load only the index to count rows quickly
                    df_temp = pd.read_excel(excel_file, sheet_name=name, usecols=[0])
                    rows = len(df_temp)
                    if rows > max_rows:
                        max_rows = rows
                        best_sheet = name
                    del df_temp
                gc.collect()

            s_idx = sheet_names.index(best_sheet) + 1
            excel_file.close()

            f_processor = UnifiedRowProcessor(output_path, job_id, best_sheet)
            log.info(f"[JOB {job_id}] Streaming XLSX filter via xlsx2csv...")
            xlsx2csv.Xlsx2csv(str(input_path), skip_empty_lines=True).convert(f_processor, sheetid=s_idx)
            f_processor.finalize()
            log.info(f"[JOB {job_id}] ✅ Complete. Kept {f_processor.match_count:,} rows.")

        elif ext == ".xlsb":
            from pyxlsb import open_workbook
            log.info(f"[JOB {job_id}] Opening XLSB with streaming reader...")
            best_sheet = None
            with open_workbook(str(input_path)) as wb:
                sheet_names = wb.sheets
                
                # Fast-path
                RAW_SHEET_CANDIDATES = {"raw", "raw data", "raw_data", "row data", "row_data"}
                for name in sheet_names:
                    if name.strip().lower() in RAW_SHEET_CANDIDATES:
                        best_sheet = name
                        break
                
                if not best_sheet:
                    log.info(f"[JOB {job_id}] Scoring XLSB sheets by row count...")
                    max_rows = -1
                    for name in sheet_names:
                        row_count = 0
                        with wb.get_sheet(name) as sheet:
                            for _ in sheet.rows():
                                row_count += 1
                        if row_count > max_rows:
                            max_rows = row_count
                            best_sheet = name
                    log.info(f"[JOB {job_id}] Selected largest XLSB sheet: '{best_sheet}' ({max_rows:,} rows)")
                f_processor = UnifiedRowProcessor(output_path, job_id, best_sheet)
                with wb.get_sheet(best_sheet) as sheet:
                    for row in sheet.rows():
                        # row is a list of Cell objects
                        f_processor.process_row([c.v for c in row])
                f_processor.finalize()
                log.info(f"[JOB {job_id}] ✅ Complete. Kept {f_processor.match_count:,} rows.")

        elif ext == ".xls":
            import xlrd
            log.info(f"[JOB {job_id}] Opening XLS (Legacy format)...")
            wb = xlrd.open_workbook(input_path)
            sheet_names = wb.sheet_names()
            
            best_sheet = None
            RAW_SHEET_CANDIDATES = {"raw", "raw data", "raw_data", "row data", "row_data"}
            for name in sheet_names:
                if name.strip().lower() in RAW_SHEET_CANDIDATES:
                    best_sheet = name
                    break
            
            if not best_sheet:
                # Find sheet with most rows
                max_rows = -1
                for name in sheet_names:
                    s = wb.sheet_by_name(name)
                    if s.nrows > max_rows:
                        max_rows = s.nrows
                        best_sheet = name
            
            log.info(f"[JOB {job_id}] Selected XLS sheet: '{best_sheet}'")
            sheet = wb.sheet_by_name(best_sheet)
            f_processor = UnifiedRowProcessor(output_path, job_id, best_sheet)
            for i in range(sheet.nrows):
                f_processor.process_row(sheet.row_values(i))
            f_processor.finalize()
            log.info(f"[JOB {job_id}] ✅ Complete. Kept {f_processor.match_count:,} rows.")

    except ValueError as ve:
        if "NO_VALID_HEADERS" in str(ve):
            # Try to report what we found
            hdrs = f_processor.error_headers if 'f_processor' in locals() else "Unknown"
            raise HTTPException(status_code=400, detail=f"No valid DC or Hub column found. Detected headers: {hdrs}")
        raise HTTPException(status_code=422, detail=str(ve))
    except Exception as e:
        log.error(f"[JOB {job_id}] Processing error: {e}")
        raise HTTPException(status_code=500, detail=f"Internal processing failed: {str(e)}")

    log.info(f"[JOB {job_id}] Processing complete. Ready for download.")

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
