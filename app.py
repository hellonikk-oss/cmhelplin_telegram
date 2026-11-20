import os
import sys
import base64
import datetime
import threading
from typing import Optional
from pydantic import BaseModel
from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv

# Ensure project directories are in sys.path
project_root = os.path.dirname(os.path.abspath(__file__))
code_dir = os.path.join(project_root, "cmhelpline_code")
if project_root not in sys.path:
    sys.path.insert(0, project_root)
if code_dir not in sys.path:
    sys.path.insert(0, code_dir)

load_dotenv(os.path.join(project_root, ".env"))

# Import core portal automation functions
try:
    from cmhelpline_code.storage_state_cmhelpline import save_cmhelpline_storage_state
    from cmhelpline_code.cmhlpline_all_complain_download import (
        is_session_valid,
        ensure_valid_session,
        download_all_cmhelpline_excel,
    )
    from cmhelpline_code.copy_datatoExcel import (
        cmhelpline_copy_data_L1,
        download_Final_cmhelpline,
    )
except ImportError:
    from storage_state_cmhelpline import save_cmhelpline_storage_state
    from cmhlpline_all_complain_download import (
        is_session_valid,
        ensure_valid_session,
        download_all_cmhelpline_excel,
    )
    from copy_datatoExcel import (
        cmhelpline_copy_data_L1,
        download_Final_cmhelpline,
    )

app = FastAPI(
    title="CM Helpline Complaints Downloader",
    description="FastAPI Backend to download L1, L2, L3, and L4 complaint Excel files from CM Helpline portal with auto-session management.",
    version="1.0.0"
)

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Global paths
STATE_FILE = os.path.join(code_dir, "state_CM.json")
OUTPUT_DIR = os.path.join(project_root, "cmhelpline")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Global State & Thread Synchronization for Captcha & Background tasks
task_lock = threading.Lock()
captcha_event = threading.Event()
captcha_input_store = {"value": ""}

download_state = {
    "status": "idle",       # "idle", "running", "waiting_captcha", "completed", "error"
    "stage": "idle",
    "message": "Ready to download",
    "percent": 0,
    "captcha_image": None,  # data:image/png;base64,...
    "logs": [],
    "error": None,
    "last_run": None,
}


def add_log(message: str, log_type: str = "info"):
    """Appends a timestamped log entry to download_state."""
    now_str = datetime.datetime.now().strftime("%H:%M:%S")
    entry = {"time": now_str, "message": message, "type": log_type}
    download_state["logs"].append(entry)
    # Keep last 150 logs
    if len(download_state["logs"]) > 150:
        download_state["logs"].pop(0)


def progress_callback(stage: str, message: str, percent: int):
    """Callback invoked by download routines to report live status."""
    download_state["stage"] = stage
    download_state["message"] = message
    download_state["percent"] = percent
    add_log(message, "info")


def web_captcha_callback(captcha_path: str) -> str:
    """
    Called when save_cmhelpline_storage_state() needs a CAPTCHA.
    Encodes the captcha image into base64, updates download_state,
    and blocks until the user submits the CAPTCHA from the HTML frontend.
    """
    add_log("🔐 Portal CAPTCHA required for authentication.", "warning")
    if os.path.exists(captcha_path):
        with open(captcha_path, "rb") as f:
            b64_data = base64.b64encode(f.read()).decode("utf-8")
        download_state["captcha_image"] = f"data:image/png;base64,{b64_data}"
    else:
        download_state["captcha_image"] = None

    download_state["status"] = "waiting_captcha"
    download_state["message"] = "CAPTCHA required. Please enter CAPTCHA in the web UI."

    # Clear previous event and wait for web submission
    captcha_event.clear()
    captcha_input_store["value"] = ""

    # Wait up to 180 seconds for user input
    got_signal = captcha_event.wait(timeout=180)
    if not got_signal or not captcha_input_store["value"]:
        download_state["status"] = "error"
        download_state["error"] = "Timed out waiting for CAPTCHA input."
        add_log("Timed out waiting for CAPTCHA input from web UI.", "error")
        raise TimeoutError("CAPTCHA input timed out.")

    entered_captcha = captcha_input_store["value"]
    add_log(f"Submitting CAPTCHA '{entered_captcha}' to portal...", "info")
    download_state["status"] = "running"
    download_state["captcha_image"] = None
    return entered_captcha


class DownloadRequest(BaseModel):
    consolidate: bool = True
    headless: bool = True


class CaptchaSubmitRequest(BaseModel):
    captcha: str


# -------------------------------------------------------------
# Routes
# -------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
def get_frontend():
    """Serves the main HTML dashboard."""
    html_path = os.path.join(project_root, "templates", "index.html")
    if os.path.exists(html_path):
        with open(html_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse("<h1>Templates/index.html not found</h1>", status_code=404)


@app.get("/api/session/status")
def get_session_status():
    """
    Checks if session is valid and active.
    Returns status, configured username, and state file presence.
    """
    username = os.getenv("CMHELIPLINE_USENAME") or os.getenv("CMHELPLINE_USERNAME") or "Not configured"
    file_exists = os.path.exists(STATE_FILE)
    valid = is_session_valid(STATE_FILE) if file_exists else False

    return {
        "valid": valid,
        "exists": file_exists,
        "username": username,
        "state_file": STATE_FILE,
        "message": "Session is active and valid." if valid else "Session is expired or missing."
    }


@app.post("/api/captcha/submit")
def submit_captcha(req: CaptchaSubmitRequest):
    """
    Receives CAPTCHA input from the frontend and signals the waiting
    login process to continue.
    """
    captcha_clean = req.captcha.strip()
    if not captcha_clean:
        raise HTTPException(status_code=400, detail="CAPTCHA cannot be empty.")

    captcha_input_store["value"] = captcha_clean
    captcha_event.set()
    return {"status": "ok", "message": "CAPTCHA received, processing authentication..."}


@app.post("/api/captcha/request")
def request_fresh_captcha():
    """
    Manually triggers login flow in background to capture a fresh CAPTCHA image
    for re-authentication.
    """
    if download_state["status"] == "running":
        raise HTTPException(status_code=400, detail="A download or login task is already running.")

    def run_login_worker():
        with task_lock:
            download_state["status"] = "running"
            download_state["message"] = "Opening login page to load fresh CAPTCHA..."
            add_log("Starting portal login worker to load CAPTCHA...", "info")
            try:
                success = save_cmhelpline_storage_state(
                    state_file=STATE_FILE,
                    headless=True,
                    captcha_callback=web_captcha_callback
                )
                if success:
                    download_state["status"] = "completed"
                    download_state["message"] = "Logged in successfully! Session state saved."
                    add_log("✓ Successfully logged in and saved session state.", "success")
                else:
                    download_state["status"] = "error"
                    download_state["message"] = "Login failed. Please verify credentials or CAPTCHA."
                    add_log("✗ Login failed.", "error")
            except Exception as e:
                download_state["status"] = "error"
                download_state["error"] = str(e)
                add_log(f"Login error: {e}", "error")

    t = threading.Thread(target=run_login_worker, daemon=True)
    t.start()

    # Wait shortly for captcha to become available
    for _ in range(40):
        if download_state["status"] == "waiting_captcha" and download_state["captcha_image"]:
            return {
                "status": "waiting_captcha",
                "captcha_image": download_state["captcha_image"],
                "message": "CAPTCHA loaded. Please submit characters."
            }
        threading.Event().wait(0.25)

    return {
        "status": download_state["status"],
        "captcha_image": download_state.get("captcha_image"),
        "message": download_state.get("message", "Loading CAPTCHA...")
    }


def download_worker(consolidate: bool = True, headless: bool = True):
    """
    Background worker that runs the full download process:
    1. Checks if session is valid. If not, automatically calls save_cmhelpline_storage_state()
       via ensure_valid_session().
    2. Downloads L1, L2, L3, and L4 complaint files from CM Helpline portal.
    3. Optionally runs Excel consolidation (L-1.xlsx & high_level.xlsx).
    """
    with task_lock:
        download_state["status"] = "running"
        download_state["percent"] = 5
        download_state["message"] = "Starting CM Helpline complaints download..."
        download_state["error"] = None
        add_log("==========================================", "info")
        add_log("Starting download process for L1, L2, L3, L4", "info")

        try:
            # 1. Download complaints (with auto session verification & login)
            success = download_all_cmhelpline_excel(
                headless=headless,
                state_file=STATE_FILE,
                output_dir=OUTPUT_DIR,
                progress_callback=progress_callback,
                captcha_callback=web_captcha_callback
            )

            if not success:
                download_state["status"] = "error"
                download_state["message"] = "Download failed. Please check logs."
                add_log("Download process terminated with error.", "error")
                return

            # 2. Consolidate Excel workbooks if requested
            if consolidate:
                progress_callback("consolidating", "Updating consolidated Excel workbooks (L-1.xlsx & high_level.xlsx)...", 96)
                try:
                    cmhelpline_copy_data_L1()
                    add_log("✓ Updated L-1.xlsx successfully.", "success")
                except Exception as ex:
                    add_log(f"⚠️ Note on L1 workbook update: {ex}", "warning")

                try:
                    download_Final_cmhelpline()
                    add_log("✓ Updated high_level.xlsx successfully.", "success")
                except Exception as ex:
                    add_log(f"⚠️ Note on high_level workbook update: {ex}", "warning")

            download_state["status"] = "completed"
            download_state["percent"] = 100
            download_state["message"] = "All L1, L2, L3, and L4 complaints downloaded and processed successfully!"
            download_state["last_run"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            add_log("✓ All tasks completed successfully!", "success")

        except Exception as e:
            download_state["status"] = "error"
            download_state["error"] = str(e)
            download_state["message"] = f"Error during download: {e}"
            add_log(f"Exception during download: {e}", "error")


@app.post("/api/download")
def start_download(req: DownloadRequest, background_tasks: BackgroundTasks):
    """
    Endpoint to trigger downloading L1, L2, L3, L4 complaints.
    Runs asynchronously in the background.
    """
    if download_state["status"] == "running":
        raise HTTPException(status_code=400, detail="A download task is already in progress.")

    # Reset state
    download_state["percent"] = 0
    download_state["captcha_image"] = None
    download_state["error"] = None

    t = threading.Thread(target=download_worker, args=(req.consolidate, req.headless), daemon=True)
    t.start()

    return {
        "status": "started",
        "message": "Complaints download initiated in background.",
        "consolidate": req.consolidate,
        "headless": req.headless
    }


@app.get("/api/download/status")
def get_download_status():
    """Returns the current download progress, logs, and status."""
    return download_state


@app.get("/api/files")
def list_downloaded_files():
    """Lists all Excel files in the cmhelpline directory."""
    files_info = []
    if not os.path.exists(OUTPUT_DIR):
        return []

    for f in sorted(os.listdir(OUTPUT_DIR)):
        if f.lower().endswith((".xls", ".xlsx")) and not f.startswith("~$"):
            fp = os.path.join(OUTPUT_DIR, f)
            size = os.path.getsize(fp)
            mtime = datetime.datetime.fromtimestamp(os.path.getmtime(fp)).strftime("%Y-%m-%d %H:%M:%S")

            # Format file size
            if size < 1024:
                size_str = f"{size} B"
            elif size < 1024 * 1024:
                size_str = f"{size / 1024:.1f} KB"
            else:
                size_str = f"{size / (1024 * 1024):.2f} MB"

            files_info.append({
                "name": f,
                "size_bytes": size,
                "size_formatted": size_str,
                "modified": mtime
            })

    return files_info


@app.get("/api/files/download/{filename}")
def download_file_by_name(filename: str):
    """Downloads a specific Excel file from the cmhelpline directory."""
    # Prevent directory traversal
    clean_name = os.path.basename(filename)
    fp = os.path.join(OUTPUT_DIR, clean_name)
    if not os.path.exists(fp):
        raise HTTPException(status_code=404, detail=f"File '{clean_name}' not found.")
    return FileResponse(
        path=fp,
        filename=clean_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.post("/api/process-excel")
def trigger_excel_processing():
    """Manually triggers the consolidation of L1 and high-level Excel files."""
    try:
        cmhelpline_copy_data_L1()
        download_Final_cmhelpline()
        return {"status": "ok", "message": "Successfully updated L-1.xlsx and high_level.xlsx."}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error consolidating Excel files: {e}")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="127.0.0.1", port=8000, reload=True)
