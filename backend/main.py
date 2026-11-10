import os
import secrets
import time
import uuid
from collections import defaultdict, deque
from pathlib import Path
from urllib.parse import quote
 
from dotenv import load_dotenv
 
load_dotenv()  # reads .env locally; does nothing on the host, where real env vars are set
 
import pymupdf  # PyMuPDF
from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
 
from generator import generate_proposal_json
from ppt_builder import OUTPUT_DIR, build_ppt  # share one output folder
 
app = FastAPI(title="Proposal Generator API")
 
# CORS is only needed if the frontend is served from a different origin.
# When FastAPI serves the frontend itself (the default setup), leave CORS_ORIGINS unset.
cors_origins = [o.strip() for o in os.environ.get("CORS_ORIGINS", "").split(",") if o.strip()]
if cors_origins:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=cors_origins,
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )
 
# ==================== CONFIG ====================
 
ALLOWED_EXTENSIONS = {".pdf"}
MAX_UPLOAD_MB = 20
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024
MIN_TEXT_CHARS = 50  # below this we assume extraction failed (e.g. scanned PDF)
MAX_RFQ_CHARS = 60_000  # truncate very long RFQs before sending to the model
PPTX_MEDIA_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)
 
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")
MAX_FAILED_ATTEMPTS = 5          # wrong passwords allowed...
LOCKOUT_SECONDS = 15 * 60        # ...per IP within this window
 
if not APP_PASSWORD:
    print("WARNING: APP_PASSWORD is not set. Protected endpoints will refuse all requests.")
 
 
# ==================== PASSWORD PROTECTION ====================
 
_failed_attempts = defaultdict(deque)  # ip -> timestamps of recent wrong passwords
 
 
def _client_ip(request: Request) -> str:
    # Behind Render's proxy the real client is the LAST entry the proxy appended
    forwarded = request.headers.get("x-forwarded-for")
    if forwarded:
        return forwarded.split(",")[-1].strip()
    return request.client.host if request.client else "unknown"
 
 
def require_password(request: Request, x_access_password: str = Header(default="")):
    """Dependency: rejects the request unless it carries the correct password header."""
    if not APP_PASSWORD:
        # Fail closed: never leave the app open because of a missing setting
        raise HTTPException(status_code=503, detail="Server is not configured (APP_PASSWORD is not set)")
 
    ip = _client_ip(request)
    now = time.monotonic()
    attempts = _failed_attempts[ip]
    while attempts and now - attempts[0] > LOCKOUT_SECONDS:
        attempts.popleft()
 
    if len(attempts) >= MAX_FAILED_ATTEMPTS:
        raise HTTPException(
            status_code=429,
            detail="Too many incorrect attempts. Please try again in 15 minutes.",
        )
 
    if not secrets.compare_digest(x_access_password.encode(), APP_PASSWORD.encode()):
        attempts.append(now)
        raise HTTPException(status_code=401, detail="Incorrect password")
 
    attempts.clear()
 
 
# ==================== HELPERS ====================
 
def extract_text(file_path: Path) -> str:
    """Extract text from a PDF."""
    parts = []
    with pymupdf.open(file_path) as doc:
        if doc.needs_pass:
            raise ValueError("PDF is password-protected")
        for page in doc:
            parts.append(page.get_text())
    return "\n".join(parts).strip()
 
 
def save_upload(file: UploadFile, dest: Path) -> None:
    """Stream the upload to disk, enforcing a maximum size."""
    written = 0
    with open(dest, "wb") as buffer:
        while chunk := file.file.read(1024 * 1024):
            written += len(chunk)
            if written > MAX_UPLOAD_BYTES:
                raise HTTPException(
                    status_code=413,
                    detail=f"File too large (max {MAX_UPLOAD_MB} MB)",
                )
            buffer.write(chunk)
 
 
def is_real_pdf(path: Path) -> bool:
    """Real PDFs start with the bytes %PDF-"""
    with open(path, "rb") as f:
        return f.read(5) == b"%PDF-"
 
 
def make_safe_name(name: str) -> str:
    cleaned = "".join(c if c.isalnum() or c in " -_" else "_" for c in name.strip())
    return cleaned or "Client"
 
 
# ==================== ROUTES ====================
 
@app.post("/generate", dependencies=[Depends(require_password)])
async def generate(
    file: UploadFile = File(...),
    client_name: str = Form("Client"),
):
    print(f"Received request - Client: '{client_name}', File: '{file.filename}'")
 
    if not file.filename:
        raise HTTPException(status_code=400, detail="No file uploaded")
 
    ext = Path(file.filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(status_code=400, detail="Only PDF files are supported")
 
    file_id = uuid.uuid4().hex
    safe_name = make_safe_name(client_name)
    # Short UUID suffix prevents two requests for the same client overwriting each other
    ppt_filename = f"{safe_name}_Proposal_{file_id[:8]}.pptx"
    rfq_path = OUTPUT_DIR / f"{file_id}_rfq.pdf"
 
    try:
        save_upload(file, rfq_path)
 
        if not is_real_pdf(rfq_path):
            raise HTTPException(status_code=400, detail="File is not a valid PDF")
 
        try:
            rfq_text = await run_in_threadpool(extract_text, rfq_path)
        except Exception as e:
            print(f"PDF read error: {e}")
            raise HTTPException(
                status_code=422,
                detail="Could not read this PDF. It may be corrupted or password-protected.",
            )
 
        if len(rfq_text) < MIN_TEXT_CHARS:
            raise HTTPException(
                status_code=422,
                detail=(
                    "Could not extract readable text from the PDF. "
                    "It may be a scanned (image-only) document."
                ),
            )
 
        if len(rfq_text) > MAX_RFQ_CHARS:
            print(f"RFQ text truncated from {len(rfq_text)} to {MAX_RFQ_CHARS} chars")
            rfq_text = rfq_text[:MAX_RFQ_CHARS]
 
        # These are blocking calls, so run them off the event loop
        proposal = await run_in_threadpool(generate_proposal_json, rfq_text, client_name)
 
        await run_in_threadpool(
            build_ppt,
            proposal=proposal,
            client_name=client_name,
            ppt_filename=ppt_filename,
        )
 
        return {
            "message": "Proposal generated successfully",
            "download_url": f"/download/{quote(ppt_filename)}",
            "filename": ppt_filename,
            "client_name": client_name,
        }
 
    except HTTPException:
        # Let intentional 4xx errors through instead of converting them to 500
        raise
    except Exception as e:
        print(f"Error generating proposal: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to generate proposal: {e}")
    finally:
        # Runs on success and failure
        rfq_path.unlink(missing_ok=True)
 
 
@app.get("/download/{filename}", dependencies=[Depends(require_password)])
async def download(filename: str):
    ppt_path = (OUTPUT_DIR / filename).resolve()
 
    # Block path traversal: file must sit directly inside OUTPUT_DIR
    if ppt_path.parent != OUTPUT_DIR.resolve() or ppt_path.suffix.lower() != ".pptx":
        raise HTTPException(status_code=404, detail="Proposal file not found")
 
    if not ppt_path.is_file():
        raise HTTPException(
            status_code=404, detail="Proposal file not found or has expired"
        )
 
    return FileResponse(
        path=ppt_path,
        filename=ppt_path.name,
        media_type=PPTX_MEDIA_TYPE,
    )
 
 
@app.get("/health")
async def health():
    # Left public so the host's health checks work; reveals nothing sensitive
    return {"status": "healthy"}
 
 
# ==================== FRONTEND ====================
 
FRONTEND_DIR = Path(__file__).resolve().parent.parent / "frontend"
if FRONTEND_DIR.is_dir():
    app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")
else:
    print(f"WARNING: frontend folder not found at {FRONTEND_DIR}; only the API is available.")
 