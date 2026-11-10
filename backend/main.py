
import uuid
from pathlib import Path
from urllib.parse import quote
 
import pymupdf  # PyMuPDF
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
 
from generator import generate_proposal_json
from ppt_builder import OUTPUT_DIR, build_ppt  # share one output folder
 
app = FastAPI(title="Proposal Generator API")
 
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5500",
        "http://127.0.0.1:5500",
    ],
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
 
@app.post("/generate")
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
 
 
@app.get("/download/{filename}")
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
    return {"status": "healthy"}

from fastapi.staticfiles import StaticFiles

FRONTEND_DIR = Path(__file__).resolve().parent.parent / "frontend"
app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")