// ==================== CONFIG ====================
// These values match the backend in main.py

const API_BASE = "http://localhost:8000";
const ALLOWED_EXTENSIONS = [".pdf"];
const MAX_FILE_SIZE_MB = 20;
const REQUEST_TIMEOUT_MS = 3 * 60 * 1000; // 3 minutes

let progressInterval = null;

// ==================== MAIN ====================

async function generateProposal() {
    const fileInput = document.getElementById("rfqFile");
    const clientInput = document.getElementById("client");
    const button = document.getElementById("generateBtn");
    const statusEl = document.getElementById("status");
    const progressBar = document.getElementById("progressBar");
    const progressContainer = document.getElementById("progressContainer");

    const file = fileInput.files[0];
    const clientName = clientInput.value.trim();

    // ---- Validation ----
    if (!file) return alert("Please select an RFQ PDF");
    if (!clientName) return alert("Please enter a client name");

    const dotIndex = file.name.lastIndexOf(".");
    const ext = dotIndex >= 0 ? file.name.slice(dotIndex).toLowerCase() : "";
    if (!ALLOWED_EXTENSIONS.includes(ext)) {
        return alert("Unsupported file type. Please upload a PDF file.");
    }
    if (file.size > MAX_FILE_SIZE_MB * 1024 * 1024) {
        return alert(`File is too large (max ${MAX_FILE_SIZE_MB} MB)`);
    }

    // ---- UI: loading state ----
    if (progressContainer) progressContainer.style.display = "block";
    button.disabled = true;
    button.innerHTML = `<span class="spinner"></span> Generating...`;
    statusEl.innerText = "Generating proposal... this may take 15-30 seconds";
    startProgressBar(progressBar);

    const formData = new FormData();
    formData.append("file", file);
    formData.append("client_name", clientName);

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);

    try {
        let response;
        try {
            response = await fetch(`${API_BASE}/generate`, {
                method: "POST",
                body: formData,
                signal: controller.signal,
            });
        } catch (err) {
            if (err.name === "AbortError") {
                throw new Error("The request timed out. Please try again.");
            }
            throw new Error("Can't reach the server. Is the backend running on port 8000?");
        }

        if (!response.ok) {
            const errorData = await response.json().catch(() => ({}));
            throw new Error(
                parseErrorDetail(errorData.detail) || `Generation failed (${response.status})`
            );
        }

        const data = await response.json();

        completeProgressBar(progressBar);
        statusEl.innerText = "Proposal generated successfully. Download should start automatically.";

        triggerDownload(data.download_url, data.filename);

        // Only reset the form on success so the user can retry after an error
        fileInput.value = "";
        clientInput.value = "";
    } catch (error) {
        console.error(error);
        statusEl.innerText = `Error: ${error.message}`;
        resetProgressBar(progressBar);
    } finally {
        clearTimeout(timeoutId);
        button.disabled = false;
        button.textContent = "Generate Proposal";
    }
}

// ==================== HELPERS ====================

function parseErrorDetail(detail) {
    if (!detail) return "";
    if (typeof detail === "string") return detail;
    // FastAPI validation errors look like: [{ loc: [...], msg: "..." }]
    if (Array.isArray(detail)) return detail.map((d) => d.msg).join("; ");
    return JSON.stringify(detail);
}

function triggerDownload(url, filename) {
    // download_url is already URL-encoded by the backend.
    // The `download` attribute is ignored cross-origin, but the server's
    // Content-Disposition header makes the browser save it correctly anyway.
    const a = document.createElement("a");
    a.href = url;
    a.download = filename || "Proposal.pptx";
    document.body.appendChild(a);
    a.click();
    a.remove();
}

function startProgressBar(progressBar) {
    if (!progressBar) return;
    clearInterval(progressInterval);
    progressBar.style.width = "0%";
    let progress = 0;

    progressInterval = setInterval(() => {
        progress = Math.min(progress + Math.random() * 12, 92);
        progressBar.style.width = `${progress}%`;
    }, 280);
}

function completeProgressBar(progressBar) {
    if (!progressBar) return;
    clearInterval(progressInterval);
    progressBar.style.width = "100%";
    setTimeout(() => resetProgressBar(progressBar), 800);
}

function resetProgressBar(progressBar) {
    clearInterval(progressInterval);
    if (progressBar) progressBar.style.width = "0%";
    const container = document.getElementById("progressContainer");
    if (container) container.style.display = "none";
}