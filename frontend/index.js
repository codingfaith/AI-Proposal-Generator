// ==================== CONFIG ====================
// These values match the backend in main.py

const API_BASE = "";   // same origin as the page (FastAPI serves both)
const ALLOWED_EXTENSIONS = [".pdf"];
const MAX_FILE_SIZE_MB = 20;
const REQUEST_TIMEOUT_MS = 3 * 60 * 1000; // 3 minutes
const PASSWORD_HEADER = "X-Access-Password";

let progressInterval = null;

// Remember the password for this browser tab only (cleared when the tab closes)
function loadSavedPassword() {
    try { return sessionStorage.getItem("accessPassword") || ""; } catch { return ""; }
}
function savePassword(value) {
    try { sessionStorage.setItem("accessPassword", value); } catch { /* ignore */ }
}
function clearSavedPassword() {
    try { sessionStorage.removeItem("accessPassword"); } catch { /* ignore */ }
}

document.addEventListener("DOMContentLoaded", () => {
    const passwordInput = document.getElementById("accessPassword");
    if (passwordInput) passwordInput.value = loadSavedPassword();
});

// ==================== MAIN ====================

async function generateProposal() {
    const fileInput = document.getElementById("rfqFile");
    const clientInput = document.getElementById("client");
    const passwordInput = document.getElementById("accessPassword");
    const button = document.getElementById("generateBtn");
    const statusEl = document.getElementById("status");
    const progressBar = document.getElementById("progressBar");
    const progressContainer = document.getElementById("progressContainer");

    const file = fileInput.files[0];
    const clientName = clientInput.value.trim();
    const password = passwordInput ? passwordInput.value : "";

    // ---- Validation ----
    if (!password) return alert("Please enter the access password");
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
                headers: { [PASSWORD_HEADER]: password },
                body: formData,
                signal: controller.signal,
            });
        } catch (err) {
            if (err.name === "AbortError") {
                throw new Error("The request timed out. Please try again.");
            }
            throw new Error("Can't reach the server. If it was idle, it may be waking up, so please try again in a minute.");
        }

        if (response.status === 401) {
            clearSavedPassword();
        }
        if (!response.ok) {
            const errorData = await response.json().catch(() => ({}));
            throw new Error(
                parseErrorDetail(errorData.detail) || `Generation failed (${response.status})`
            );
        }

        savePassword(password); // correct password: remember it for this tab

        const data = await response.json();

        completeProgressBar(progressBar);
        statusEl.innerText = "Proposal generated successfully. Download should start automatically.";

        await downloadFile(data.download_url, data.filename, password);

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

// Downloads are password-protected too, and a plain link can't send a header,
// so we fetch the file ourselves and save it from memory.
async function downloadFile(url, filename, password) {
    const response = await fetch(`${API_BASE}${url}`, {
        headers: { [PASSWORD_HEADER]: password },
    });
    if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(parseErrorDetail(errorData.detail) || `Download failed (${response.status})`);
    }

    const blob = await response.blob();
    const objectUrl = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = objectUrl;
    a.download = filename || "Proposal.pptx";
    document.body.appendChild(a);
    a.click();
    a.remove();

    setTimeout(() => URL.revokeObjectURL(objectUrl), 10000);
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