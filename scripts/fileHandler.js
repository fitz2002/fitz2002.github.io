/* ─────────────────────────────────────────
   fileHandler.js
   Handles template file upload (TXT + PDF),
   drag-and-drop, and the template preview.
───────────────────────────────────────── */

// Shared state: the raw text of the uploaded template
let templateText = '';

// ── Entry point called by the file input's onchange ──
async function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  // Show the filename pill
  const display = document.getElementById('file-name-display');
  display.style.display = 'block';
  display.textContent = `📎 ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;

  if (file.name.endsWith('.docx')) {
    await readDocx(file);
  } else if (file.name.endsWith('.txt')) {
    readText(file);
  } else {
    addLog('error', 'Unsupported file type. Please upload a .txt or .docx file.');
  }
}

// ── Plain-text files ──
function readText(file) {
  const reader = new FileReader();
  reader.onload = () => {
    templateText = reader.result;
    showTemplatePreview();
  };
  reader.onerror = () => addLog('error', 'Failed to read text file.');
  reader.readAsText(file);
}

// ── Word documents (.docx) via Mammoth.js ──
// Mammoth extracts clean text from docx, preserving paragraph and line breaks
// exactly as they appear in the document — no coordinate guesswork needed.
async function readDocx(file) {
  if (!window.mammoth) {
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js');
  }

  const arrayBuffer = await file.arrayBuffer();

  // extractRawText preserves paragraph breaks as newlines
  const result = await mammoth.extractRawText({ arrayBuffer });

  if (result.messages.length) {
    result.messages.forEach(m => addLog('warn', `docx: ${m.message}`));
  }

  templateText = result.value.replace(/\n{2,}/g, '\n');
  showTemplatePreview();
}

// ── Renders a truncated preview of the uploaded template ──
function showTemplatePreview() {
  const wrap = document.getElementById('template-preview-wrap');
  const pre  = document.getElementById('template-preview');
  const PREVIEW_LIMIT = 1500;

  pre.textContent = templateText.length > PREVIEW_LIMIT
    ? templateText.substring(0, PREVIEW_LIMIT) + '\n\n[... preview truncated ...]'
    : templateText;

  wrap.style.display = 'block';
}

// ── Lazy-loads an external script by appending a <script> tag ──
function loadScript(src) {
  return new Promise((resolve, reject) => {
    const script  = document.createElement('script');
    script.src    = src;
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

// ── Drag-and-drop wiring (runs after DOM is ready) ──
document.addEventListener('DOMContentLoaded', () => {
  const dropZone = document.getElementById('drop-zone');

  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });

  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) {
      // Sync the hidden file input so its change event still fires if needed
      document.getElementById('file-input').files = e.dataTransfer.files;
      handleFileUpload({ target: { files: [file] } });
    }
  });
});
