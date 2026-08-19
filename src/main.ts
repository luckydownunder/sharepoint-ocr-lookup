import "./style.css";
import type { AccountInfo } from "@azure/msal-browser";
import { authConfigured, initAuth, signIn, signOut } from "./auth";
import { extractName, readDocument } from "./ocr";
import { isImage, isPdf, pdfPreviewDataUrl } from "./pdf";
import { searchSharePoint, type SharePointHit } from "./graph";

const app = document.querySelector<HTMLDivElement>("#app")!;

let account: AccountInfo | null = null;
let selectedFile: File | null = null;
let previewSrc = "";
let ocrText = "";
let extractedName = "";
let results: SharePointHit[] = [];
let status = "";
let error = "";
let busy = false;

function render() {
  app.innerHTML = `
    <header>
      <div>
        <h1>SharePoint OCR Lookup</h1>
        <p class="lede">Upload a PDF or photo. We detect the name first. Later we will upload the file to a matching SharePoint subsite.</p>
      </div>
      <div class="row">
        ${
          account
            ? `<span class="status">${escapeHtml(account.username)}</span>
               <button class="secondary" id="sign-out" type="button">Sign out</button>`
            : `<button id="sign-in" type="button" ${authConfigured ? "" : "disabled"}>Sign in with Microsoft</button>`
        }
      </div>
    </header>

    ${
      authConfigured
        ? ""
        : `<div class="config-warn">
            Copy <code>.env.example</code> to <code>.env</code> and add your Azure app Client ID and Tenant ID, then restart <code>npm run dev</code>.
            Redirect URI must be <code>${window.location.origin}</code>.
          </div>`
    }

    <section class="card">
      <div class="dropzone" id="dropzone" tabindex="0">
        Drop a PDF or image here, or click to choose a file
        <input id="file" type="file" accept="application/pdf,image/*" hidden />
      </div>
      ${
        previewSrc
          ? `<img class="preview" alt="File preview" src="${previewSrc}" />`
          : selectedFile
            ? `<p class="status" style="margin-top:1rem">${escapeHtml(selectedFile.name)}</p>`
            : ""
      }
      <div class="row" style="margin-top:1rem">
        <button id="ocr" type="button" ${selectedFile && !busy ? "" : "disabled"}>Read name</button>
      </div>
    </section>

    <section class="card">
      <p class="detected-label">Detected name</p>
      <p class="detected-name">${extractedName ? escapeHtml(extractedName) : "No name yet — upload a file to start."}</p>
      <label for="name">Correct it if needed</label>
      <div class="row">
        <input id="name" type="text" value="${escapeAttr(extractedName)}" placeholder="e.g. Jane Citizen" />
        <button id="search" type="button" ${!busy && account ? "" : "disabled"}>Search SharePoint</button>
      </div>
      ${ocrText ? `<p class="status" style="margin-top:1rem">Text we read</p><pre class="ocr">${escapeHtml(ocrText)}</pre>` : ""}
    </section>

    <section class="card">
      <p class="${error ? "error" : "status"}">${escapeHtml(error || status || "Upload a file and we will try to detect the name.")}</p>
      <div id="results">
        ${results
          .map(
            (hit) => `
          <article class="result">
            <a href="${escapeAttr(hit.url)}" target="_blank" rel="noreferrer">${escapeHtml(hit.title)}</a>
            <div class="meta">${escapeHtml(hit.source)}${hit.summary ? " · " + escapeHtml(hit.summary) : ""}</div>
          </article>
        `,
          )
          .join("")}
      </div>
    </section>
  `;

  bind();
}

function bind() {
  const dropzone = document.querySelector<HTMLDivElement>("#dropzone");
  const fileInput = document.querySelector<HTMLInputElement>("#file");
  const nameInput = document.querySelector<HTMLInputElement>("#name");

  document.querySelector("#sign-in")?.addEventListener("click", async () => {
    try {
      error = "";
      account = await signIn();
      status = "Signed in.";
    } catch (err) {
      error = toMessage(err);
    }
    render();
  });

  document.querySelector("#sign-out")?.addEventListener("click", async () => {
    await signOut();
    account = null;
    render();
  });

  dropzone?.addEventListener("click", () => fileInput?.click());
  dropzone?.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropzone.classList.add("drag");
  });
  dropzone?.addEventListener("dragleave", () => dropzone.classList.remove("drag"));
  dropzone?.addEventListener("drop", (event) => {
    event.preventDefault();
    dropzone.classList.remove("drag");
    const file = event.dataTransfer?.files[0];
    if (file) void setFile(file);
  });
  fileInput?.addEventListener("change", () => {
    const file = fileInput.files?.[0];
    if (file) void setFile(file);
  });

  nameInput?.addEventListener("input", () => {
    extractedName = nameInput.value;
    const heading = document.querySelector(".detected-name");
    if (heading) heading.textContent = extractedName || "No name yet — upload a file to start.";
  });

  document.querySelector("#ocr")?.addEventListener("click", () => void runOcr());
  document.querySelector("#search")?.addEventListener("click", () => void runSearch());
}

async function setFile(file: File) {
  if (!isPdf(file) && !isImage(file)) {
    error = "Please choose a PDF or an image file.";
    selectedFile = null;
    previewSrc = "";
    render();
    return;
  }

  selectedFile = file;
  ocrText = "";
  extractedName = "";
  results = [];
  error = "";
  previewSrc = "";
  status = `Selected ${file.name}. Reading name…`;
  render();

  try {
    if (isPdf(file)) {
      previewSrc = await pdfPreviewDataUrl(file);
    } else {
      previewSrc = await fileToDataUrl(file);
    }
  } catch {
    previewSrc = "";
  }
  render();
  await runOcr();
}

async function runOcr() {
  if (!selectedFile) return;
  busy = true;
  error = "";
  status = "Reading the file…";
  render();
  try {
    const { text, source } = await readDocument(selectedFile, (step, progress) => {
      status = `${step} (${Math.round(progress * 100)}%)`;
      const el = document.querySelector(".card .status, .error");
      if (el && !error) el.textContent = status;
    });
    ocrText = text;
    extractedName = extractName(text);
    status = extractedName
      ? `Detected name: ${extractedName}${source === "pdf-text" ? " (from PDF text)" : " (from OCR)"}.`
      : "Could not detect a name automatically. Type it in the box above.";
  } catch (err) {
    error = toMessage(err);
  } finally {
    busy = false;
    render();
  }
}

async function runSearch() {
  const name = document.querySelector<HTMLInputElement>("#name")?.value.trim() ?? "";
  extractedName = name;
  if (!name) {
    error = "Enter or detect a name first.";
    render();
    return;
  }
  busy = true;
  error = "";
  status = `Searching SharePoint for “${name}”…`;
  results = [];
  render();
  try {
    results = await searchSharePoint(name);
    status = results.length
      ? `Found ${results.length} item${results.length === 1 ? "" : "s"}.`
      : "No SharePoint matches for that name.";
  } catch (err) {
    error = toMessage(err);
  } finally {
    busy = false;
    render();
  }
}

function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function escapeHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function escapeAttr(value: string): string {
  return escapeHtml(value);
}

function toMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

initAuth()
  .then((signedIn) => {
    account = signedIn;
    if (account) status = "Signed in. Upload a PDF or image to detect a name.";
    render();
  })
  .catch((err) => {
    error = toMessage(err);
    render();
  });
