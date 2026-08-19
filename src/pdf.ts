import * as pdfjs from "pdfjs-dist";
import pdfWorker from "pdfjs-dist/build/pdf.worker.min.mjs?url";

pdfjs.GlobalWorkerOptions.workerSrc = pdfWorker;

export function isPdf(file: File): boolean {
  return file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf");
}

export function isImage(file: File): boolean {
  return file.type.startsWith("image/");
}

async function loadDocument(file: File) {
  const data = await file.arrayBuffer();
  return pdfjs.getDocument({ data }).promise;
}

export async function extractPdfTextLayer(file: File, maxPages = 3): Promise<string> {
  const doc = await loadDocument(file);
  const parts: string[] = [];
  const last = Math.min(doc.numPages, maxPages);
  for (let pageNum = 1; pageNum <= last; pageNum += 1) {
    const page = await doc.getPage(pageNum);
    const content = await page.getTextContent();
    const line = content.items
      .map((item) => ("str" in item ? item.str : ""))
      .join(" ")
      .replace(/\s+/g, " ")
      .trim();
    if (line) parts.push(line);
  }
  return parts.join("\n");
}

export async function renderPdfPage(file: File, pageNum = 1, scale = 2): Promise<HTMLCanvasElement> {
  const doc = await loadDocument(file);
  const page = await doc.getPage(Math.min(pageNum, doc.numPages));
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");
  if (!context) {
    throw new Error("Could not create a canvas to preview the PDF.");
  }
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  await page.render({ canvasContext: context, viewport, canvas }).promise;
  return canvas;
}

export async function renderPdfPages(
  file: File,
  maxPages = 2,
  scale = 2,
): Promise<HTMLCanvasElement[]> {
  const doc = await loadDocument(file);
  const last = Math.min(doc.numPages, maxPages);
  const pages: HTMLCanvasElement[] = [];
  for (let pageNum = 1; pageNum <= last; pageNum += 1) {
    pages.push(await renderPdfPage(file, pageNum, scale));
  }
  return pages;
}

export async function pdfPreviewDataUrl(file: File): Promise<string> {
  const canvas = await renderPdfPage(file, 1, 1.4);
  return canvas.toDataURL("image/png");
}
