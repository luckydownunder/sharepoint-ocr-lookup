import { createWorker } from "tesseract.js";
import { extractPdfTextLayer, isPdf, renderPdfPages } from "./pdf";

export async function recognizeImage(
  image: File | Blob | HTMLCanvasElement,
  onProgress?: (status: string, progress: number) => void,
): Promise<string> {
  const worker = await createWorker("eng", undefined, {
    logger: (message) => {
      if (message.status && typeof message.progress === "number") {
        onProgress?.(message.status, message.progress);
      }
    },
  });
  try {
    const { data } = await worker.recognize(image);
    return data.text.trim();
  } finally {
    await worker.terminate();
  }
}

export async function readDocument(
  file: File,
  onProgress?: (status: string, progress: number) => void,
): Promise<{ text: string; source: "pdf-text" | "ocr" }> {
  if (isPdf(file)) {
    onProgress?.("Reading PDF text", 0.15);
    const embedded = await extractPdfTextLayer(file);
    if (embedded.replace(/\s/g, "").length >= 20) {
      return { text: embedded, source: "pdf-text" };
    }
    onProgress?.("PDF looks scanned — running OCR", 0.25);
    const pages = await renderPdfPages(file, 2, 2);
    const chunks: string[] = [];
    for (let i = 0; i < pages.length; i += 1) {
      const pageText = await recognizeImage(pages[i], (status, progress) => {
        const overall = 0.25 + ((i + progress) / pages.length) * 0.75;
        onProgress?.(`OCR page ${i + 1}/${pages.length}: ${status}`, overall);
      });
      chunks.push(pageText);
    }
    return { text: chunks.join("\n\n"), source: "ocr" };
  }

  const text = await recognizeImage(file, onProgress);
  return { text, source: "ocr" };
}

const SKIP = new Set([
  "THE",
  "AND",
  "FOR",
  "NAME",
  "DATE",
  "FORM",
  "PAGE",
  "SHAREPOINT",
  "LOOKUP",
  "RESIDENT",
  "CLIENT",
  "PATIENT",
]);

/** Prefer a labeled Name field; otherwise a 2–4 word proper-name-like phrase. */
export function extractName(ocrText: string): string {
  const labeled = ocrText.match(
    /(?:resident|client|patient|full)?\s*name\s*[:\-]\s*([A-Za-z][A-Za-z'’\-]+(?:\s+[A-Za-z][A-Za-z'’\-]+){0,3})/i,
  );
  if (labeled?.[1]) {
    return cleanName(labeled[1]);
  }

  const lines = ocrText
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  for (const line of lines) {
    const mixed = line.match(/\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3})\b/);
    if (mixed) return cleanName(mixed[1]);
    const caps = line.match(/\b([A-Z]{2,}(?:\s+[A-Z]{2,}){1,3})\b/);
    if (caps && !SKIP.has(caps[1].split(/\s+/)[0] ?? "")) {
      return cleanName(caps[1]);
    }
  }

  return "";
}

function cleanName(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}
