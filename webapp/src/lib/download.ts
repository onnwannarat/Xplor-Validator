/** Triggers a browser download for an in-memory blob without any network round-trip. */
export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/** Downloads a plain-text file (e.g. a CSV template) generated entirely in-browser. */
export function downloadTextFile(content: string, filename: string, mimeType = "text/csv"): void {
  downloadBlob(new Blob([content], { type: mimeType }), filename);
}
