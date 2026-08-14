import JSZip from "jszip";

export interface ZipEntry {
  filename: string;
  content: string | ArrayBuffer;
}

/** Bundles all generated output files into a single downloadable zip — nothing touches disk or a server. */
export async function buildOutputZip(entries: ZipEntry[]): Promise<Blob> {
  const zip = new JSZip();
  for (const entry of entries) {
    zip.file(entry.filename, entry.content);
  }
  return zip.generateAsync({ type: "blob" });
}

export { downloadBlob } from "../../download";
