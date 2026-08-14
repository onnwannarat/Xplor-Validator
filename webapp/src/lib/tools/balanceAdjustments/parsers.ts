import Papa from "papaparse";
import ExcelJS from "exceljs";

export type Row = Record<string, string>;

function safeStr(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "object" && value !== null && "result" in value) {
    return safeStr((value as { result: unknown }).result);
  }
  return String(value);
}

function isHtmlBytes(bytes: ArrayBuffer): boolean {
  const head = new TextDecoder("utf-8", { fatal: false }).decode(bytes.slice(0, 512));
  return head.includes("<") && head.includes(">");
}

/** Parses the first HTML <table> in a string into header + row arrays, mirroring pandas.read_html()[0]. */
function parseHtmlTable(html: string): Row[] {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const table = doc.querySelector("table");
  if (!table) throw new Error("No <table> found in HTML content.");

  const trs = [...table.querySelectorAll("tr")];
  if (trs.length === 0) return [];

  const cellText = (cell: Element) => (cell.textContent ?? "").trim();
  const firstRowCells = [...trs[0].children];
  const firstRowIsHeader = firstRowCells.length > 0 && firstRowCells.every((c) => c.tagName.toLowerCase() === "th");

  const headerCells = firstRowIsHeader ? firstRowCells : [...trs[0].children];
  const headers = headerCells.map((c, idx) => {
    const text = cellText(c);
    return text || `Column${idx + 1}`;
  });

  const dataTrs = firstRowIsHeader ? trs.slice(1) : trs.slice(1);
  return dataTrs.map((tr) => {
    const cells = [...tr.children].map(cellText);
    const row: Row = {};
    headers.forEach((h, idx) => {
      row[h] = cells[idx] ?? "";
    });
    return row;
  });
}

function parseCsvText(text: string): Row[] {
  const stripped = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  const result = Papa.parse<Row>(stripped, { header: true, skipEmptyLines: true });
  return result.data.map((row) => {
    const out: Row = {};
    for (const [k, v] of Object.entries(row)) out[k.trim()] = safeStr(v);
    return out;
  });
}

async function parseXlsxBuffer(bytes: ArrayBuffer): Promise<Row[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  const ws = wb.worksheets[0];
  if (!ws) return [];

  const headerRow = ws.getRow(1);
  const headers: string[] = [];
  headerRow.eachCell({ includeEmpty: false }, (cell) => headers.push(safeStr(cell.value).trim()));

  const rows: Row[] = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const excelRow = ws.getRow(r);
    if (excelRow.cellCount === 0) continue;
    const row: Row = {};
    let hasValue = false;
    headers.forEach((h, idx) => {
      const v = safeStr(excelRow.getCell(idx + 1).value);
      if (v) hasValue = true;
      row[h] = v;
    });
    if (hasValue) rows.push(row);
  }
  return rows;
}

/**
 * Reads a balance-adjustment input file (CSV, real XLSX, or HTML-disguised-as-.xls)
 * into normalised string rows. Port of read_input_bytes.
 */
export async function readBalanceAdjustmentFile(file: File): Promise<Row[]> {
  const bytes = await file.arrayBuffer();
  const ext = file.name.toLowerCase().split(".").pop();

  if (ext === "csv") {
    return parseCsvText(new TextDecoder("utf-8").decode(bytes));
  }
  if ((ext === "xlsx" || ext === "xls") && isHtmlBytes(bytes)) {
    return parseHtmlTable(new TextDecoder("utf-8", { fatal: false }).decode(bytes));
  }
  return parseXlsxBuffer(bytes);
}

/** Converts a currency-formatted string like '$1,234.56' or '-$51.67' to a number, or null if empty/zero/unparseable. Port of parse_amount. */
export function parseAmount(val: string | undefined | null): number | null {
  if (val === null || val === undefined) return null;
  const s = String(val).trim();
  if (s === "" || s === "nan" || s === "None") return null;
  const cleaned = s.replace(/[,$]/g, "");
  const num = Number(cleaned);
  if (!Number.isFinite(num) || cleaned === "") return null;
  return num !== 0 ? num : null;
}

/** Splits "Last Name, First Name" into [firstName, lastName]. Port of parse_name. */
export function parseName(accountName: string | undefined | null): [string, string] {
  if (!accountName || String(accountName).trim() === "") return ["", ""];
  const name = String(accountName).trim();
  const commaIdx = name.indexOf(",");
  if (commaIdx === -1) return [name, ""];
  const last = name.slice(0, commaIdx).trim();
  const first = name.slice(commaIdx + 1).trim();
  return [first, last];
}
