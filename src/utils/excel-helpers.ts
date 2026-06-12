import * as XLSX from "xlsx";

export function normalizeHeader(value: unknown): string {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

export function findColumnIndex(
  headers: string[],
  matchers: string[]
): number {
  const normalized = headers.map(normalizeHeader);
  for (const matcher of matchers) {
    const idx = normalized.findIndex((h) => h.includes(matcher));
    if (idx >= 0) return idx;
  }
  return -1;
}

/** Lee .xls y .xlsx usando SheetJS */
export function bufferToRows(buffer: Buffer): unknown[][] {
  let workbook: XLSX.WorkBook;
  try {
    workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
  } catch {
    throw new Error(
      "No se pudo leer el archivo Excel. Verifica que sea .xls o .xlsx válido."
    );
  }

  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("El archivo Excel no contiene hojas de cálculo.");
  }

  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: null,
    raw: false,
  });

  return rows as unknown[][];
}

export function detectHeaderRow(rows: unknown[][], maxScan = 50): number {
  // Prioridad: fila que contenga "haber" (extractos bancarios BNC)
  for (let i = 0; i < Math.min(rows.length, maxScan); i++) {
    const row = rows[i];
    if (!row) continue;
    const normalized = row.map((c) => normalizeHeader(c));
    if (normalized.some((h) => h === "haber" || h.includes("haber"))) {
      return i;
    }
  }

  for (let i = 0; i < Math.min(rows.length, maxScan); i++) {
    const row = rows[i];
    if (!row) continue;
    const text = row.map((c) => normalizeHeader(c)).join(" ");
    if (
      text.includes("factura") ||
      text.includes("referencia") ||
      (text.includes("fecha") && (text.includes("monto") || text.includes("debe")))
    ) {
      return i;
    }
  }

  return 0;
}

export function cellToString(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  return String(value).trim();
}

export function cellToDateString(value: unknown): string {
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  const str = cellToString(value);
  if (str.includes("T")) return str.slice(0, 10);
  const dmy = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (dmy) {
    const [, d, m, y] = dmy;
    return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  return str.slice(0, 10);
}

export function cellToFactura(value: unknown): number {
  const n = parseInt(String(value).replace(/\D/g, ""), 10);
  return Number.isNaN(n) ? NaN : n;
}
