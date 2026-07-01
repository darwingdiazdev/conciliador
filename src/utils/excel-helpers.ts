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

const USD_HEADER_HINTS = ["usd", "dolar", "dollar"];
const BS_HEADER_HINTS = ["bs", "bolivar", "ves"];

/** Busca columna en dólares por símbolo $ o etiqueta USD en el encabezado. */
export function findColumnDollarSign(
  headers: string[],
  keywordGroups: string[][]
): number {
  const normalized = headers.map(normalizeHeader);
  const rawHeaders = headers.map((h) => String(h ?? ""));

  for (const terms of keywordGroups) {
    for (let i = 0; i < headers.length; i++) {
      if (!terms.every((t) => normalized[i].includes(t))) continue;
      if (rawHeaders[i].includes("$")) return i;
      if (USD_HEADER_HINTS.some((hint) => normalized[i].includes(hint))) return i;
    }
  }

  return -1;
}

/** Elige la columna en USD entre varias candidatas (p. ej. cuota en $ vs cuota en Bs). */
export function findUsdAmountColumn(
  headers: string[],
  dataRows: unknown[][],
  options: {
    requiredTerms: string[][];
    fallbackMatchers?: string[];
  }
): number {
  const normalized = headers.map(normalizeHeader);
  const rawHeaders = headers.map((h) => String(h ?? "").toLowerCase());

  let candidates = normalized
    .map((h, idx) => ({ idx, h, raw: rawHeaders[idx] }))
    .filter(({ h }) =>
      options.requiredTerms.some((terms) => terms.every((t) => h.includes(t)))
    )
    .filter(({ h }) => !BS_HEADER_HINTS.some((bs) => h.includes(bs)));

  if (candidates.length === 0 && options.fallbackMatchers) {
    const fallbackIdx = findColumnIndex(headers, options.fallbackMatchers);
    if (fallbackIdx >= 0) {
      const h = normalized[fallbackIdx];
      if (!BS_HEADER_HINTS.some((bs) => h.includes(bs))) {
        return fallbackIdx;
      }
    }
    return -1;
  }

  if (candidates.length === 0) return -1;
  if (candidates.length === 1) return candidates[0].idx;

  const usdLabeled = candidates.filter(({ h, raw }) =>
    USD_HEADER_HINTS.some((hint) => h.includes(hint) || raw.includes("$"))
  );
  if (usdLabeled.length === 1) return usdLabeled[0].idx;

  const pool = usdLabeled.length > 0 ? usdLabeled : candidates;
  return pickColumnWithSmallestMedian(
    pool.map((c) => c.idx),
    dataRows
  );
}

function pickColumnWithSmallestMedian(
  indices: number[],
  dataRows: unknown[][]
): number {
  let bestIdx = indices[0];
  let bestMedian = Infinity;

  for (const idx of indices) {
    const vals = dataRows
      .map((row) => {
        const raw = row[idx];
        if (raw === null || raw === undefined || raw === "") return NaN;
        const n =
          typeof raw === "number"
            ? raw
            : parseFloat(
                String(raw)
                  .replace(/,/g, "")
                  .replace(/[^\d.-]/g, "")
              );
        return Number.isNaN(n) ? NaN : n;
      })
      .filter((v) => !Number.isNaN(v) && v > 0);

    if (vals.length === 0) continue;
    vals.sort((a, b) => a - b);
    const median = vals[Math.floor(vals.length / 2)];
    if (median < bestMedian) {
      bestMedian = median;
      bestIdx = idx;
    }
  }

  return bestIdx;
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
