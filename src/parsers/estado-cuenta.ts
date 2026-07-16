import * as XLSX from "xlsx";
import {
  cellToDateString,
  cellToString,
  detectHeaderRow,
  findColumnIndex,
  normalizeHeader,
} from "../utils/excel-helpers";
import { parseMontoFlexible } from "../utils/amounts";
import { EstadoCuentaRow } from "../types";

/**
 * Layout típico del export bancario:
 * A Fecha | B Descripción | C Monto | D Saldo
 *
 * Los montos vienen en miles (1.25 = 1.250,00 Bs); la escala se corrige en reconcile.
 */
const LAYOUT = {
  fecha: 0,
  descripcion: 1,
  monto: 2,
  saldo: 3,
} as const;

export async function parseEstadoCuenta(buffer: Buffer): Promise<EstadoCuentaRow[]> {
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("El estado de cuenta no contiene hojas de cálculo.");
  }

  const sheet = workbook.Sheets[sheetName];
  const allRows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: null,
    raw: true,
  }) as unknown[][];

  const headerIdx = detectHeaderRowEstado(allRows);
  const headers = (allRows[headerIdx] ?? []).map((h) =>
    h === null || h === undefined ? "" : String(h)
  );
  const cols = resolveColumns(headers);

  if (cols.monto < 0) {
    throw new Error(
      'No se encontró la columna "Monto" en el estado de cuenta.'
    );
  }

  const movimientos: EstadoCuentaRow[] = [];

  for (let i = headerIdx + 1; i < allRows.length; i++) {
    const row = allRows[i];
    if (!row || row.every((c) => c === null || c === "")) continue;

    const descripcion =
      cols.descripcion >= 0 ? cellToString(row[cols.descripcion]) : "";
    const descLower = descripcion.toLowerCase();

    if (
      descLower.includes("comision") ||
      descLower.includes("comisión") ||
      descLower.includes("comis.")
    ) {
      continue;
    }

    const montoRaw = parseMontoFlexible(row[cols.monto]);
    if (Number.isNaN(montoRaw) || montoRaw <= 0) continue;

    const saldoRaw =
      cols.saldo >= 0 ? parseMontoFlexible(row[cols.saldo]) : 0;

    movimientos.push({
      fecha: cols.fecha >= 0 ? cellToDateString(row[cols.fecha]) : "",
      descripcion,
      monto: montoRaw,
      saldo: Number.isNaN(saldoRaw) ? 0 : saldoRaw,
    });
  }

  if (movimientos.length === 0) {
    throw new Error(
      "No se encontraron abonos válidos en el estado de cuenta."
    );
  }

  return movimientos;
}

function resolveColumns(headers: string[]) {
  const normalized = headers.map(normalizeHeader);
  const joined = normalized.join(" ");

  // Archivo estándar: Fecha | Descripción | Monto | Saldo
  if (
    joined.includes("fecha") &&
    joined.includes("monto") &&
    (joined.includes("descrip") || joined.includes("saldo"))
  ) {
    return {
      fecha: indexOr(headers, ["fecha"], LAYOUT.fecha),
      descripcion: indexOr(
        headers,
        ["descripcion", "descripci", "descrip", "concepto", "detalle"],
        LAYOUT.descripcion
      ),
      monto: indexOr(headers, ["monto", "importe", "haber", "credito"], LAYOUT.monto),
      saldo: indexOr(headers, ["saldo"], LAYOUT.saldo),
    };
  }

  return {
    fecha: findColumnIndex(headers, ["fecha"]),
    descripcion: findColumnIndex(headers, [
      "descripcion",
      "descripci",
      "descrip",
      "concepto",
      "detalle",
    ]),
    monto: findColumnIndex(headers, ["monto", "importe", "haber", "credito"]),
    saldo: findColumnIndex(headers, ["saldo"]),
  };
}

function indexOr(
  headers: string[],
  matchers: string[],
  fallback: number
): number {
  const idx = findColumnIndex(headers, matchers);
  return idx >= 0 ? idx : fallback;
}

function detectHeaderRowEstado(rows: unknown[][], maxScan = 30): number {
  for (let i = 0; i < Math.min(rows.length, maxScan); i++) {
    const row = rows[i];
    if (!row) continue;
    const text = row.map((c) => normalizeHeader(c)).join(" ");
    if (
      (text.includes("fecha") && text.includes("monto")) ||
      (text.includes("descrip") && text.includes("monto")) ||
      (text.includes("fecha") && text.includes("saldo"))
    ) {
      return i;
    }
  }
  return detectHeaderRow(rows, maxScan);
}
