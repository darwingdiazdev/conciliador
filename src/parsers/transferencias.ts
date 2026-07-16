import {
  cellToDateString,
  cellToString,
  detectHeaderRow,
  findColumnIndex,
  normalizeHeader,
  bufferToRows,
} from "../utils/excel-helpers";
import { unificarMonto } from "../utils/amounts";
import { TransferenciaRow } from "../types";

/**
 * Layout del archivo de transferencias del día:
 * A Fecha | B Sucursal | C Razón social | D Ticket | E Monto | F Banco | G Referencia
 */
const LAYOUT = {
  fecha: 0,
  sucursal: 1,
  razonSocial: 2,
  ticket: 3,
  monto: 4,
  banco: 5,
  referencia: 6,
} as const;

export async function parseTransferencias(
  buffer: Buffer
): Promise<TransferenciaRow[]> {
  const allRows = bufferToRows(buffer);
  const headerIdx = detectHeaderRowTransferencias(allRows);
  const headers = (allRows[headerIdx] ?? []).map(String);
  const cols = resolveColumns(headers);

  if (cols.monto < 0) {
    throw new Error(
      'No se encontró la columna "Monto" en el archivo de transferencias.'
    );
  }

  const transferencias: TransferenciaRow[] = [];

  for (let i = headerIdx + 1; i < allRows.length; i++) {
    const row = allRows[i];
    if (!row || row.every((c) => c === null || c === "")) continue;

    const monto = unificarMonto(row[cols.monto]);
    if (Number.isNaN(monto) || monto <= 0) continue;

    transferencias.push({
      fecha: cellToDateString(row[cols.fecha] ?? ""),
      sucursal: cellToString(row[cols.sucursal] ?? ""),
      razonSocial: cellToString(row[cols.razonSocial] ?? ""),
      ticket: formatTicket(row[cols.ticket]),
      monto,
      banco: cellToString(row[cols.banco] ?? ""),
      referencia: formatReferencia(row[cols.referencia]),
    });
  }

  if (transferencias.length === 0) {
    throw new Error(
      "No se encontraron transferencias válidas en el archivo."
    );
  }

  return transferencias;
}

function resolveColumns(headers: string[]) {
  const normalized = headers.map(normalizeHeader);
  const joined = normalized.join(" ");

  // Debe ser el archivo de transferencias (Sucursal, Ticket, Banco…)
  if (
    !joined.includes("sucursal") ||
    (!joined.includes("ticket") && !joined.includes("banco"))
  ) {
    throw new Error(
      'El archivo de transferencias no es válido. Debe tener columnas como Fecha, Sucursal, Razón social, Ticket, Monto, Banco y Referencia. ¿Subiste el estado de cuenta en ese campo por error?'
    );
  }

  // Archivo estándar → orden fijo
  if (
    joined.includes("sucursal") &&
    (joined.includes("ticket") || joined.includes("razonsocial") || joined.includes("razon"))
  ) {
    return { ...LAYOUT };
  }

  const byName = {
    fecha: findColumnIndex(headers, ["fecha"]),
    sucursal: findColumnIndex(headers, ["sucursal"]),
    razonSocial: findColumnIndex(headers, ["razonsocial", "razonso", "empresa"]),
    ticket: findColumnIndex(headers, ["ticket"]),
    monto: findColumnIndex(headers, ["monto", "importe"]),
    banco: findColumnIndex(headers, ["banco"]),
    referencia: findExactReferencia(normalized),
  };

  if (byName.sucursal < 0 || byName.monto < 0) {
    throw new Error(
      'No se encontraron las columnas Sucursal/Monto en el archivo de transferencias.'
    );
  }

  return {
    fecha: byName.fecha >= 0 ? byName.fecha : LAYOUT.fecha,
    sucursal: byName.sucursal,
    razonSocial:
      byName.razonSocial >= 0 ? byName.razonSocial : LAYOUT.razonSocial,
    ticket: byName.ticket >= 0 ? byName.ticket : LAYOUT.ticket,
    monto: byName.monto,
    banco: byName.banco >= 0 ? byName.banco : LAYOUT.banco,
    referencia:
      byName.referencia >= 0 ? byName.referencia : LAYOUT.referencia,
  };
}

function findExactReferencia(normalized: string[]): number {
  const exact = normalized.findIndex((h) => h === "referencia");
  if (exact >= 0) return exact;
  return normalized.findIndex((h) => h.includes("referencia"));
}

function formatTicket(value: unknown): string {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return "";
    return String(Math.round(value));
  }
  const str = String(value).trim();
  if (/e\+/i.test(str)) {
    const n = Number(str);
    if (Number.isFinite(n)) return String(Math.round(n));
  }
  return str;
}

function formatReferencia(value: unknown): string {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return "";
    // Conservar decimales si los tiene (ej. 3911.72)
    return Number.isInteger(value) ? String(value) : String(value);
  }
  return String(value).trim();
}

function detectHeaderRowTransferencias(
  rows: unknown[][],
  maxScan = 30
): number {
  for (let i = 0; i < Math.min(rows.length, maxScan); i++) {
    const row = rows[i];
    if (!row) continue;
    const text = row.map((c) => normalizeHeader(c)).join(" ");
    if (
      (text.includes("ticket") && text.includes("monto")) ||
      (text.includes("sucursal") && text.includes("monto")) ||
      (text.includes("referencia") && text.includes("banco"))
    ) {
      return i;
    }
  }
  return detectHeaderRow(rows, maxScan);
}
