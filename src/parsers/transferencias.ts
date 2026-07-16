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
 * Layout típico del archivo de transferencias del día:
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
      referencia: cellToString(row[cols.referencia] ?? ""),
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
  const byName = {
    fecha: findColumnIndex(headers, ["fecha"]),
    sucursal: findColumnIndex(headers, ["sucursal"]),
    razonSocial: findColumnIndex(headers, [
      "razonsocial",
      "razonso",
      "empresa",
    ]),
    ticket: findColumnIndex(headers, ["ticket"]),
    monto: findColumnIndex(headers, ["monto", "importe"]),
    banco: findColumnIndex(headers, ["banco"]),
    referencia: findExactOrIncludes(headers, ["referencia"]),
  };

  // Si faltan columnas clave, usar el orden fijo del archivo de transferencias
  const missingCore =
    byName.sucursal < 0 ||
    byName.razonSocial < 0 ||
    byName.ticket < 0 ||
    byName.banco < 0;

  if (missingCore && looksLikeTransferenciasLayout(headers)) {
    return { ...LAYOUT };
  }

  return {
    fecha: byName.fecha >= 0 ? byName.fecha : LAYOUT.fecha,
    sucursal: byName.sucursal >= 0 ? byName.sucursal : LAYOUT.sucursal,
    razonSocial:
      byName.razonSocial >= 0 ? byName.razonSocial : LAYOUT.razonSocial,
    ticket: byName.ticket >= 0 ? byName.ticket : LAYOUT.ticket,
    monto: byName.monto >= 0 ? byName.monto : LAYOUT.monto,
    banco: byName.banco >= 0 ? byName.banco : LAYOUT.banco,
    referencia:
      byName.referencia >= 0 ? byName.referencia : LAYOUT.referencia,
  };
}

function looksLikeTransferenciasLayout(headers: string[]): boolean {
  const joined = headers.map(normalizeHeader).join(" ");
  return (
    joined.includes("sucursal") ||
    joined.includes("ticket") ||
    joined.includes("banco")
  );
}

function findExactOrIncludes(headers: string[], matchers: string[]): number {
  const normalized = headers.map(normalizeHeader);
  for (const matcher of matchers) {
    const exact = normalized.findIndex((h) => h === matcher);
    if (exact >= 0) return exact;
  }
  for (const matcher of matchers) {
    const idx = normalized.findIndex((h) => h.includes(matcher));
    if (idx >= 0) return idx;
  }
  return -1;
}

function formatTicket(value: unknown): string {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return "";
    return Number.isInteger(value)
      ? String(value)
      : String(Math.round(value));
  }
  const str = String(value).trim();
  if (/e\+/i.test(str)) {
    const n = Number(str);
    if (Number.isFinite(n)) return String(Math.round(n));
  }
  return str;
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
