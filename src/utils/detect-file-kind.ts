import * as XLSX from "xlsx";
import { normalizeHeader } from "../utils/excel-helpers";

export type ArchivoKind = "transferencias" | "estadoCuenta" | "desconocido";

/** Detecta el tipo de Excel por sus encabezados. */
export function detectArchivoKind(buffer: Buffer): ArchivoKind {
  const headers = readHeaders(buffer);
  const joined = headers.map(normalizeHeader).join(" ");

  const isTransferencias =
    joined.includes("sucursal") &&
    (joined.includes("ticket") || joined.includes("banco")) &&
    joined.includes("monto");

  const isEstadoCuenta =
    joined.includes("monto") &&
    (joined.includes("descrip") || joined.includes("saldo")) &&
    !joined.includes("sucursal") &&
    !joined.includes("ticket");

  if (isTransferencias) return "transferencias";
  if (isEstadoCuenta) return "estadoCuenta";
  return "desconocido";
}

function readHeaders(buffer: Buffer): string[] {
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) return [];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: null,
    raw: false,
  }) as unknown[][];

  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const row = rows[i];
    if (!row) continue;
    const text = row.map((c) => normalizeHeader(c)).join(" ");
    if (
      text.includes("monto") ||
      text.includes("ticket") ||
      text.includes("descrip") ||
      text.includes("sucursal")
    ) {
      return row.map((c) => (c == null ? "" : String(c)));
    }
  }
  return (rows[0] ?? []).map((c) => (c == null ? "" : String(c)));
}
