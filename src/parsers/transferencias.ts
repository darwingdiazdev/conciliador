import {
  cellToDateString,
  cellToString,
  detectHeaderRow,
  findColumnIndex,
  bufferToRows,
} from "../utils/excel-helpers";
import { unificarMonto } from "../utils/amounts";
import { TransferenciaRow } from "../types";

export async function parseTransferencias(
  buffer: Buffer
): Promise<TransferenciaRow[]> {
  const allRows = bufferToRows(buffer);
  const headerIdx = detectHeaderRowTransferencias(allRows);
  const headers = (allRows[headerIdx] ?? []).map(String);

  const colFecha = findColumnIndex(headers, ["fecha"]);
  const colSucursal = findColumnIndex(headers, ["sucursal"]);
  const colRazon = findColumnIndex(headers, ["razonsocial", "empresa", "razon"]);
  const colTicket = findColumnIndex(headers, ["ticket"]);
  const colMonto = findColumnIndex(headers, ["monto", "importe"]);
  const colBanco = findColumnIndex(headers, ["banco"]);
  const colRef = findColumnIndex(headers, ["referencia", "ref"]);
  const colPrevia = findColumnIndex(headers, [
    "transferenciaprevia",
    "transfprevia",
    "previa",
  ]);

  if (colMonto < 0) {
    throw new Error(
      'No se encontró la columna "Monto" en el archivo de transferencias.'
    );
  }

  const transferencias: TransferenciaRow[] = [];

  for (let i = headerIdx + 1; i < allRows.length; i++) {
    const row = allRows[i];
    if (!row || row.every((c) => c === null || c === "")) continue;

    // Unifica 3,250.00 → 3250.00 antes de conciliar
    const monto = unificarMonto(row[colMonto]);
    if (Number.isNaN(monto) || monto <= 0) continue;

    transferencias.push({
      fecha: colFecha >= 0 ? cellToDateString(row[colFecha]) : "",
      sucursal: colSucursal >= 0 ? cellToString(row[colSucursal]) : "",
      razonSocial: colRazon >= 0 ? cellToString(row[colRazon]) : "",
      ticket: colTicket >= 0 ? cellToString(row[colTicket]) : "",
      monto,
      banco: colBanco >= 0 ? cellToString(row[colBanco]) : "",
      referencia: colRef >= 0 ? cellToString(row[colRef]) : "",
      transferenciaPrevia:
        colPrevia >= 0 ? cellToString(row[colPrevia]) : "",
    });
  }

  if (transferencias.length === 0) {
    throw new Error(
      "No se encontraron transferencias válidas en el archivo."
    );
  }

  return transferencias;
}

function detectHeaderRowTransferencias(
  rows: unknown[][],
  maxScan = 30
): number {
  for (let i = 0; i < Math.min(rows.length, maxScan); i++) {
    const row = rows[i];
    if (!row) continue;
    const text = row
      .map((c) =>
        String(c ?? "")
          .normalize("NFD")
          .replace(/[\u0300-\u036f]/g, "")
          .toLowerCase()
      )
      .join(" ");
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
