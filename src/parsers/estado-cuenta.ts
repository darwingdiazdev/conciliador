import {
  cellToDateString,
  cellToString,
  detectHeaderRow,
  findColumnIndex,
  bufferToRows,
} from "../utils/excel-helpers";
import { unificarMonto } from "../utils/amounts";
import { EstadoCuentaRow } from "../types";

export async function parseEstadoCuenta(buffer: Buffer): Promise<EstadoCuentaRow[]> {
  const allRows = bufferToRows(buffer);
  const headerIdx = detectHeaderRowEstado(allRows);
  const headers = (allRows[headerIdx] ?? []).map(String);

  const colFecha = findColumnIndex(headers, ["fecha"]);
  const colDesc = findColumnIndex(headers, ["descripcion", "descrip", "concepto"]);
  const colMonto = findColumnIndex(headers, ["monto", "importe", "haber", "credito"]);
  const colSaldo = findColumnIndex(headers, ["saldo"]);

  if (colMonto < 0) {
    throw new Error(
      'No se encontró la columna "Monto" en el estado de cuenta.'
    );
  }

  const movimientos: EstadoCuentaRow[] = [];

  for (let i = headerIdx + 1; i < allRows.length; i++) {
    const row = allRows[i];
    if (!row || row.every((c) => c === null || c === "")) continue;

    const descripcion = colDesc >= 0 ? cellToString(row[colDesc]) : "";
    const descLower = descripcion.toLowerCase();

    if (
      descLower.includes("comision") ||
      descLower.includes("comisión") ||
      descLower.includes("comis.")
    ) {
      continue;
    }

    // Unifica 3.250,00 → 3250.00 antes de conciliar
    const monto = unificarMonto(row[colMonto]);
    if (Number.isNaN(monto) || monto <= 0) continue;

    movimientos.push({
      fecha: colFecha >= 0 ? cellToDateString(row[colFecha]) : "",
      descripcion,
      monto,
      saldo: colSaldo >= 0 ? unificarMonto(row[colSaldo]) : 0,
    });
  }

  if (movimientos.length === 0) {
    throw new Error(
      "No se encontraron abonos válidos en el estado de cuenta."
    );
  }

  return movimientos;
}

function detectHeaderRowEstado(rows: unknown[][], maxScan = 30): number {
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
      (text.includes("fecha") && text.includes("monto")) ||
      (text.includes("descripcion") && text.includes("monto")) ||
      (text.includes("fecha") && text.includes("saldo"))
    ) {
      return i;
    }
  }
  return detectHeaderRow(rows, maxScan);
}
