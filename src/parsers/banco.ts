import {
  cellToDateString,
  cellToString,
  detectHeaderRow,
  findColumnIndex,
  bufferToRows,
} from "../utils/excel-helpers";
import { parseMonto } from "../utils/amounts";
import { BancoParseResult, BancoRow } from "../types";

export async function parseBanco(buffer: Buffer): Promise<BancoParseResult> {
  const allRows = bufferToRows(buffer);
  const headerIdx = detectHeaderRow(allRows);
  const headers = (allRows[headerIdx] ?? []).map(String);

  const colFecha = findColumnIndex(headers, ["fecha"]);
  const colHaber = findColumnIndex(headers, ["haber", "credito", "abono"]);
  const colSaldo = findColumnIndex(headers, ["saldo"]);
  const colRef = findColumnIndex(headers, ["referencia"]);
  const colDesc = findColumnIndex(headers, ["descripcion", "descrip"]);
  const colTipo = findColumnIndex(headers, ["tipooperacion", "tipotransaccion", "tipo"]);

  if (colHaber < 0) {
    throw new Error('No se encontró la columna "Haber" en el extracto del banco.');
  }
  if (colSaldo < 0) {
    throw new Error('No se encontró la columna "Saldo" en el extracto del banco.');
  }

  const movimientos: BancoRow[] = [];
  let saldoFinal = NaN;

  for (let i = headerIdx + 1; i < allRows.length; i++) {
    const row = allRows[i];
    if (!row || row.every((c) => c === null || c === "")) continue;

    const saldo = parseMonto(row[colSaldo]);
    if (!Number.isNaN(saldo)) {
      saldoFinal = saldo;
    }

    const firstCell = cellToString(row[0]).toLowerCase();
    if (firstCell.includes("total")) continue;

    const tipoOperacion =
      colTipo >= 0 ? cellToString(row[colTipo]) : colDesc >= 0 ? cellToString(row[colDesc]) : "";
    const descripcion = colDesc >= 0 ? cellToString(row[colDesc]) : tipoOperacion;

    const tipoLower = tipoOperacion.toLowerCase();
    const descLower = descripcion.toLowerCase();
    if (
      tipoLower.includes("comision") ||
      tipoLower.includes("comisión") ||
      descLower.includes("comision") ||
      descLower.includes("comisión")
    ) {
      continue;
    }

    const haber = parseMonto(row[colHaber]);
    if (Number.isNaN(haber) || haber <= 0) continue;

    let referencia = colRef >= 0 ? cellToString(row[colRef]) : `${i}`;
    if (/e\+/i.test(referencia)) {
      referencia = Number(referencia).toFixed(0);
    }

    movimientos.push({
      fecha: colFecha >= 0 ? cellToDateString(row[colFecha]) : "",
      referencia,
      haber,
      descripcion,
      tipoOperacion,
    });
  }

  if (movimientos.length === 0) {
    throw new Error("No se encontraron abonos válidos en el extracto del banco.");
  }
  if (Number.isNaN(saldoFinal)) {
    throw new Error("No se pudo determinar el saldo final del extracto bancario.");
  }

  return { movimientos, saldoFinal };
}
