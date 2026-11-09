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

  const periodo = periodoDelTitulo(allRows.slice(0, headerIdx));
  const movimientos: BancoRow[] = [];
  let saldoFinal = NaN;

  for (let i = headerIdx + 1; i < allRows.length; i++) {
    const row = allRows[i];
    if (!row || row.every((c) => c === null || c === "")) continue;

    const saldo = parseMonto(row[colSaldo]);
    if (!Number.isNaN(saldo)) {
      saldoFinal = saldo;
    }

    const filaTexto = row.map((c) => cellToString(c).toLowerCase()).join(" ");
    if (filaTexto.includes("total")) continue;

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

  return { movimientos, saldoFinal, ...periodo };
}

/** "MOVIMIENTOS ENTRE FECHAS DEL 19/10/2026 AL 1/11/2026" */
function periodoDelTitulo(rows: unknown[][]): {
  fechaDesde?: string;
  fechaHasta?: string;
} {
  const re =
    /del\s+(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})\s+al\s+(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})/i;

  for (const row of rows) {
    if (!row) continue;
    for (const cell of row) {
      const match = cellToString(cell).match(re);
      if (!match) continue;
      return {
        fechaDesde: iso(match[3], match[2], match[1]),
        fechaHasta: iso(match[6], match[5], match[4]),
      };
    }
  }

  return {};
}

function iso(anio: string, mes: string, dia: string): string {
  return `${anio}-${mes.padStart(2, "0")}-${dia.padStart(2, "0")}`;
}
