import {
  cellToDateString,
  cellToFactura,
  detectHeaderRow,
  findColumnIndex,
  bufferToRows,
} from "../utils/excel-helpers";
import { parseMonto } from "../utils/amounts";
import { VentaRow } from "../types";

export async function parseVentas(buffer: Buffer): Promise<VentaRow[]> {
  const allRows = bufferToRows(buffer);
  const headerIdx = detectHeaderRow(allRows);
  const headers = (allRows[headerIdx] ?? []).map(String);

  const colFecha = findColumnIndex(headers, ["fechaemision", "fechaemitida"]);
  const colFactura = findColumnIndex(headers, ["nfactura", "factura", "numerofactura"]);
  const colEmpresa = findColumnIndex(headers, ["empresa"]);
  const colCuotaIniUsd = findColumnIndex(headers, [
    "cuotainicialusd",
    "cuotainicial",
  ]);
  const colCuotaPenUsd = findColumnIndex(headers, [
    "cuotapendienteusd",
    "cuotapendiente",
    "cuotapendie",
  ]);
  const colTotalUsd = findColumnIndex(headers, ["totalusd", "total"]);
  const colTasa = findColumnIndex(headers, ["tasadecambio", "tasacamb", "tasabcv"]);

  if (colFactura < 0) {
    throw new Error(
      'No se encontró la columna "N° Factura" en el archivo de ventas.'
    );
  }

  const ventas: VentaRow[] = [];
  const seen = new Set<number>();

  for (let i = headerIdx + 1; i < allRows.length; i++) {
    const row = allRows[i];
    if (!row || row.every((c) => c === null || c === "")) continue;

    const numeroFactura = cellToFactura(row[colFactura]);
    if (Number.isNaN(numeroFactura) || seen.has(numeroFactura)) continue;
    seen.add(numeroFactura);

    let cuotaInicialUsd = colCuotaIniUsd >= 0 ? parseMonto(row[colCuotaIniUsd]) : NaN;
    let cuotaPendienteUsd = colCuotaPenUsd >= 0 ? parseMonto(row[colCuotaPenUsd]) : NaN;

    // Si hay varias columnas "cuota inicial", tomar la que parece USD (valor menor)
    if (colCuotaIniUsd >= 0) {
      const candidates = headers
        .map((h, idx) => ({ h: h.toLowerCase(), idx }))
        .filter(({ h }) => h.includes("cuota") && h.includes("inicial"));
      if (candidates.length > 1) {
        const usdCol = candidates
          .map((c) => ({ idx: c.idx, val: parseMonto(row[c.idx]) }))
          .filter((c) => !Number.isNaN(c.val))
          .sort((a, b) => a.val - b.val)[0];
        if (usdCol) cuotaInicialUsd = usdCol.val;
      }
    }

    if (colCuotaPenUsd >= 0) {
      const candidates = headers
        .map((h, idx) => ({ h: h.toLowerCase(), idx }))
        .filter(({ h }) => h.includes("cuota") && h.includes("pend"));
      if (candidates.length > 1) {
        const usdCol = candidates
          .map((c) => ({ idx: c.idx, val: parseMonto(row[c.idx]) }))
          .filter((c) => !Number.isNaN(c.val))
          .sort((a, b) => a.val - b.val)[0];
        if (usdCol) cuotaPendienteUsd = usdCol.val;
      }
    }

    const totalUsd =
      colTotalUsd >= 0 && parseMonto(row[colTotalUsd]) < 1000
        ? parseMonto(row[colTotalUsd])
        : (Number.isNaN(cuotaInicialUsd) ? 0 : cuotaInicialUsd) +
          (Number.isNaN(cuotaPendienteUsd) ? 0 : cuotaPendienteUsd);

    ventas.push({
      fechaEmision: colFecha >= 0 ? cellToDateString(row[colFecha]) : "",
      numeroFactura,
      empresa: colEmpresa >= 0 ? String(row[colEmpresa] ?? "") : "",
      cuotaInicialUsd: Number.isNaN(cuotaInicialUsd) ? 0 : cuotaInicialUsd,
      cuotaPendienteUsd: Number.isNaN(cuotaPendienteUsd) ? 0 : cuotaPendienteUsd,
      totalUsd,
      tasaCambio: colTasa >= 0 ? parseMonto(row[colTasa]) : undefined,
    });
  }

  if (ventas.length === 0) {
    throw new Error("No se encontraron ventas válidas en el archivo.");
  }

  return ventas;
}
