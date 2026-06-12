import {
  cellToDateString,
  cellToFactura,
  cellToString,
  detectHeaderRow,
  findColumnIndex,
  bufferToRows,
} from "../utils/excel-helpers";
import { parseMonto } from "../utils/amounts";
import { MerchantRow } from "../types";

export async function parseMerchant(buffer: Buffer): Promise<MerchantRow[]> {
  const allRows = bufferToRows(buffer);
  const headerIdx = detectHeaderRow(allRows);
  const headers = (allRows[headerIdx] ?? []).map(String);

  const colFecha = findColumnIndex(headers, ["fechadetransaccion", "fechatransaccion", "fecha"]);
  const colFactura = findColumnIndex(headers, ["nfactura", "factura", "numerofactura"]);
  const colMontoBs = findColumnIndex(headers, ["montopagado", "montopagad"]);
  const colMontoUsd = findColumnIndex(headers, ["montopagadousd", "montoasignado", "montousd"]);
  const colTasa = findColumnIndex(headers, ["tasadecambio", "tasacamb"]);
  const colRef = findColumnIndex(headers, ["referencia", "nreferencia"]);

  if (colFactura < 0) {
    throw new Error(
      'No se encontró la columna "# Factura" en el archivo de merchant (Cashea).'
    );
  }

  const merchants: MerchantRow[] = [];
  const byFactura = new Map<number, MerchantRow>();

  for (let i = headerIdx + 1; i < allRows.length; i++) {
    const row = allRows[i];
    if (!row || row.every((c) => c === null || c === "")) continue;

    const numeroFactura = cellToFactura(row[colFactura]);
    if (Number.isNaN(numeroFactura)) continue;

    let montoPagadoBs = colMontoBs >= 0 ? parseMonto(row[colMontoBs]) : NaN;
    let montoPagadoUsd = colMontoUsd >= 0 ? parseMonto(row[colMontoUsd]) : NaN;

    // Monto Bs suele ser el valor mayor; USD el menor
    const montoCols = headers
      .map((h, idx) => ({ h: normalizeMontoHeader(h), idx }))
      .filter(({ h }) => h.includes("monto"));

    if (montoCols.length >= 2 && (Number.isNaN(montoPagadoBs) || Number.isNaN(montoPagadoUsd))) {
      const vals = montoCols
        .map((c) => ({ idx: c.idx, val: parseMonto(row[c.idx]) }))
        .filter((c) => !Number.isNaN(c.val))
        .sort((a, b) => a.val - b.val);
      if (vals.length >= 2) {
        montoPagadoUsd = vals[0].val;
        montoPagadoBs = vals[vals.length - 1].val;
      } else if (vals.length === 1) {
        montoPagadoBs = vals[0].val;
      }
    }

    const entry: MerchantRow = {
      fechaTransaccion: colFecha >= 0 ? cellToDateString(row[colFecha]) : "",
      numeroFactura,
      montoPagadoBs: Number.isNaN(montoPagadoBs) ? 0 : montoPagadoBs,
      montoPagadoUsd: Number.isNaN(montoPagadoUsd) ? 0 : montoPagadoUsd,
      tasaCambio: colTasa >= 0 ? parseMonto(row[colTasa]) : 0,
      referencia: colRef >= 0 ? cellToString(row[colRef]) : "",
    };

    // Una fila por factura: conservar el registro más reciente o con mayor monto
    const existing = byFactura.get(numeroFactura);
    if (!existing || entry.montoPagadoBs >= existing.montoPagadoBs) {
      byFactura.set(numeroFactura, entry);
    }
  }

  byFactura.forEach((m) => merchants.push(m));

  if (merchants.length === 0) {
    throw new Error("No se encontraron registros válidos en el archivo de Cashea.");
  }

  return merchants;
}

function normalizeMontoHeader(h: string): string {
  return h.toLowerCase().replace(/[^a-z0-9]/g, "");
}
