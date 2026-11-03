import {
  cellToDateString,
  cellToFactura,
  cellToString,
  detectHeaderRow,
  findColumnIndex,
  findTasaColumn,
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
  const colMontoBs = findColumnIndex(headers, [
    "montopagadoenves",
    "montopagadoenbs",
    "montopagado",
    "montopagad",
  ]);
  const colMontoUsd = findColumnIndex(headers, [
    "montopagadoenusd",
    "montopagadousd",
    "montoasignado",
    "montousd",
  ]);
  const colTasa = findTasaColumn(headers);
  const colRef = findColumnIndex(headers, ["referencia", "nreferencia"]);

  if (colFactura < 0) {
    throw new Error(
      'No se encontró la columna "# Factura" en el archivo de merchant (Cashea).'
    );
  }

  const merchants: MerchantRow[] = [];

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

    merchants.push(entry);
  }

  merchants.sort(
    (a, b) =>
      a.numeroFactura - b.numeroFactura ||
      a.fechaTransaccion.localeCompare(b.fechaTransaccion) ||
      a.referencia.localeCompare(b.referencia)
  );

  if (merchants.length === 0) {
    throw new Error("No se encontraron registros válidos en el archivo de Cashea.");
  }

  return merchants;
}

function normalizeMontoHeader(h: string): string {
  return h.toLowerCase().replace(/[^a-z0-9]/g, "");
}
