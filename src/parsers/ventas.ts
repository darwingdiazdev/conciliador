import {
  cellToDateString,
  cellToFactura,
  detectHeaderRow,
  findColumnIndex,
  findColumnDollarSign,
  findUsdAmountColumn,
  bufferToRows,
} from "../utils/excel-helpers";
import { parseMonto } from "../utils/amounts";
import { VentaRow } from "../types";

export async function parseVentas(buffer: Buffer): Promise<VentaRow[]> {
  const allRows = bufferToRows(buffer);
  const headerIdx = detectHeaderRow(allRows);
  const headers = (allRows[headerIdx] ?? []).map(String);
  const dataRows = allRows.slice(headerIdx + 1);

  const colFecha = findColumnIndex(headers, ["fechaemision", "fechaemitida"]);
  const colFactura = findColumnIndex(headers, ["nfactura", "factura", "numerofactura"]);
  const colEmpresa = findColumnIndex(headers, ["empresa"]);
  let colCuotaIniUsd = findColumnDollarSign(headers, [["cuota", "inicial"]]);
  if (colCuotaIniUsd < 0) {
    colCuotaIniUsd = findUsdAmountColumn(headers, dataRows, {
      requiredTerms: [["cuota", "inicial"]],
      fallbackMatchers: ["cuotainicialusd", "cuotainicial"],
    });
  }

  let colCuotaPenUsd = findColumnDollarSign(headers, [["cuota", "pend"]]);
  if (colCuotaPenUsd < 0) {
    colCuotaPenUsd = findUsdAmountColumn(headers, dataRows, {
      requiredTerms: [
        ["cuota", "pend"],
        ["saldo", "pend"],
        ["montocuota"],
        ["valorcuota"],
        ["importecuota"],
      ],
      fallbackMatchers: ["cuotapendienteusd", "cuotapendiente", "cuotapendie"],
    });
  }
  const colTotalUsd = findUsdAmountColumn(headers, dataRows, {
    requiredTerms: [
      ["total", "usd"],
      ["totalusd"],
      ["montototal"],
      ["total", "venta"],
    ],
    fallbackMatchers: ["totalusd", "total"],
  });
  const colTasa = findColumnIndex(headers, [
    "tasadecambio",
    "tasadecamb",
    "tasadecam",
    "tasacamb",
    "tasabcv",
  ]);

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

    const cuotaInicialUsd =
      colCuotaIniUsd >= 0 ? parseMonto(row[colCuotaIniUsd]) : NaN;
    const cuotaPendienteUsd =
      colCuotaPenUsd >= 0 ? parseMonto(row[colCuotaPenUsd]) : NaN;

    let totalUsd = NaN;
    if (colTotalUsd >= 0) {
      const parsedTotal = parseMonto(row[colTotalUsd]);
      if (!Number.isNaN(parsedTotal)) totalUsd = parsedTotal;
    }
    if (Number.isNaN(totalUsd)) {
      totalUsd =
        (Number.isNaN(cuotaInicialUsd) ? 0 : cuotaInicialUsd) +
        (Number.isNaN(cuotaPendienteUsd) ? 0 : cuotaPendienteUsd);
    }

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
