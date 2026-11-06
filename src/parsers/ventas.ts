import {
  cellToDateString,
  cellToFactura,
  detectHeaderRow,
  findColumnIndex,
  findTasaColumn,
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
  const colSucursal = findColumnIndex(headers, ["sucursal", "tienda"]);
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
  const colTasa = findTasaColumn(headers);

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
      sucursal: colSucursal >= 0 ? String(row[colSucursal] ?? "").trim() : "",
      empresa: colEmpresa >= 0 ? String(row[colEmpresa] ?? "").trim() : "",
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

/** Título "tienda - empresa" y rango de fechas del archivo de ventas. */
export function encabezadoVentas(
  ventas: VentaRow[],
  hoy = new Date()
): { titulo: string; fechaDesde: string; fechaHasta: string } {
  const conteo = new Map<string, number>();
  for (const venta of ventas) {
    const tituloFila = [nombreTienda(venta.sucursal), venta.empresa.trim()]
      .filter(Boolean)
      .join(" - ");
    if (!tituloFila) continue;
    conteo.set(tituloFila, (conteo.get(tituloFila) ?? 0) + 1);
  }
  const titulo = [...conteo.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([nombre]) => nombre)
    .join(" · ");
  const { fechaDesde, fechaHasta } = rangoPorDistancia(ventas, hoy);
  return { titulo: titulo || "Conciliación de pagos", fechaDesde, fechaHasta };
}

function nombreTienda(sucursal: string): string {
  return sucursal.split("-")[0]?.trim() ?? "";
}

function rangoPorDistancia(
  ventas: VentaRow[],
  hoy: Date
): { fechaDesde: string; fechaHasta: string } {
  const hoyMs = Date.UTC(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
  let fechaDesde = "";
  let fechaHasta = "";
  let mayorDistancia = -1;
  let menorDistancia = Infinity;

  for (const venta of ventas) {
    const fecha = venta.fechaEmision;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) continue;
    const [anio, mes, dia] = fecha.split("-").map(Number);
    const distancia = Math.abs(Date.UTC(anio, mes - 1, dia) - hoyMs);
    if (distancia > mayorDistancia) {
      mayorDistancia = distancia;
      fechaDesde = fecha;
    }
    if (distancia < menorDistancia) {
      menorDistancia = distancia;
      fechaHasta = fecha;
    }
  }

  return { fechaDesde, fechaHasta };
}
