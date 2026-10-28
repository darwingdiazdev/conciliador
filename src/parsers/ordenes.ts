import { parseMonto } from "../utils/amounts";
import {
  bufferToRows,
  cellToDateString,
  detectHeaderRow,
  findColumnIndex,
  normalizeHeader,
} from "../utils/excel-helpers";
import {
  EstadoCuotaResumen,
  QuincenaRow,
  QuincenaTotales,
  ReporteOrdenes,
  ResumenQuincenas,
} from "../types";

const MESES = [
  "enero",
  "febrero",
  "marzo",
  "abril",
  "mayo",
  "junio",
  "julio",
  "agosto",
  "septiembre",
  "octubre",
  "noviembre",
  "diciembre",
];

interface CuotaCol {
  monto: number;
  estado: number;
}

function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

function montoOrZero(value: unknown): number {
  const n = parseMonto(value);
  return Number.isNaN(n) ? 0 : n;
}

function findCuotaColumns(headers: string[]): CuotaCol[] {
  const normalized = headers.map(normalizeHeader);
  const byN = new Map<number, CuotaCol>();

  normalized.forEach((header, idx) => {
    const monto = header.match(/^montocuota(\d+)$/);
    if (monto) {
      const n = Number(monto[1]);
      const current = byN.get(n) ?? { monto: -1, estado: -1 };
      current.monto = idx;
      byN.set(n, current);
    }
    const estado = header.match(/^estadocuota(\d+)$/);
    if (estado) {
      const n = Number(estado[1]);
      const current = byN.get(n) ?? { monto: -1, estado: -1 };
      current.estado = idx;
      byN.set(n, current);
    }
  });

  return [...byN.entries()]
    .filter(([, col]) => col.monto >= 0)
    .sort((a, b) => a[0] - b[0])
    .map(([, col]) => col);
}

function etiquetaQuincena(anio: number, mes: number, quincena: 1 | 2): string {
  const nombre = MESES[mes - 1] ?? String(mes);
  if (quincena === 1) return `1–15 ${nombre} ${anio}`;
  const ultimo = new Date(anio, mes, 0).getDate();
  return `16–${ultimo} ${nombre} ${anio}`;
}

function emptyTotales(): QuincenaTotales {
  return {
    ordenes: 0,
    totalVendido: 0,
    totalCuotaInicial: 0,
    totalCuotaPendiente: 0,
  };
}

interface GrupoAcc {
  empresa: string;
  quincenas: Map<string, QuincenaTotales & { anio: number; mes: number; q: 1 | 2 }>;
  estados: Map<
    string,
    {
      ordenes: number;
      totalVendido: number;
      totalCuotaInicial: number;
      montoCuota: number;
    }
  >;
  ordenesCanceladas: number;
  desde: string;
  hasta: string;
}

const ESTADO_RANK = ["DONE", "SCHEDULED", "GRACED", "PENDING", "DELAYED"];

function rankEstado(estado: string): number {
  if (estado === "CANCELLED") return 1000;
  const idx = ESTADO_RANK.indexOf(estado);
  return idx === -1 ? 500 : idx;
}

function estadoCuota(
  row: unknown[],
  cuotas: CuotaCol[],
  estadoOrden: string
): { estado: string; monto: number } {
  const encontrados: { estado: string; monto: number }[] = [];
  for (const cuota of cuotas) {
    const monto = montoOrZero(row[cuota.monto]);
    const raw = cuota.estado >= 0 ? String(row[cuota.estado] ?? "").trim() : "";
    if (!raw && !monto) continue;
    encontrados.push({
      estado: (raw || "SIN ESTADO").toUpperCase(),
      monto,
    });
  }

  if (encontrados.length === 0) {
    return { estado: estadoOrden || "SIN ESTADO", monto: 0 };
  }

  const preferido =
    encontrados.find((item) => item.estado !== "DONE" && item.estado !== "CANCELLED") ??
    encontrados.find((item) => item.estado === "DONE") ??
    encontrados[0];

  const monto = encontrados
    .filter((item) => item.estado === preferido.estado)
    .reduce((sum, item) => sum + item.monto, 0);

  return { estado: preferido.estado, monto };
}

export async function resumirOrdenesPorQuincena(
  buffer: Buffer
): Promise<ReporteOrdenes> {
  const allRows = bufferToRows(buffer);
  const headerIdx = detectHeaderRow(allRows);
  const headers = (allRows[headerIdx] ?? []).map((cell) => String(cell ?? ""));

  const colFechaCompra = findColumnIndex(headers, ["fechacompra"]);
  const colFechaFactura = findColumnIndex(headers, [
    "fechafacturacion",
    "fechafactura",
  ]);
  const colVenta = findColumnIndex(headers, ["ventatotal", "totalventa"]);
  const colInicial = findColumnIndex(headers, [
    "pagoencaja",
    "cuotainicial",
    "pagoinicial",
  ]);
  const colEstado = findColumnIndex(headers, ["estadoorden", "estadodelaorden"]);
  const colGrupo = findColumnIndex(headers, [
    "nombredelgrupo",
    "nombregrupo",
    "grupo",
  ]);
  const cuotas = findCuotaColumns(headers);

  if (colVenta < 0 && colInicial < 0 && cuotas.length === 0) {
    throw new Error(
      'No se reconocieron las columnas de órdenes (Venta total, Pago en caja o Monto cuota).'
    );
  }

  if (colFechaCompra < 0 && colFechaFactura < 0) {
    throw new Error(
      'No se encontró la columna "Fecha compra" en el archivo de órdenes.'
    );
  }

  const grupos = new Map<string, GrupoAcc>();

  for (let i = headerIdx + 1; i < allRows.length; i++) {
    const row = allRows[i];
    if (!row || row.every((cell) => cell === null || cell === "")) continue;

    const fechaRaw =
      colFechaCompra >= 0 && row[colFechaCompra]
        ? row[colFechaCompra]
        : colFechaFactura >= 0
          ? row[colFechaFactura]
          : "";
    const fecha = cellToDateString(fechaRaw);
    const match = fecha.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) continue;

    const anio = Number(match[1]);
    const mes = Number(match[2]);
    const dia = Number(match[3]);
    if (mes < 1 || mes > 12 || dia < 1 || dia > 31) continue;

    const estadoOrden =
      colEstado >= 0 ? normalizeHeader(row[colEstado]).toUpperCase() : "";
    const cuota = estadoCuota(row, cuotas, estadoOrden);
    if (estadoOrden.includes("CANCEL") || cuota.estado.includes("CANCEL")) {
      continue;
    }

    const empresaRaw =
      colGrupo >= 0 ? String(row[colGrupo] ?? "").trim() : "";
    const empresa = (empresaRaw || "Sin grupo").toUpperCase();
    const grupo = grupos.get(empresa) ?? {
      empresa,
      quincenas: new Map(),
      estados: new Map(),
      ordenesCanceladas: 0,
      desde: fecha,
      hasta: fecha,
    };
    if (fecha < grupo.desde) grupo.desde = fecha;
    if (fecha > grupo.hasta) grupo.hasta = fecha;

    const cuotaInicial = round2(colInicial >= 0 ? montoOrZero(row[colInicial]) : 0);
    const montoCuota = round2(cuota.monto);
    let venta = round2(colVenta >= 0 ? montoOrZero(row[colVenta]) : 0);
    if (!venta) venta = round2(cuotaInicial + montoCuota);
    if (!venta && !cuotaInicial && !montoCuota) continue;

    const bucket = grupo.estados.get(cuota.estado) ?? {
      ordenes: 0,
      totalVendido: 0,
      totalCuotaInicial: 0,
      montoCuota: 0,
    };
    bucket.ordenes += 1;
    bucket.totalVendido += venta;
    bucket.totalCuotaInicial += cuotaInicial;
    bucket.montoCuota += montoCuota;
    grupo.estados.set(cuota.estado, bucket);

    let cuotaPendiente = 0;
    for (const col of cuotas) {
      const monto = round2(montoOrZero(row[col.monto]));
      if (!monto) continue;
      const estado =
        col.estado >= 0 ? normalizeHeader(row[col.estado]) : "";
      if (estado.includes("cancel")) continue;
      cuotaPendiente += monto;
    }

    const quincena: 1 | 2 = dia <= 15 ? 1 : 2;
    const clave = `${anio}-${String(mes).padStart(2, "0")}-Q${quincena}`;
    const acc = grupo.quincenas.get(clave) ?? {
      ...emptyTotales(),
      anio,
      mes,
      q: quincena,
    };
    acc.ordenes += 1;
    acc.totalVendido += venta;
    acc.totalCuotaInicial += cuotaInicial;
    acc.totalCuotaPendiente += cuotaPendiente;
    grupo.quincenas.set(clave, acc);
    grupos.set(empresa, grupo);
  }

  if (grupos.size === 0) {
    throw new Error("No se encontraron órdenes válidas en el archivo.");
  }

  const reporte: ResumenQuincenas[] = [...grupos.values()]
    .sort((a, b) => a.empresa.localeCompare(b.empresa))
    .map((grupo) => {
      const quincenas: QuincenaRow[] = [...grupo.quincenas.entries()]
        .sort((a, b) => a[0].localeCompare(b[0]))
        .map(([clave, item]) => ({
          clave,
          etiqueta: etiquetaQuincena(item.anio, item.mes, item.q),
          ordenes: item.ordenes,
          totalVendido: round2(item.totalVendido),
          totalCuotaInicial: round2(item.totalCuotaInicial),
          totalCuotaPendiente: round2(item.totalCuotaPendiente),
        }));

      const totales = quincenas.reduce<QuincenaTotales>((acc, row) => {
        acc.ordenes += row.ordenes;
        acc.totalVendido += row.totalVendido;
        acc.totalCuotaInicial += row.totalCuotaInicial;
        acc.totalCuotaPendiente += row.totalCuotaPendiente;
        return acc;
      }, emptyTotales());
      totales.totalVendido = round2(totales.totalVendido);
      totales.totalCuotaInicial = round2(totales.totalCuotaInicial);
      totales.totalCuotaPendiente = round2(totales.totalCuotaPendiente);

      const estados: EstadoCuotaResumen[] = [...grupo.estados.entries()]
        .filter(([estado]) => !estado.includes("CANCEL"))
        .sort((a, b) => rankEstado(a[0]) - rankEstado(b[0]) || a[0].localeCompare(b[0]))
        .map(([estado, item]) => ({
          estado,
          ordenes: item.ordenes,
          totalVendido: round2(item.totalVendido),
          totalCuotaInicial: round2(item.totalCuotaInicial),
          montoCuota: round2(item.montoCuota),
        }));

      return {
        empresa: grupo.empresa,
        desde: grupo.desde,
        hasta: grupo.hasta,
        quincenas,
        totales,
        ordenesCanceladas: grupo.ordenesCanceladas,
        estados,
      };
    });

  return { grupos: reporte };
}
