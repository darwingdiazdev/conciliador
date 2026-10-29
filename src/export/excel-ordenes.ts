import ExcelJS from "exceljs";
import { ReporteOrdenes, ResumenQuincenas } from "../types";

const ORANGE = "FFE8590C";
const ORANGE_SOFT = "FFFFF1E6";
const GREY = "FFF2F2F2";
const WHITE = "FFFFFFFF";
const BLACK = "FF000000";
const USD = '"$"#,##0.00';

export async function exportOrdenesQuincenas(
  reporte: ReporteOrdenes
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const usedNames = new Set<string>();

  for (const grupo of reporte.grupos) {
    const sheet = workbook.addWorksheet(uniqueSheetName(grupo.empresa, usedNames));
    writeGrupo(sheet, grupo);
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

function writeGrupo(sheet: ExcelJS.Worksheet, grupo: ResumenQuincenas): void {
  sheet.columns = [
    { width: 28 },
    { width: 14 },
    { width: 28 },
    { width: 24 },
    { width: 26 },
  ];

  sheet.getRow(1).height = 28;
  merge(sheet, "A1:E1", grupo.empresa, {
    bold: true,
    size: 16,
    color: { argb: WHITE },
  }, ORANGE);

  sheet.getRow(2).height = 18;
  merge(
    sheet,
    "A2:E2",
    formatRango(grupo.desde, grupo.hasta),
    { bold: true, size: 11, color: { argb: BLACK } },
    GREY
  );

  let row = 4;
  row = writeSection(sheet, row, "Ventas por quincena");
  row = writeTable(
    sheet,
    row,
    ["Quincena", "Órdenes", "Total vendido por Cashea", "Total cuota inicial", "Total cuota pendiente"],
    grupo.quincenas.map((item) => [
      item.etiqueta,
      item.ordenes,
      item.totalVendido,
      item.totalCuotaInicial,
      item.totalCuotaPendiente,
    ]),
    [
      "Total",
      grupo.totales.ordenes,
      grupo.totales.totalVendido,
      grupo.totales.totalCuotaInicial,
      grupo.totales.totalCuotaPendiente,
    ]
  );

  row += 1;
  row = writeSection(sheet, row, "Estados de la cuota");
  const totalEstados = grupo.estados.reduce(
    (acc, item) => {
      acc.ordenes += item.ordenes;
      acc.totalVendido += item.totalVendido;
      acc.totalCuotaInicial += item.totalCuotaInicial;
      acc.montoCuota += item.montoCuota;
      return acc;
    },
    { ordenes: 0, totalVendido: 0, totalCuotaInicial: 0, montoCuota: 0 }
  );
  totalEstados.totalVendido = round2(totalEstados.totalVendido);
  totalEstados.totalCuotaInicial = round2(totalEstados.totalCuotaInicial);
  totalEstados.montoCuota = round2(totalEstados.montoCuota);

  row = writeTable(
    sheet,
    row,
    ["Estado", "Órdenes", "Total vendido por Cashea", "Total cuota inicial", "Monto cuota"],
    grupo.estados.map((item) => [
      traducirEstado(item.estado),
      item.ordenes,
      item.totalVendido,
      item.totalCuotaInicial,
      item.montoCuota,
    ]),
    [
      "Total",
      totalEstados.ordenes,
      totalEstados.totalVendido,
      totalEstados.totalCuotaInicial,
      totalEstados.montoCuota,
    ]
  );

  sheet.views = [{ state: "frozen", ySplit: 2 }];
}

function traducirEstado(estado: string): string {
  const etiquetas: Record<string, string> = {
    DONE: "Pagada",
    SCHEDULED: "Programada",
    SHEDULED: "Programada",
    SHEDULE: "Programada",
    GRACED: "En gracia",
    CANCELLED: "Cancelada",
    CANCELED: "Cancelada",
    "SIN ESTADO": "Sin estado",
  };
  return etiquetas[estado] ?? estado;
}

function writeSection(sheet: ExcelJS.Worksheet, row: number, title: string): number {
  const current = sheet.getRow(row);
  current.height = 20;
  merge(sheet, `A${row}:E${row}`, title, { bold: true, size: 12, color: { argb: BLACK } }, ORANGE_SOFT);
  return row + 1;
}

function writeTable(
  sheet: ExcelJS.Worksheet,
  startRow: number,
  headers: string[],
  data: (string | number)[][],
  total: (string | number)[]
): number {
  const header = sheet.getRow(startRow);
  header.height = 22;
  headers.forEach((label, index) => {
    const cell = header.getCell(index + 1);
    cell.value = label;
    cell.font = { bold: true, size: 10, color: { argb: WHITE } };
    cell.fill = fill(ORANGE);
    cell.alignment = {
      horizontal: index === 0 ? "left" : "right",
      vertical: "middle",
      wrapText: true,
    };
    cell.border = border();
  });

  data.forEach((values, idx) => {
    paintRow(sheet.getRow(startRow + 1 + idx), values, idx % 2 === 1 ? GREY : WHITE, false);
  });

  const totalRow = startRow + 1 + data.length;
  paintRow(sheet.getRow(totalRow), total, ORANGE_SOFT, true);
  return totalRow + 1;
}

function paintRow(
  row: ExcelJS.Row,
  values: (string | number)[],
  bg: string,
  bold: boolean
): void {
  row.height = 18;
  values.forEach((value, index) => {
    const cell = row.getCell(index + 1);
    cell.value = value;
    cell.font = { bold, size: 10, color: { argb: BLACK } };
    cell.fill = fill(bg);
    cell.alignment = { horizontal: index === 0 ? "left" : "right", vertical: "middle" };
    cell.border = border();
    if (index >= 2 && typeof value === "number") cell.numFmt = USD;
  });
}

function merge(
  sheet: ExcelJS.Worksheet,
  range: string,
  value: string,
  font: Partial<ExcelJS.Font>,
  bg: string
): void {
  sheet.mergeCells(range);
  const cell = sheet.getCell(range.split(":")[0]);
  cell.value = value;
  cell.font = font;
  cell.fill = fill(bg);
  cell.alignment = { horizontal: "center", vertical: "middle" };
}

function fill(argb: string): ExcelJS.Fill {
  return { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function border(): Partial<ExcelJS.Borders> {
  const side: Partial<ExcelJS.Border> = { style: "thin", color: { argb: "FFE7D3C4" } };
  return { top: side, left: side, bottom: side, right: side };
}

export function nombreArchivoQuincenas(reporte: ReporteOrdenes): string {
  const razones = [
    ...new Set(reporte.grupos.map((grupo) => grupo.empresa.trim()).filter(Boolean)),
  ];
  const fechas = reporte.grupos
    .flatMap((grupo) => [grupo.desde, grupo.hasta])
    .filter(Boolean)
    .sort();
  const desde = fechas[0] ?? "";
  const hasta = fechas[fechas.length - 1] ?? "";
  const fecha =
    desde && hasta ? `${fechaArchivo(desde)} al ${fechaArchivo(hasta)}` : "";
  const raw = ["quincenas", fecha, razones.join(", ")].filter(Boolean).join(" ");
  const limpio = raw.replace(/[\\/:*?"<>|]/g, " ").replace(/\s+/g, " ").trim();
  return `${limpio.slice(0, 180)}.xlsx`;
}

function fechaArchivo(iso: string): string {
  const [anio, mes, dia] = iso.split("-");
  if (!anio || !mes || !dia) return iso;
  return `${dia}-${mes}-${anio}`;
}

function formatRango(desde: string, hasta: string): string {
  const fmt = (iso: string) => {
    const [y, m, d] = iso.split("-");
    if (!y || !m || !d) return iso;
    return `${d}/${m}/${y}`;
  };
  if (desde && hasta) return `${fmt(desde)} al ${fmt(hasta)}`;
  return "";
}

function uniqueSheetName(empresa: string, used: Set<string>): string {
  const clean = empresa.replace(/[\\/*?:\[\]]/g, " ").trim() || "Ordenes";
  let name = clean.slice(0, 31);
  let n = 2;
  while (used.has(name)) {
    const suffix = ` ${n}`;
    name = `${clean.slice(0, 31 - suffix.length)}${suffix}`;
    n += 1;
  }
  used.add(name);
  return name;
}

function round2(n: number): number {
  return Math.round(n * 100) / 100;
}
