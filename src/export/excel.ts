import fs from "fs";
import path from "path";
import ExcelJS from "exceljs";
import { COMPANIES } from "../constants/companies";
import { formatBs, formatUsd } from "../utils/amounts";
import { ResumenConciliacion } from "../types";

const COLS = 10;
const YELLOW_BANNER = "FFFFFF00";
const YELLOW_CASHEA = "FFFFF2CC"; // amarillo claro
const GREY_HEADER = "FFE8E8E8";
const GREY_ROW = "FFF5F5F5";
const GREEN = "FF008000";
const RED = "FFFF0000";
const BLACK = "FF000000";
const WHITE = "FFFFFFFF";

const CASHEA_BANNER = "Gemini_Generated_Image_6aycd26aycd26ayc (1).png";
const BNC_LOGO = "BNC.png";

export async function exportConciliacion(
  resumen: ResumenConciliacion,
  options: { empresa?: string; fechaDesde?: string; fechaHasta?: string } = {}
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const headerRow = 4;
  const sheet = workbook.addWorksheet("Pagos conciliados", {
    views: [{ state: "frozen", ySplit: headerRow }],
  });

  const empresa =
    options.empresa ||
    resumen.filas.find((f) => f.empresa)?.empresa ||
    COMPANIES[COMPANIES.length - 1];

  const fechas = resumen.filas.map((f) => f.fechaEmitida).filter(Boolean).sort();
  const desde = options.fechaDesde || fechas[0] || "";
  const hasta = options.fechaHasta || fechas[fechas.length - 1] || "";
  const rangoTitulo = formatRangoFechas(desde, hasta);
  const lastCol = colLetter(COLS);

  // ── Fila 1: banner con imágenes ──
  sheet.getRow(1).height = 52;
  mergeAndStyle(sheet, `A1:${lastCol}1`, {
    value: "",
    fill: YELLOW_BANNER,
  });
  addHeaderImages(workbook, sheet);

  // ── Fila 2: empresa ──
  sheet.getRow(2).height = 22;
  mergeAndStyle(sheet, `A2:${lastCol}2`, {
    value: empresa.toUpperCase(),
    font: { bold: true, size: 11, color: { argb: BLACK } },
    fill: GREY_HEADER,
    alignment: { horizontal: "center", vertical: "middle" },
  });

  // ── Fila 3: fechas ──
  sheet.getRow(3).height = 18;
  mergeAndStyle(sheet, `A3:${lastCol}3`, {
    value: rangoTitulo,
    font: { bold: true, size: 10, color: { argb: BLACK } },
    alignment: { horizontal: "center", vertical: "middle" },
  });

  // ── Resumen + encabezados (fila 4) ──
  buildSummaryBox(sheet, headerRow, resumen);

  const headers = [
    "Fecha emitida",
    "N° Factura",
    "Empresa",
    "C. inicial $",
    "C. pendiente $",
    "Total $",
    "Cashea",
    "BNC",
    "Pagado",
    "Tipo pago",
  ];

  const header = sheet.getRow(headerRow);
  header.height = 20;
  headers.forEach((h, i) => {
    const cell = header.getCell(i + 1);
    cell.value = h;
    cell.font = { bold: true, size: 10, color: { argb: BLACK } };
    cell.fill = solidFill(GREY_HEADER);
    cell.alignment = {
      horizontal: i >= 3 && i <= 7 ? "right" : i === 9 ? "center" : "left",
      vertical: "middle",
    };
    cell.border = thinBorder();
  });

  sheet.autoFilter = { from: `A${headerRow}`, to: `J${headerRow}` };

  // ── Datos ──
  resumen.filas.forEach((fila, idx) => {
    const rowNum = headerRow + 1 + idx;
    const row = sheet.getRow(rowNum);
    row.height = 18;
    const isAlt = idx % 2 === 1;
    const rowFill = isAlt ? GREY_ROW : WHITE;

    const values: (string | number)[] = [
      fila.fechaEmitida,
      fila.numeroFactura,
      fila.empresa,
      formatUsd(fila.cuotaInicialUsd),
      formatUsd(fila.cuotaPendienteUsd),
      formatUsd(fila.totalUsd),
      formatBs(fila.casheaBs),
      formatBs(fila.bncBs),
      fila.pagado ? "✓" : "",
      fila.tipoPago ?? "",
    ];

    values.forEach((v, i) => {
      const cell = row.getCell(i + 1);
      cell.value = v;
      cell.border = thinBorder();
      cell.font = { size: 10, color: { argb: BLACK } };

      if (i === 6) {
        cell.fill = solidFill(YELLOW_CASHEA);
        cell.alignment = { horizontal: "right", vertical: "middle" };
      } else if (i === 8) {
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.fill = solidFill(rowFill);
        if (fila.pagado) {
          cell.font = { bold: true, size: 13, color: { argb: GREEN } };
        }
      } else if (i === 9) {
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.fill = solidFill(rowFill);
        if (fila.tipoPago === "Pago completo") {
          cell.font = { bold: true, color: { argb: GREEN } };
        } else if (fila.tipoPago === "Abono") {
          cell.font = { bold: true, color: { argb: RED } };
        }
      } else {
        cell.fill = solidFill(rowFill);
        cell.alignment = {
          horizontal: i >= 3 && i <= 7 ? "right" : "left",
          vertical: "middle",
        };
      }
    });
  });

  sheet.columns = [
    { width: 14 },
    { width: 11 },
    { width: 32 },
    { width: 13 },
    { width: 13 },
    { width: 11 },
    { width: 17 },
    { width: 17 },
    { width: 9 },
    { width: 14 },
    { width: 2 },
    { width: 22 },
    { width: 18 },
  ];

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

function addHeaderImages(workbook: ExcelJS.Workbook, sheet: ExcelJS.Worksheet) {
  const casheaPath = assetPath(CASHEA_BANNER);
  const bncPath = assetPath(BNC_LOGO);

  if (fs.existsSync(casheaPath)) {
    const casheaImg = workbook.addImage({
      buffer: fs.readFileSync(casheaPath) as unknown as ExcelJS.Buffer,
      extension: "png",
    });
    sheet.addImage(casheaImg, {
      tl: { col: 0.05, row: 0.08 },
      ext: { width: 380, height: 44 },
    });
  }

  if (fs.existsSync(bncPath)) {
    const bncImg = workbook.addImage({
      buffer: fs.readFileSync(bncPath) as unknown as ExcelJS.Buffer,
      extension: "png",
    });
    sheet.addImage(bncImg, {
      tl: { col: 6.2, row: 0.15 },
      ext: { width: 200, height: 40 },
    });
  }
}

function assetPath(filename: string): string {
  return path.join(__dirname, "../../public/assets", filename);
}

function buildSummaryBox(
  sheet: ExcelJS.Worksheet,
  startRow: number,
  resumen: ResumenConciliacion
) {
  const rows = [
    { label: "INGRESOS POR CASHEA", value: formatBs(resumen.totalCashea), highlight: true },
    { label: "TOTAL EN CUENTA", value: formatBs(resumen.totalBnc), highlight: false },
    {
      label: "DIF.",
      value: formatBs(resumen.diferencia),
      highlight: false,
      valueColor: resumen.diferencia < 0 ? RED : GREEN,
    },
  ];

  rows.forEach((item, i) => {
    const r = startRow + i;
    const labelCell = sheet.getCell(`K${r}`);
    const valueCell = sheet.getCell(`L${r}`);

    labelCell.value = item.label;
    labelCell.font = { bold: true, size: 10 };
    labelCell.alignment = { horizontal: "left", vertical: "middle" };
    labelCell.border = mediumBorder();
    labelCell.fill = solidFill(item.highlight ? YELLOW_CASHEA : WHITE);

    valueCell.value = item.value;
    valueCell.font = {
      bold: true,
      size: 10,
      color: { argb: item.valueColor ?? BLACK },
    };
    valueCell.alignment = { horizontal: "right", vertical: "middle" };
    valueCell.border = mediumBorder();
    valueCell.fill = solidFill(WHITE);
  });
}

function mergeAndStyle(
  sheet: ExcelJS.Worksheet,
  range: string,
  opts: {
    value?: string;
    font?: Partial<ExcelJS.Font>;
    fill?: string;
    alignment?: Partial<ExcelJS.Alignment>;
  }
) {
  sheet.mergeCells(range);
  const cell = sheet.getCell(range.split(":")[0]);
  if (opts.value !== undefined) cell.value = opts.value;
  if (opts.font) cell.font = opts.font;
  if (opts.fill) cell.fill = solidFill(opts.fill);
  if (opts.alignment) cell.alignment = opts.alignment;
}

function solidFill(argb: string): ExcelJS.Fill {
  return { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function thinBorder(): Partial<ExcelJS.Borders> {
  const side: Partial<ExcelJS.Border> = { style: "thin", color: { argb: "FFBFBFBF" } };
  return { top: side, left: side, bottom: side, right: side };
}

function mediumBorder(): Partial<ExcelJS.Borders> {
  const side: Partial<ExcelJS.Border> = { style: "medium", color: { argb: BLACK } };
  return { top: side, left: side, bottom: side, right: side };
}

function colLetter(n: number): string {
  return String.fromCharCode(64 + n);
}

function formatRangoFechas(desde: string, hasta: string): string {
  const fmt = (d: string) => {
    const parts = d.split("-");
    if (parts.length === 3) return `${parts[2]}-${parts[1]}`;
    return d;
  };
  if (desde && hasta) return `${fmt(desde)} al ${fmt(hasta)}`;
  return "Conciliación de pagos";
}
