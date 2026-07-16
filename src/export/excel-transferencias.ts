import ExcelJS from "exceljs";
import { formatMontoUnificado } from "../utils/amounts";
import { ResumenTransferencias } from "../types";
import { TOLERANCIA_BS } from "../reconcile-transferencias";

const COLS = 11;
const GREY_HEADER = "FFE8E8E8";
const GREY_ROW = "FFF5F5F5";
const GREEN = "FF008000";
const ORANGE = "FFC05600";
const BLACK = "FF000000";
const WHITE = "FFFFFFFF";
const BLUE_SOFT = "FFEBF4FF";
const YELLOW_ALERT = "FFFFF3CD";

export async function exportTransferencias(
  resumen: ResumenTransferencias
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const headerRow = 4;
  const sheet = workbook.addWorksheet("Transferencias", {
    views: [{ state: "frozen", ySplit: headerRow }],
  });

  const lastCol = colLetter(COLS);

  sheet.getRow(1).height = 28;
  mergeAndStyle(sheet, `A1:${lastCol}1`, {
    value: "CONCILIADOR DE TRANSFERENCIAS",
    font: { bold: true, size: 14, color: { argb: BLACK } },
    fill: BLUE_SOFT,
    alignment: { horizontal: "center", vertical: "middle" },
  });

  sheet.getRow(2).height = 18;
  mergeAndStyle(sheet, `A2:${lastCol}2`, {
    value: `Tolerancia: hasta ${TOLERANCIA_BS} Bs · ✓ exacto · ⚠ con diferencia`,
    font: { size: 10, color: { argb: BLACK } },
    alignment: { horizontal: "center", vertical: "middle" },
  });

  sheet.getRow(3).height = 8;

  buildSummaryBox(sheet, headerRow, resumen);

  const headers = [
    "Fecha",
    "Sucursal",
    "Razón social",
    "Ticket",
    "Monto",
    "Banco",
    "Referencia",
    "Descripción",
    "Monto edo. cta.",
    "Dif.",
    "Conciliado",
  ];

  const header = sheet.getRow(headerRow);
  header.height = 20;
  headers.forEach((h, i) => {
    const cell = header.getCell(i + 1);
    cell.value = h;
    cell.font = { bold: true, size: 10, color: { argb: BLACK } };
    cell.fill = solidFill(GREY_HEADER);
    cell.alignment = {
      horizontal:
        i === 4 || i === 8 || i === 9 ? "right" : i === 10 ? "center" : "left",
      vertical: "middle",
    };
    cell.border = thinBorder();
  });

  sheet.autoFilter = { from: `A${headerRow}`, to: `K${headerRow}` };

  resumen.filas.forEach((fila, idx) => {
    const rowNum = headerRow + 1 + idx;
    const row = sheet.getRow(rowNum);
    row.height = 18;
    const isAlt = idx % 2 === 1;
    const rowFill = isAlt ? GREY_ROW : WHITE;

    let estadoIcon = "";
    if (fila.conciliado && fila.conDiferencia) {
      estadoIcon = "⚠";
    } else if (fila.conciliado) {
      estadoIcon = "✓";
    }

    const values: (string | number)[] = [
      fila.fecha,
      fila.sucursal,
      fila.razonSocial,
      fila.ticket,
      formatMontoUnificado(fila.montoTransferencia),
      fila.banco,
      fila.referencia,
      fila.descripcion,
      formatMontoUnificado(fila.montoEstadoCuenta),
      fila.conDiferencia && fila.diferencia !== null
        ? formatMontoUnificado(fila.diferencia)
        : "",
      estadoIcon,
    ];

    values.forEach((v, i) => {
      const cell = row.getCell(i + 1);
      cell.value = v;
      cell.border = thinBorder();
      cell.font = { size: 10, color: { argb: BLACK } };
      cell.fill = solidFill(rowFill);

      if (i === 4 || i === 8) {
        cell.alignment = { horizontal: "right", vertical: "middle" };
      } else if (i === 9) {
        cell.alignment = { horizontal: "right", vertical: "middle" };
        if (fila.conDiferencia) {
          cell.font = { bold: true, size: 10, color: { argb: ORANGE } };
        }
      } else if (i === 10) {
        cell.alignment = { horizontal: "center", vertical: "middle" };
        if (fila.conDiferencia) {
          cell.font = { bold: true, size: 13, color: { argb: ORANGE } };
          cell.fill = solidFill(YELLOW_ALERT);
        } else if (fila.conciliado) {
          cell.font = { bold: true, size: 13, color: { argb: GREEN } };
        }
      } else {
        cell.alignment = { horizontal: "left", vertical: "middle" };
      }
    });
  });

  sheet.columns = [
    { width: 12 },
    { width: 16 },
    { width: 20 },
    { width: 10 },
    { width: 12 },
    { width: 18 },
    { width: 14 },
    { width: 28 },
    { width: 14 },
    { width: 10 },
    { width: 11 },
    { width: 2 },
    { width: 22 },
    { width: 14 },
  ];

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

function buildSummaryBox(
  sheet: ExcelJS.Worksheet,
  startRow: number,
  resumen: ResumenTransferencias
) {
  const conDif = resumen.filas.filter((f) => f.conDiferencia).length;
  const exactas = resumen.conciliadas - conDif;

  const rows = [
    {
      label: "TRANSFERENCIAS",
      value: formatMontoUnificado(resumen.totalTransferencias),
      color: BLACK,
    },
    {
      label: "CONCILIADAS EN BANCO",
      value: formatMontoUnificado(resumen.totalEstadoCuenta),
      color: BLACK,
    },
    {
      label: "✓ / ⚠ / TOTAL",
      value: `${exactas} / ${conDif} / ${resumen.filas.length}`,
      color: resumen.sinConciliar > 0 || conDif > 0 ? ORANGE : GREEN,
    },
  ];

  rows.forEach((item, i) => {
    const r = startRow + i;
    const labelCell = sheet.getCell(`M${r}`);
    const valueCell = sheet.getCell(`N${r}`);

    labelCell.value = item.label;
    labelCell.font = { bold: true, size: 10 };
    labelCell.alignment = { horizontal: "left", vertical: "middle" };
    labelCell.border = mediumBorder();
    labelCell.fill = solidFill(i === 0 ? BLUE_SOFT : WHITE);

    valueCell.value = item.value;
    valueCell.font = { bold: true, size: 10, color: { argb: item.color } };
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
  const side: Partial<ExcelJS.Border> = {
    style: "thin",
    color: { argb: "FFBFBFBF" },
  };
  return { top: side, left: side, bottom: side, right: side };
}

function mediumBorder(): Partial<ExcelJS.Borders> {
  const side: Partial<ExcelJS.Border> = {
    style: "medium",
    color: { argb: BLACK },
  };
  return { top: side, left: side, bottom: side, right: side };
}

function colLetter(n: number): string {
  return String.fromCharCode(64 + n);
}
