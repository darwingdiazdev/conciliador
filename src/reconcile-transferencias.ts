import {
  EstadoCuentaRow,
  ResumenTransferencias,
  TransferenciaConciliacionRow,
  TransferenciaRow,
} from "./types";

/** Diferencia máxima permitida entre montos (Bs) para considerar conciliación. */
export const TOLERANCIA_BS = 20;

export function reconcileTransferencias(
  transferencias: TransferenciaRow[],
  estadoCuenta: EstadoCuentaRow[]
): ResumenTransferencias {
  // El banco exporta montos ÷1000 (3.12072 = 3,120.72). Alinear escala con transferencias.
  const estado = alignEstadoScale(transferencias, estadoCuenta);

  const usados = new Set<number>();
  const filas: TransferenciaConciliacionRow[] = [];

  for (const t of transferencias) {
    const match = findBestMatch(t.monto, estado, usados);

    if (match) {
      usados.add(match.index);
      const diferencia =
        Math.round((match.row.monto - t.monto) * 100) / 100;
      const conDiferencia = diferencia !== 0;
      filas.push({
        fecha: t.fecha,
        sucursal: t.sucursal,
        razonSocial: t.razonSocial,
        ticket: t.ticket,
        montoTransferencia: t.monto,
        banco: t.banco,
        referencia: t.referencia,
        descripcion: match.row.descripcion,
        montoEstadoCuenta: match.row.monto,
        diferencia: conDiferencia ? diferencia : 0,
        conciliado: true,
        conDiferencia,
      });
    } else {
      filas.push({
        fecha: t.fecha,
        sucursal: t.sucursal,
        razonSocial: t.razonSocial,
        ticket: t.ticket,
        montoTransferencia: t.monto,
        banco: t.banco,
        referencia: t.referencia,
        descripcion: "",
        montoEstadoCuenta: null,
        diferencia: null,
        conciliado: false,
        conDiferencia: false,
      });
    }
  }

  const conciliadas = filas.filter((f) => f.conciliado).length;
  const totalTransferencias = filas.reduce(
    (s, f) => s + f.montoTransferencia,
    0
  );
  const totalEstadoCuenta = filas
    .filter((f) => f.conciliado)
    .reduce((s, f) => s + (f.montoEstadoCuenta ?? 0), 0);

  return {
    totalTransferencias,
    totalEstadoCuenta,
    conciliadas,
    sinConciliar: filas.length - conciliadas,
    filas,
  };
}

/**
 * Si la mediana del estado de cuenta es ~1000× menor que la de transferencias,
 * los montos del banco vienen en miles y hay que multiplicar por 1000.
 */
function alignEstadoScale(
  transferencias: TransferenciaRow[],
  estadoCuenta: EstadoCuentaRow[]
): EstadoCuentaRow[] {
  const tMed = median(transferencias.map((t) => t.monto).filter((n) => n > 0));
  const eMed = median(estadoCuenta.map((e) => e.monto).filter((n) => n > 0));
  if (!tMed || !eMed) return estadoCuenta;

  const ratio = tMed / eMed;
  if (ratio < 50 || ratio > 5000) return estadoCuenta;

  return estadoCuenta.map((e) => ({
    ...e,
    monto: Math.round(e.monto * 1000 * 100) / 100,
    saldo:
      e.saldo && !Number.isNaN(e.saldo)
        ? Math.round(e.saldo * 1000 * 100) / 100
        : e.saldo,
  }));
}

function median(values: number[]): number {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0
    ? (sorted[mid - 1] + sorted[mid]) / 2
    : sorted[mid];
}

function findBestMatch(
  monto: number,
  estadoCuenta: EstadoCuentaRow[],
  usados: Set<number>
): { index: number; row: EstadoCuentaRow; diff: number } | null {
  let best: { index: number; row: EstadoCuentaRow; diff: number } | null =
    null;

  for (let i = 0; i < estadoCuenta.length; i++) {
    if (usados.has(i)) continue;
    const row = estadoCuenta[i];
    const diff = Math.abs(row.monto - monto);
    if (diff > TOLERANCIA_BS) continue;

    if (!best || diff < best.diff) {
      best = { index: i, row, diff };
      if (diff === 0) break;
    }
  }

  return best;
}
