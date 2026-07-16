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
  const usados = new Set<number>();
  const filas: TransferenciaConciliacionRow[] = [];

  for (const t of transferencias) {
    const match = findBestMatch(t.monto, estadoCuenta, usados);

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
