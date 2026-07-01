import { amountsMatch, roundMontoKey } from "./utils/amounts";
import {
  BancoRow,
  ConciliacionRow,
  MerchantRow,
  ResumenConciliacion,
  TipoPagoCuota,
  VentaRow,
} from "./types";

export function reconcile(
  ventas: VentaRow[],
  merchant: MerchantRow[],
  banco: BancoRow[],
  saldoFinal: number
): ResumenConciliacion {
  const merchantByFactura = new Map<number, MerchantRow>();
  for (const m of merchant) {
    merchantByFactura.set(m.numeroFactura, m);
  }

  const bancoByMonto = new Map<string, BancoRow[]>();
  for (const b of banco) {
    const key = roundMontoKey(b.haber);
    const list = bancoByMonto.get(key) ?? [];
    list.push(b);
    bancoByMonto.set(key, list);
  }

  const usadosBanco = new Set<string>();
  const filas: ConciliacionRow[] = [];

  for (const venta of ventas) {
    const m = merchantByFactura.get(venta.numeroFactura);
    const casheaBs = m?.montoPagadoBs ?? null;

    let bncBs: number | null = null;

    if (m && casheaBs !== null) {
      const candidatos = findBancoCandidates(casheaBs, bancoByMonto);
      const match = candidatos.find(
        (b) => !usadosBanco.has(b.referencia) && amountsMatch(b.haber, casheaBs)
      );
      if (match) {
        bncBs = match.haber;
        usadosBanco.add(match.referencia);
      }
    }

    filas.push({
      fechaEmitida: venta.fechaEmision,
      numeroFactura: venta.numeroFactura,
      empresa: venta.empresa,
      cuotaInicialUsd: venta.cuotaInicialUsd,
      cuotaPendienteUsd: venta.cuotaPendienteUsd,
      totalUsd: venta.totalUsd,
      casheaBs,
      bncBs,
      pagado: casheaBs !== null && bncBs !== null && amountsMatch(casheaBs, bncBs),
      tipoPago: classifyTipoPago(m, casheaBs, venta.cuotaPendienteUsd),
    });
  }

  // Ingresos Cashea = solo ventas conciliadas (✓)
  const totalCashea = filas
    .filter((r) => r.pagado)
    .reduce((s, r) => s + (r.casheaBs ?? 0), 0);

  // Total en cuenta = saldo final del extracto BNC
  const totalBnc = saldoFinal;

  return {
    totalCashea,
    totalBnc,
    diferencia: totalBnc - totalCashea,
    filas,
  };
}

function classifyTipoPago(
  merchant: MerchantRow | undefined,
  casheaBs: number | null,
  cuotaPendienteUsd: number
): TipoPagoCuota | null {
  if (!merchant || casheaBs === null || casheaBs <= 0) return null;
  if (!merchant.tasaCambio || merchant.tasaCambio <= 0) return null;

  const pagadoUsd = casheaBs / merchant.tasaCambio;
  if (amountsMatch(pagadoUsd, cuotaPendienteUsd)) return "Pago completo";
  return "Abono";
}

function findBancoCandidates(
  monto: number,
  bancoByMonto: Map<string, BancoRow[]>
): BancoRow[] {
  const exact = bancoByMonto.get(roundMontoKey(monto)) ?? [];
  if (exact.length > 0) return exact;

  const all: BancoRow[] = [];
  bancoByMonto.forEach((rows) => all.push(...rows));
  return all.filter((b) => amountsMatch(b.haber, monto));
}
