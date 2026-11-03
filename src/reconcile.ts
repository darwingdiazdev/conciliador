import { amountsMatch, amountsMatchUsd, roundMontoKey } from "./utils/amounts";
import {
  BancoRow,
  ConciliacionRow,
  MerchantRow,
  PendienteConfirmarRow,
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
  const merchantByFactura = new Map<number, MerchantRow[]>();
  for (const m of merchant) {
    const list = merchantByFactura.get(m.numeroFactura) ?? [];
    list.push(m);
    merchantByFactura.set(m.numeroFactura, list);
  }

  const bancoByMonto = new Map<string, BancoRow[]>();
  for (const b of banco) {
    const key = roundMontoKey(b.haber);
    const list = bancoByMonto.get(key) ?? [];
    list.push(b);
    bancoByMonto.set(key, list);
  }

  const bancosUsados = new Set<BancoRow>();
  const merchantConciliados = new Set<MerchantRow>();
  const filas: ConciliacionRow[] = [];

  for (const venta of ventas) {
    const pagos = [...(merchantByFactura.get(venta.numeroFactura) ?? [])].sort(
      (a, b) =>
        a.fechaTransaccion.localeCompare(b.fechaTransaccion) ||
        a.referencia.localeCompare(b.referencia) ||
        a.montoPagadoBs - b.montoPagadoBs
    );

    for (const m of pagos) {
      const casheaBs = m.montoPagadoBs > 0 ? m.montoPagadoBs : null;
      let bncBs: number | null = null;

      if (casheaBs !== null) {
        const candidatos = findBancoCandidates(casheaBs, bancoByMonto);
        const match = candidatos.find(
          (b) => !bancosUsados.has(b) && amountsMatch(b.haber, casheaBs)
        );
        if (match) {
          bncBs = match.haber;
          bancosUsados.add(match);
          merchantConciliados.add(m);
        }
      }

      filas.push({
        fechaEmitida: venta.fechaEmision,
        fechaPago: m.fechaTransaccion,
        numeroFactura: venta.numeroFactura,
        empresa: venta.empresa,
        cuotaInicialUsd: venta.cuotaInicialUsd,
        cuotaPendienteUsd: venta.cuotaPendienteUsd,
        totalUsd: venta.totalUsd,
        casheaBs,
        bncBs,
        pagado: casheaBs !== null && bncBs !== null && amountsMatch(casheaBs, bncBs),
        tipoPago: classifyTipoPago(m, casheaBs, venta),
      });
    }
  }

  // Ingresos Cashea = solo ventas conciliadas (✓)
  const totalCashea = filas
    .filter((r) => r.pagado)
    .reduce((s, r) => s + (r.casheaBs ?? 0), 0);

  // Total en cuenta = saldo final del extracto BNC
  const totalBnc = saldoFinal;

  const facturasVenta = new Set(ventas.map((v) => v.numeroFactura));
  const pendientesConfirmar: PendienteConfirmarRow[] = [];
  const merchantSinFactura = merchant
    .filter((m) => !facturasVenta.has(m.numeroFactura) && m.montoPagadoBs > 0)
    .sort(
      (a, b) =>
        a.fechaTransaccion.localeCompare(b.fechaTransaccion) ||
        a.numeroFactura - b.numeroFactura ||
        a.referencia.localeCompare(b.referencia)
    );

  for (const m of merchantSinFactura) {
    const match = findBancoCandidates(m.montoPagadoBs, bancoByMonto).find(
      (b) => !bancosUsados.has(b) && amountsMatch(b.haber, m.montoPagadoBs)
    );
    if (!match) continue;

    bancosUsados.add(match);
    merchantConciliados.add(m);
    pendientesConfirmar.push({
      fechaMerchant: m.fechaTransaccion,
      numeroFactura: m.numeroFactura,
      referenciaMerchant: m.referencia,
      montoMerchantBs: m.montoPagadoBs,
      montoUsd: m.montoPagadoUsd,
      tasaCambio: m.tasaCambio,
      fechaBanco: match.fecha,
      referenciaBanco: match.referencia,
      descripcionBanco: match.descripcion,
      montoBanco: match.haber,
    });
  }

  const bancoSinConciliar = banco
    .filter((b) => !bancosUsados.has(b))
    .sort((a, b) => a.fecha.localeCompare(b.fecha) || a.referencia.localeCompare(b.referencia));
  const merchantSinVenta = merchant
    .filter((m) => !facturasVenta.has(m.numeroFactura) && !merchantConciliados.has(m))
    .sort(
      (a, b) =>
        a.fechaTransaccion.localeCompare(b.fechaTransaccion) ||
        a.numeroFactura - b.numeroFactura
    );
  const merchantSinBanco = merchant
    .filter((m) => facturasVenta.has(m.numeroFactura) && !merchantConciliados.has(m))
    .sort(
      (a, b) =>
        a.numeroFactura - b.numeroFactura ||
        a.fechaTransaccion.localeCompare(b.fechaTransaccion) ||
        a.referencia.localeCompare(b.referencia)
    );

  return {
    totalCashea,
    totalBnc,
    diferencia: totalBnc - totalCashea,
    filas,
    bancoSinConciliar,
    merchantSinVenta,
    merchantSinBanco,
    pendientesConfirmar,
  };
}

function classifyTipoPago(
  merchant: MerchantRow | undefined,
  casheaBs: number | null,
  venta: VentaRow
): TipoPagoCuota | null {
  if (!merchant || casheaBs === null || casheaBs <= 0) return null;

  const cuotaPendienteUsd = venta.cuotaPendienteUsd;
  if (!cuotaPendienteUsd || cuotaPendienteUsd <= 0) return null;

  const pagadoUsd = resolvePagadoUsd(merchant, casheaBs);
  if (pagadoUsd === null) return null;

  if (amountsMatchUsd(pagadoUsd, cuotaPendienteUsd)) return "Pago completo";
  return "Abono";
}

function resolvePagadoUsd(
  merchant: MerchantRow,
  casheaBs: number
): number | null {
  if (merchant.montoPagadoUsd > 0) {
    return merchant.montoPagadoUsd;
  }
  if (merchant.tasaCambio > 0 && casheaBs > 0) {
    return casheaBs / merchant.tasaCambio;
  }
  return null;
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
