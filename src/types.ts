export interface VentaRow {
  fechaEmision: string;
  numeroFactura: number;
  empresa: string;
  cuotaInicialUsd: number;
  cuotaPendienteUsd: number;
  totalUsd: number;
  tasaCambio?: number;
}

export interface MerchantRow {
  fechaTransaccion: string;
  numeroFactura: number;
  montoPagadoBs: number;
  montoPagadoUsd: number;
  tasaCambio: number;
  referencia: string;
}

export interface BancoRow {
  fecha: string;
  referencia: string;
  haber: number;
  descripcion: string;
  tipoOperacion: string;
}

export interface BancoParseResult {
  movimientos: BancoRow[];
  saldoFinal: number;
}

export type TipoPagoCuota = "Pago completo" | "Abono";

export interface ConciliacionRow {
  fechaEmitida: string;
  numeroFactura: number;
  empresa: string;
  cuotaInicialUsd: number;
  cuotaPendienteUsd: number;
  totalUsd: number;
  casheaBs: number | null;
  bncBs: number | null;
  pagado: boolean;
  tipoPago: TipoPagoCuota | null;
}

export interface ResumenConciliacion {
  totalCashea: number;
  totalBnc: number;
  diferencia: number;
  filas: ConciliacionRow[];
}
