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

export interface EstadoCuentaRow {
  fecha: string;
  descripcion: string;
  monto: number;
  saldo: number;
}

export interface TransferenciaRow {
  fecha: string;
  sucursal: string;
  razonSocial: string;
  ticket: string;
  monto: number;
  banco: string;
  referencia: string;
}

export interface TransferenciaConciliacionRow {
  fecha: string;
  sucursal: string;
  razonSocial: string;
  ticket: string;
  montoTransferencia: number;
  banco: string;
  referencia: string;
  descripcion: string;
  montoEstadoCuenta: number | null;
  diferencia: number | null;
  /** Exacto = true; con diferencia dentro de tolerancia = true pero con alerta en Excel */
  conciliado: boolean;
  /** true si se emparejó pero con diferencia (≤ 20 Bs) */
  conDiferencia: boolean;
}

export interface ResumenTransferencias {
  totalTransferencias: number;
  totalEstadoCuenta: number;
  conciliadas: number;
  sinConciliar: number;
  filas: TransferenciaConciliacionRow[];
}
