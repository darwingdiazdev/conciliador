export function parseMonto(value: unknown): number {
  if (value === null || value === undefined || value === "") return NaN;
  if (typeof value === "number") return value;
  const str = String(value)
    .replace(/[+]/g, "")
    .replace(/,/g, "")
    .replace(/Bs\.?S\.?/gi, "")
    .replace(/\$/g, "")
    .trim();
  return parseFloat(str);
}

/**
 * Unifica montos de ambos formatos a un número comparable:
 *   - Europeo / VE: 3.250,00  →  3250.00
 *   - US:           3,250.00  →  3250.00
 *   - Solo miles EU: 3.250    →  3250.00
 * Resultado canónico: miles con coma, decimales con punto (3,250.00).
 */
export function unificarMonto(value: unknown): number {
  const n = parseMontoFlexible(value);
  if (Number.isNaN(n)) return NaN;
  return Math.round(n * 100) / 100;
}

/** Parsea montos en formato VE/EU (3.250,00) o US (3,250.00 / 3250.00). */
export function parseMontoFlexible(value: unknown): number {
  if (value === null || value === undefined || value === "") return NaN;
  if (typeof value === "number") return value;

  let str = String(value)
    .replace(/\u00a0/g, " ")
    .replace(/[+]/g, "")
    .replace(/Bs\.?S\.?/gi, "")
    .replace(/\$/g, "")
    .trim();

  if (!str) return NaN;

  // Negativos con paréntesis: (1.250,00)
  let negative = false;
  if (/^\(.*\)$/.test(str)) {
    negative = true;
    str = str.slice(1, -1).trim();
  }
  if (str.startsWith("-")) {
    negative = true;
    str = str.slice(1).trim();
  }

  const hasComma = str.includes(",");
  const hasDot = str.includes(".");

  if (hasComma && hasDot) {
    const lastComma = str.lastIndexOf(",");
    const lastDot = str.lastIndexOf(".");
    if (lastComma > lastDot) {
      // 3.250,00 → europeo
      str = str.replace(/\./g, "").replace(",", ".");
    } else {
      // 3,250.00 → US
      str = str.replace(/,/g, "");
    }
  } else if (hasComma && !hasDot) {
    // Solo comas: 3250,00 (decimal EU) o 3,250 (miles US)
    const parts = str.split(",");
    if (parts.length === 2 && parts[1].length <= 2) {
      str = `${parts[0].replace(/\./g, "")}.${parts[1]}`;
    } else {
      str = str.replace(/,/g, "");
    }
  } else if (hasDot && !hasComma) {
    // Varios grupos de miles europeos: 1.250.000
    if (/^\d{1,3}(\.\d{3}){2,}$/.test(str)) {
      str = str.replace(/\./g, "");
    }
    // Un solo punto (3.121 / 3250.00 / 1.25) → decimal US; NO tratar como miles
  }

  const n = parseFloat(str);
  if (Number.isNaN(n)) return NaN;
  return negative ? -n : n;
}

/** Formato canónico: 3,250.00 (coma miles, punto decimales). */
export function formatMontoUnificado(amount: number | null): string {
  if (amount === null || Number.isNaN(amount)) return "";
  return amount.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

export function amountsMatch(a: number, b: number, epsilon = 0.02): boolean {
  if (Number.isNaN(a) || Number.isNaN(b)) return false;
  return Math.abs(a - b) <= epsilon;
}

export function roundMontoKey(amount: number): string {
  return amount.toFixed(2);
}

export function formatBs(amount: number | null): string {
  if (amount === null || Number.isNaN(amount)) return "";
  return `Bs.S ${amount.toLocaleString("es-VE", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`;
}

export function amountsMatchUsd(a: number, b: number): boolean {
  const ra = Math.round(a * 100) / 100;
  const rb = Math.round(b * 100) / 100;
  return amountsMatch(ra, rb, 0.05);
}

export function formatUsd(amount: number | null): string {
  if (amount === null || Number.isNaN(amount)) return "";
  return `$ ${amount.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`;
}
