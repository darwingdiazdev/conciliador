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

/** Parsea montos en formato VE/EU (1.250,00) o US (1,250.00 / 1250.00). */
export function parseMontoFlexible(value: unknown): number {
  if (value === null || value === undefined || value === "") return NaN;
  if (typeof value === "number") return value;

  let str = String(value)
    .replace(/[+]/g, "")
    .replace(/Bs\.?S\.?/gi, "")
    .replace(/\$/g, "")
    .trim();

  const lastComma = str.lastIndexOf(",");
  const lastDot = str.lastIndexOf(".");

  if (lastComma > lastDot) {
    // Formato europeo: 1.250,00
    str = str.replace(/\./g, "").replace(",", ".");
  } else {
    // Formato US o sin miles: 1,250.00 / 1250.00
    str = str.replace(/,/g, "");
  }

  return parseFloat(str);
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
