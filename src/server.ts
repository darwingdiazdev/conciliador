import express from "express";
import multer from "multer";
import path from "path";
import { parseVentas } from "./parsers/ventas";
import { parseMerchant } from "./parsers/merchant";
import { parseBanco } from "./parsers/banco";
import { parseEstadoCuenta } from "./parsers/estado-cuenta";
import { parseTransferencias } from "./parsers/transferencias";
import { reconcile } from "./reconcile";
import { reconcileTransferencias } from "./reconcile-transferencias";
import { COMPANIES } from "./constants/companies";
import { exportConciliacion } from "./export/excel";
import { exportTransferencias } from "./export/excel-transferencias";
import { detectArchivoKind } from "./utils/detect-file-kind";

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
const PORT = process.env.PORT || 3000;

app.use(express.static(path.join(__dirname, "../public")));

app.post(
  "/api/conciliar",
  upload.fields([
    { name: "ventas", maxCount: 1 },
    { name: "merchant", maxCount: 1 },
    { name: "banco", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      const files = req.files as Record<string, Express.Multer.File[]>;
      const ventasFile = files?.ventas?.[0];
      const merchantFile = files?.merchant?.[0];
      const bancoFile = files?.banco?.[0];

      if (!ventasFile || !merchantFile || !bancoFile) {
        res.status(400).json({
          error: "Debes cargar los 3 archivos: ventas, merchant y banco.",
        });
        return;
      }

      const [ventas, merchant, bancoData] = await Promise.all([
        parseVentas(ventasFile.buffer),
        parseMerchant(merchantFile.buffer),
        parseBanco(bancoFile.buffer),
      ]);

      const resumen = reconcile(
        ventas,
        merchant,
        bancoData.movimientos,
        bancoData.saldoFinal
      );

      const empresa = req.body.empresa as string;
      if (!empresa || !COMPANIES.includes(empresa as (typeof COMPANIES)[number])) {
        res.status(400).json({ error: "Selecciona una empresa válida." });
        return;
      }

      const fechaDesde = (req.body.fechaDesde as string) || undefined;
      const fechaHasta = (req.body.fechaHasta as string) || undefined;

      const buffer = await exportConciliacion(resumen, {
        empresa,
        fechaDesde,
        fechaHasta,
      });

      const filename = `conciliacion_${new Date().toISOString().slice(0, 10)}.xlsx`;

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
      res.send(buffer);
    } catch (err) {
      const message = err instanceof Error ? err.message : "Error desconocido";
      res.status(400).json({ error: message });
    }
  }
);

app.post(
  "/api/conciliar-transferencias",
  upload.fields([
    { name: "estadoCuenta", maxCount: 1 },
    { name: "transferencias", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      const files = req.files as Record<string, Express.Multer.File[]>;
      const estadoFile = files?.estadoCuenta?.[0];
      const transfFile = files?.transferencias?.[0];

      if (!estadoFile || !transfFile) {
        res.status(400).json({
          error:
            "Debes cargar los 2 archivos: estado de cuenta y transferencias del día.",
        });
        return;
      }

      // Autodetectar y corregir si los archivos vienen intercambiados
      let estadoBuf = estadoFile.buffer;
      let transfBuf = transfFile.buffer;
      const kindA = detectArchivoKind(estadoBuf);
      const kindB = detectArchivoKind(transfBuf);

      if (kindA === "transferencias" && kindB === "estadoCuenta") {
        estadoBuf = transfFile.buffer;
        transfBuf = estadoFile.buffer;
      } else if (kindA === "transferencias" && kindB !== "estadoCuenta") {
        res.status(400).json({
          error:
            'En "Estado de cuenta" subiste el archivo de transferencias. Sube el de movimientos (Fecha, Descripción, Monto, Saldo).',
        });
        return;
      } else if (kindB === "estadoCuenta" && kindA !== "transferencias") {
        res.status(400).json({
          error:
            'En "Transferencias" subiste el estado de cuenta. Sube el de transferencias del día (Sucursal, Ticket, Banco…).',
        });
        return;
      }

      const [estadoCuenta, transferencias] = await Promise.all([
        parseEstadoCuenta(estadoBuf),
        parseTransferencias(transfBuf),
      ]);

      const resumen = reconcileTransferencias(transferencias, estadoCuenta);
      const buffer = await exportTransferencias(resumen);

      const filename = `transferencias_${new Date().toISOString().slice(0, 10)}.xlsx`;

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
      res.send(buffer);
    } catch (err) {
      const message = err instanceof Error ? err.message : "Error desconocido";
      res.status(400).json({ error: message });
    }
  }
);

app.listen(PORT, () => {
  console.log(`Conciliador disponible en http://localhost:${PORT}`);
});
