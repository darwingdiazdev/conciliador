import express from "express";
import multer from "multer";
import path from "path";
import { parseVentas } from "./parsers/ventas";
import { parseMerchant } from "./parsers/merchant";
import { parseBanco } from "./parsers/banco";
import { reconcile } from "./reconcile";
import { COMPANIES } from "./constants/companies";
import { exportConciliacion } from "./export/excel";

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

app.listen(PORT, () => {
  console.log(`Conciliador disponible en http://localhost:${PORT}`);
});
