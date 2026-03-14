import type { Express } from "express";
import * as ExcelJS from "exceljs";

type ExcelJSImportLike = { Workbook: new () => ExcelJS.Workbook };
const ExcelJSRuntime: ExcelJSImportLike =
  (ExcelJS as unknown as { default?: ExcelJSImportLike }).default ??
  (ExcelJS as unknown as ExcelJSImportLike);

export const registerExportRoute = (app: Express) => {
  app.post("/api/files/export", async (req, res) => {
    const { fileName, headers, rows } = req.body as {
      fileName: unknown;
      headers: unknown;
      rows: unknown;
    };

    if (typeof fileName !== "string" || fileName.trim().length === 0) {
      return res
        .status(400)
        .json({ message: "fileName must be a non-empty string" });
    }
    if (
      !Array.isArray(headers) ||
      !headers.every((item) => typeof item === "string")
    ) {
      return res.status(400).json({ message: "headers must be a string array" });
    }
    if (
      !Array.isArray(rows) ||
      !rows.every(
        (row) =>
          Array.isArray(row) && row.every((cell) => typeof cell === "string"),
      )
    ) {
      return res.status(400).json({ message: "rows must be a 2d string array" });
    }

    try {
      const workbook = new ExcelJSRuntime.Workbook();
      const worksheet = workbook.addWorksheet("Sheet1");

      worksheet.addRow(headers);
      for (const row of rows) {
        worksheet.addRow(row);
      }

      worksheet.columns = headers.map((header, index) => {
        const maxLengthFromRows = rows.reduce((acc, row) => {
          const value = row[index] ?? "";
          return Math.max(acc, value.length);
        }, 0);
        return {
          header,
          key: `col_${index}`,
          width: Math.min(
            60,
            Math.max(12, Math.max(header.length, maxLengthFromRows) + 2),
          ),
        };
      });

      const headerRow = worksheet.getRow(1);
      headerRow.font = { bold: true };
      headerRow.commit();

      const baseName = fileName.replace(/\.[^.]+$/, "");
      const exportName = `${baseName}-导出.xlsx`;
      const encodedFileName = encodeURIComponent(exportName);
      const buffer = await workbook.xlsx.writeBuffer();

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      );
      res.setHeader(
        "Content-Disposition",
        `attachment; filename*=UTF-8''${encodedFileName}`,
      );
      const outputBuffer = Buffer.isBuffer(buffer)
        ? buffer
        : Buffer.from(buffer as ArrayBuffer);
      return res.send(outputBuffer);
    } catch (error) {
      const message = error instanceof Error ? error.message : "导出 Excel 失败";
      return res.status(500).json({ message });
    }
  });
};
