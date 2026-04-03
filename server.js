const express = require("express");
const multer = require("multer");
const path = require("path");
const { buildPendingDownloadWorkbook, processInvoiceBundle } = require("./processor");

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
const port = process.env.PORT || 3000;
const MAX_INVOICE_UPLOAD_COUNT = 300;

app.use(express.static(__dirname));

app.post(
  "/api/process",
  upload.fields([
    { name: "invoices", maxCount: MAX_INVOICE_UPLOAD_COUNT },
    { name: "fullExport", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      const invoiceFiles = req.files?.invoices || [];
      const fullExportFile = req.files?.fullExport?.[0];

      if (!invoiceFiles.length) {
        return res.status(400).json({ error: "请至少上传一份 PDF 发票。" });
      }
      if (!fullExportFile) {
        return res.status(400).json({ error: "请上传“全量发票查询导出结果.xlsx”。" });
      }

      const result = await processInvoiceBundle({
        invoiceFiles,
        fullExportFile,
      });

      res.json({
        summary: result.summary,
        templateHeaders: result.templateHeaders,
        matchedRows: result.matchedRows,
        ordinaryInvoices: result.ordinaryInvoices.map((item) => ({
          fileName: item.fileName,
          invoiceType: item.invoiceType,
          digitalInvoiceNumber: item.digitalInvoiceNumber,
          issueDate: item.issueDate,
          sellerName: item.sellerName,
          grossAmount: item.grossAmount,
        })),
        parsedInvoices: result.parsedInvoices.map((item) => ({
          fileName: item.fileName,
          digitalInvoiceNumber: item.digitalInvoiceNumber,
          issueDate: item.issueDate,
          sellerName: item.sellerName,
          grossAmount: item.grossAmount,
        })),
        errors: result.errors.map((item) => ({
          fileIndex: item.fileIndex ?? null,
          fileName: item.fileName,
          reason: item.reason,
          digitalInvoiceNumber: item.invoice?.digitalInvoiceNumber || "",
          sellerName: item.invoice?.sellerName || "",
          issueDate: item.invoice?.issueDate || "",
          grossAmount: item.invoice?.grossAmount ?? null,
        })),
        workbookBase64: result.workbookBuffer.toString("base64"),
        outputFileName: buildOutputFileName(),
      });
    } catch (error) {
      res.status(500).json({ error: error.message || "处理失败" });
    }
  }
);

app.post(
  "/api/pending-download",
  upload.fields([
    { name: "invoices", maxCount: MAX_INVOICE_UPLOAD_COUNT },
    { name: "fullExport", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      const invoiceFiles = req.files?.invoices || [];
      const fullExportFile = req.files?.fullExport?.[0];

      if (!invoiceFiles.length) {
        return res.status(400).json({ error: "请至少上传一份 PDF 发票。" });
      }
      if (!fullExportFile) {
        return res.status(400).json({ error: "请上传“全量发票查询导出结果.xlsx”。" });
      }

      const result = await buildPendingDownloadWorkbook({
        invoiceFiles,
        fullExportFile,
      });

      res.json({
        pendingCount: result.pendingCount,
        workbookBase64: result.workbookBuffer.toString("base64"),
        outputFileName: buildPendingOutputFileName(),
      });
    } catch (error) {
      res.status(500).json({ error: error.message || "导入清单模板生成失败" });
    }
  }
);

app.use((error, req, res, next) => {
  if (!error) {
    return next();
  }

  if (error instanceof multer.MulterError) {
    if (error.code === "LIMIT_UNEXPECTED_FILE" && error.field === "invoices") {
      return res.status(400).json({
        error: `PDF 发票最多可上传 ${MAX_INVOICE_UPLOAD_COUNT} 份，当前已超出上限。`,
      });
    }

    return res.status(400).json({
      error: error.message || "上传文件失败",
    });
  }

  return res.status(500).json({
    error: error.message || "服务器内部错误",
  });
});

app.get(/.*/, (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

app.listen(port, () => {
  console.log(`digital-invoice-organizer running at http://localhost:${port}`);
});

function buildOutputFileName() {
  const now = new Date();
  const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
  return `抵扣勾选增值税发票信息_${stamp}.xlsx`;
}

function buildPendingOutputFileName() {
  const now = new Date();
  const target = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
  return `导入清单模板_${target.getFullYear()}${String(target.getMonth() + 1).padStart(2, "0")}_${stamp}.xlsx`;
}
