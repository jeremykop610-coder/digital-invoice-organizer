const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const { PDFParse } = require("pdf-parse");

const TEMPLATE_DEFAULTS = {
  "是否勾选*": "是",
  "红字锁定标志": "未锁定",
  "发票来源": "电子发票服务平台",
};

const TEMPLATE_DIR_CANDIDATES = ["samples-private", "samples"];
const DEFAULT_TEMPLATE_PATH = resolveTemplatePath("抵扣勾选增值税发票信息.xlsx");
const IMPORT_TEMPLATE_PATH = resolveTemplatePath("导入清单模板.xlsx");

async function processInvoiceBundle({ invoiceFiles, fullExportFile }) {
  const fullExportWorkbook = XLSX.read(fullExportFile.buffer, { type: "buffer", cellDates: false });
  const templateWorkbook = loadTemplateWorkbook();

  const fullExportRows = mergeWorkbookRows(fullExportWorkbook);
  const exportIndex = buildExportIndex(fullExportRows);
  const redFlushContext = buildRedFlushContext(fullExportRows);
  const templateSheet = templateWorkbook.Sheets[templateWorkbook.SheetNames[0]];
  const templateHeaders = readTemplateHeaders(templateSheet);

  if (!templateHeaders.length) {
    throw new Error("未能读取“抵扣勾选增值税发票信息”模板表头。");
  }

  const parsedInvoices = [];
  for (const file of invoiceFiles) {
    parsedInvoices.push(await parsePdfInvoice(file));
  }

  const matchedRows = [];
  const errors = [];
  const ordinaryInvoices = [];

  parsedInvoices.forEach((invoice, index) => {
    if (isOrdinaryInvoice(invoice)) {
      ordinaryInvoices.push(invoice);
      return;
    }

    const matched = matchInvoice(invoice, exportIndex);
    if (!matched) {
      errors.push({
        fileIndex: index,
        fileName: invoice.fileName,
        reason: buildUnmatchedReason(invoice),
        invoice,
      });
      return;
    }

    if (!isNormalInvoiceStatus(matched)) {
      errors.push({
        fileIndex: index,
        fileName: invoice.fileName,
        reason: buildAbnormalStatusReason(matched),
        invoice,
      });
      return;
    }

    const matchedInvoiceNumber = normalizeNumber(pickRowValue(matched, ["数电发票号码", "发票号码", "发票代码"]));
    if (redFlushContext.redInvoiceNumbers.has(matchedInvoiceNumber)) {
      errors.push({
        fileIndex: index,
        fileName: invoice.fileName,
        reason: "红字发票不参与勾选抵扣",
        invoice,
      });
      return;
    }

    const effectiveDeductibleTax = calculateEffectiveDeductibleTax(matched, invoice, redFlushContext);
    if (redFlushContext.fullyRedFlushedInvoiceNumbers.has(matchedInvoiceNumber)) {
      errors.push({
        fileIndex: index,
        fileName: invoice.fileName,
        reason: "原发票已全额红冲，无可抵扣税额",
        invoice,
      });
      return;
    }

    matchedRows.push(buildTemplateRow({
      order: matchedRows.length + 1,
      templateHeaders,
      invoice,
      exportRow: matched,
      effectiveDeductibleTax,
    }));
  });

  const outputWorkbook = XLSX.utils.book_new();
  const worksheetRows = [templateHeaders, ...matchedRows.map((row) => templateHeaders.map((header) => row[header] ?? ""))];
  const outputSheet = XLSX.utils.aoa_to_sheet(worksheetRows);
  XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, "抵扣勾选增值税发票信息");
  const workbookBuffer = XLSX.write(outputWorkbook, { type: "buffer", bookType: "xlsx" });

  return {
    summary: {
      invoiceCount: parsedInvoices.length,
      ordinaryCount: ordinaryInvoices.length,
      matchedCount: matchedRows.length,
      errorCount: errors.length,
      totalAmount: roundMoney(sumBy(matchedRows, "金额*")),
      totalTax: roundMoney(sumBy(matchedRows, "有效抵扣税额*")),
      buyerTaxNos: uniqueValues(matchedRows, "购买方识别号*"),
    },
    templateHeaders,
    parsedInvoices,
    ordinaryInvoices,
    matchedRows,
    errors,
    workbookBuffer,
  };
}

async function buildPendingDownloadWorkbook({ invoiceFiles, fullExportFile, now = new Date() }) {
  const fullExportWorkbook = XLSX.read(fullExportFile.buffer, { type: "buffer", cellDates: false });
  const fullExportSourceRows = collectWorkbookRows(fullExportWorkbook);
  const fullExportRows = mergeWorkbookRows(fullExportWorkbook);
  const redFlushContext = buildRedFlushContext(fullExportSourceRows);
  const parsedInvoices = [];

  for (const file of invoiceFiles) {
    parsedInvoices.push(await parsePdfInvoice(file));
  }

  const uploadedKeys = new Set(
    parsedInvoices
      .filter((invoice) => !isOrdinaryInvoice(invoice))
      .flatMap((invoice) => buildInvoiceLookupKeys(invoice).filter(Boolean))
  );
  const pendingRows = [];
  const filteredRows = fullExportRows.filter((row) => {
    if (!isPendingImportEligibleRow(row)) return false;
    if (!isPreviousMonthInvoice(row["开票日期"], now)) return false;
    const invoiceNumber = normalizeNumber(pickRowValue(row, ["数电发票号码", "发票号码", "发票代码"]));
    if (redFlushContext.redInvoiceNumbers.has(invoiceNumber)) {
      return false;
    }
    if (redFlushContext.fullyRedFlushedInvoiceNumbers.has(invoiceNumber)) {
      return false;
    }

    const rowKeys = buildRowLookupKeys(row);
    if (!rowKeys.length) return true;
    return !rowKeys.some((key) => uploadedKeys.has(key));
  });

  pendingRows.push(...filteredRows);
  const pendingCount = pendingRows.length;

  if (!pendingRows.length) {
    return {
      pendingCount,
      workbookBuffer: buildEmptyImportTemplateWorkbook(),
    };
  }

  const pendingWorkbook = buildImportTemplateWorkbook(pendingRows);

  return {
    pendingCount,
    workbookBuffer: XLSX.write(pendingWorkbook, { type: "buffer", bookType: "xlsx" }),
  };
}

function loadTemplateWorkbook() {
  const templatePath = DEFAULT_TEMPLATE_PATH.path;
  if (!templatePath) {
    throw new Error(`后台模板不存在，已查找：${DEFAULT_TEMPLATE_PATH.candidates.join("、")}`);
  }
  return XLSX.read(fs.readFileSync(templatePath), { type: "buffer", cellDates: false });
}

function loadImportTemplateWorkbook() {
  const templatePath = IMPORT_TEMPLATE_PATH.path;
  if (!templatePath) {
    throw new Error(`导入清单模板不存在，已查找：${IMPORT_TEMPLATE_PATH.candidates.join("、")}`);
  }
  return XLSX.read(fs.readFileSync(templatePath), {
    type: "buffer",
    cellDates: false,
    cellStyles: true,
  });
}

function resolveTemplatePath(fileName) {
  const candidates = TEMPLATE_DIR_CANDIDATES.map((dirName) => path.join(__dirname, dirName, fileName));
  return {
    path: candidates.find((candidate) => fs.existsSync(candidate)) || "",
    candidates,
  };
}

function buildEmptyImportTemplateWorkbook() {
  const workbook = loadImportTemplateWorkbook();
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  clearImportTemplateDataRows(sheet);
  return XLSX.write(workbook, {
    type: "buffer",
    bookType: "xlsx",
    cellStyles: true,
  });
}

function buildImportTemplateWorkbook(rows) {
  const workbook = loadImportTemplateWorkbook();
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const templateHeaders = readImportTemplateHeaders(sheet);

  if (!templateHeaders.length) {
    throw new Error("未能读取“导入清单模板”表头。");
  }

  const sampleStyles = readImportTemplateSampleStyles(sheet);
  clearImportTemplateDataRows(sheet);

  rows.forEach((row, index) => {
    const dataRow = buildImportTemplateRow(row, index + 1);
    validateImportTemplateRow(dataRow, row);
    writeImportTemplateRow(sheet, index + 2, dataRow, sampleStyles);
  });

  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:G1");
  range.e.r = Math.max(range.e.r, rows.length + 1);
  range.e.c = Math.max(range.e.c, 6);
  sheet["!ref"] = XLSX.utils.encode_range(range);

  return workbook;
}

function mergeWorkbookRows(workbook) {
  return mergeRowsByInvoiceNumber(collectWorkbookRows(workbook));
}

function mergeRowsByInvoiceNumber(rows) {
  const merged = new Map();

  for (const row of rows) {
    const key = normalizeNumber(row["数电发票号码"] || row["发票号码"] || row["发票代码"]);
    if (!key) continue;
    const current = merged.get(key) || {};
    merged.set(key, mergeRowObjects(current, row));
  }

  return Array.from(merged.values());
}

function collectWorkbookRows(workbook) {
  return workbook.SheetNames.flatMap((sheetName) => XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    defval: "",
    raw: false,
  }));
}

function buildExportIndex(rows) {
  const byDigitalNumber = new Map();
  const bySellerAndAmount = new Map();

  for (const row of rows) {
    const digitalNumber = normalizeNumber(pickRowValue(row, ["数电发票号码", "发票号码", "发票代码"]));
    if (digitalNumber) {
      byDigitalNumber.set(digitalNumber, row);
    }

    const fallbackKey = buildFallbackKey({
      sellerName: pickRowValue(row, ["销方名称", "销售方纳税人名称"]),
      issueDate: pickRowValue(row, ["开票日期"]),
      grossAmount: pickRowValue(row, ["价税合计", "价税总额", "合计金额"]),
    });
    if (fallbackKey) {
      bySellerAndAmount.set(fallbackKey, row);
    }
  }

  return { byDigitalNumber, bySellerAndAmount, rows };
}

function readTemplateHeaders(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: "",
  });
  return (rows[0] || []).map((value) => String(value).trim()).filter(Boolean);
}

function readImportTemplateHeaders(sheet) {
  return ["A1", "B1", "C1", "D1", "E1", "F1", "G1"]
    .map((address) => String(sheet[address]?.v || "").trim())
    .filter(Boolean);
}

function readImportTemplateSampleStyles(sheet) {
  return ["A", "B", "C", "D", "E", "F", "G"].map((column) => {
    const style = sheet[`${column}2`]?.s;
    return style ? JSON.parse(JSON.stringify(style)) : undefined;
  });
}

function clearImportTemplateDataRows(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:G1");
  for (let row = 1; row <= range.e.r; row += 1) {
    for (let column = 0; column <= range.e.c; column += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: column });
      delete sheet[address];
    }
  }
  range.e.r = 0;
  sheet["!ref"] = XLSX.utils.encode_range(range);
}

function buildImportTemplateRow(exportRow, order) {
  return [
    String(order),
    pickRowValue(exportRow, ["数电发票号码", "发票号码", "发票代码"]),
    normalizeCompactDate(pickRowValue(exportRow, ["开票日期"])),
    pickRowValue(exportRow, ["销方名称", "销售方纳税人名称"]),
    pickRowValue(exportRow, ["销方识别号", "销售方纳税人识别号"]),
    numericOrString(pickRowValue(exportRow, ["金额", "发票金额（不含税）", "不含税金额", "金额(不含税)"]), ""),
    numericOrString(pickRowValue(exportRow, ["税额"]), ""),
  ];
}

function validateImportTemplateRow(dataRow, exportRow) {
  const requiredFields = [
    "序号",
    "数电发票号码（必填）",
    "开票日期（必填）",
    "销售方纳税人名称",
    "销售方纳税人识别号",
    "发票金额（不含税）",
    "税额",
  ];

  const missingFields = requiredFields.filter((field, index) => {
    const value = dataRow[index];
    return value === "" || value == null;
  });

  if (!missingFields.length) return;

  const invoiceRef = exportRow["数电发票号码"] || exportRow["发票号码"] || exportRow["发票代码"] || "未识别发票";
  throw new Error(`导入清单模板存在必填项缺失：${invoiceRef} 缺少 ${missingFields.join("、")}`);
}

function writeImportTemplateRow(sheet, rowNumber, dataRow, sampleStyles) {
  dataRow.forEach((value, index) => {
    const address = XLSX.utils.encode_cell({ r: rowNumber - 1, c: index });
    const cell = createImportTemplateCell(value);
    if (sampleStyles[index]) {
      cell.s = JSON.parse(JSON.stringify(sampleStyles[index]));
    }
    sheet[address] = cell;
  });
}

function createImportTemplateCell(value) {
  if (typeof value === "number") {
    return { t: "n", v: value, z: "@" };
  }
  return { t: "s", v: String(value), z: "@" };
}

async function parsePdfInvoice(file) {
  const parser = new PDFParse({ data: file.buffer });
  try {
    const result = await parser.getText();
    const text = normalizePdfText(result.text || "");
    const fileName = normalizeUploadFileName(file.originalname);

    return {
      fileName,
      invoiceType: pickInvoiceType(text),
      digitalInvoiceNumber: pickInvoiceNumber(text) || pickInvoiceNumber(fileName),
      issueDate: pickInvoiceIssueDate(text),
      buyerName: pickBuyerName(text),
      buyerTaxNo: pickFirst(text, [/购\s*买\s*方\s*信\s*息.*?统一社会信用代码\/纳税人识别号[:：]?\s*([0-9A-Z]{15,25})/]),
      sellerName: pickSellerName(text),
      sellerTaxNo: pickFirst(text, [/销\s*售\s*方\s*信\s*息.*?统一社会信用代码\/纳税人识别号[:：]?\s*([0-9A-Z]{15,25})/, /\b([0-9A-Z]{15,25})\b/g], true),
      amount: parseMoney(pickFirst(text, [/合\s*计\s*¥?\s*([0-9,]+\.[0-9]{2})/, /¥\s*([0-9,]+\.[0-9]{2})\s+¥\s*[0-9,]+\.[0-9]{2}\s+价税合计/])),
      tax: parseMoney(pickFirst(text, [/合\s*计\s*¥?\s*[0-9,]+\.[0-9]{2}\s+¥?\s*([0-9,]+\.[0-9]{2})/, /税\s*额\s.*?([0-9,]+\.[0-9]{2})\s+开票人/])),
      grossAmount: parseMoney(pickFirst(text, [/（小写）\s*¥?\s*([0-9,]+\.[0-9]{2})/, /¥\s*([0-9,]+\.[0-9]{2})\s+壹/, /¥\s*([0-9,]+\.[0-9]{2})\s+贰/, /¥\s*([0-9,]+\.[0-9]{2})\s+叁/, /票\s*价\s*[:：]?\s*¥?\s*([0-9,]+\.[0-9]{2})/])),
      rawText: text,
    };
  } finally {
    await parser.destroy();
  }
}

function matchInvoice(invoice, exportIndex) {
  const digitalNumber = normalizeNumber(invoice.digitalInvoiceNumber);
  if (digitalNumber && exportIndex.byDigitalNumber.has(digitalNumber)) {
    return exportIndex.byDigitalNumber.get(digitalNumber);
  }

  const fallbackKey = buildFallbackKey({
    sellerName: invoice.sellerName,
    issueDate: invoice.issueDate,
    grossAmount: invoice.grossAmount,
  });
  if (fallbackKey && exportIndex.bySellerAndAmount.has(fallbackKey)) {
    return exportIndex.bySellerAndAmount.get(fallbackKey);
  }

  const fileNameMatched = matchInvoiceByFileName(invoice, exportIndex.rows || []);
  if (fileNameMatched) {
    return fileNameMatched;
  }

  return null;
}

function matchInvoiceByFileName(invoice, rows) {
  const clues = extractInvoiceFileNameClues(invoice.fileName);
  if (!clues.issueDate || !clues.sellerKeyword) return null;

  const candidates = rows.filter((row) => {
    const rowDate = normalizeCompactDate(pickRowValue(row, ["开票日期"]));
    if (rowDate !== clues.issueDate) return false;

    const sellerName = cleanText(pickRowValue(row, ["销方名称", "销售方纳税人名称"]));
    if (!sellerName.includes(clues.sellerKeyword)) return false;

    return isSelectableUploadedInvoiceRow(row);
  });

  return candidates.length === 1 ? candidates[0] : null;
}

function extractInvoiceFileNameClues(fileName) {
  const baseName = cleanText(fileName).replace(/\.[^.]+$/, "");
  const dateMatch = baseName.match(/(^|[^\d])(\d{2})(\d{2})(\d{2})(?!\d)/);
  if (!dateMatch) {
    return { issueDate: "", sellerKeyword: "" };
  }

  const year = Number(`20${dateMatch[2]}`);
  const month = Number(dateMatch[3]);
  const day = Number(dateMatch[4]);
  if (month < 1 || month > 12 || day < 1 || day > 31) {
    return { issueDate: "", sellerKeyword: "" };
  }

  const dateEndIndex = (dateMatch.index || 0) + dateMatch[0].length;
  const sellerKeyword = baseName
    .slice(dateEndIndex)
    .replace(/发票.*$/, "")
    .replace(/[（(].*$/, "")
    .replace(/[^\p{Script=Han}A-Za-z0-9]/gu, "")
    .trim();

  return {
    issueDate: `${year}${String(month).padStart(2, "0")}${String(day).padStart(2, "0")}`,
    sellerKeyword: sellerKeyword.length >= 2 ? sellerKeyword : "",
  };
}

function buildTemplateRow({ order, templateHeaders, invoice, exportRow, effectiveDeductibleTax }) {
  const base = {};
  for (const header of templateHeaders) {
    base[header] = mapTemplateValue(header, invoice, exportRow, order, effectiveDeductibleTax);
  }
  return base;
}

function mapTemplateValue(header, invoice, exportRow, order, effectiveDeductibleTax) {
  const directMap = {
    "序号": order,
    "是否勾选*": TEMPLATE_DEFAULTS["是否勾选*"],
    "数电发票号码": pickRowValue(exportRow, ["数电发票号码", "发票号码", "发票代码"]) || invoice.digitalInvoiceNumber,
    "数电发票号码*": pickRowValue(exportRow, ["数电发票号码", "发票号码", "发票代码"]) || invoice.digitalInvoiceNumber,
    "发票代码": pickRowValue(exportRow, ["发票代码"]),
    "发票号码": pickRowValue(exportRow, ["发票号码"]),
    "开票日期*": pickRowValue(exportRow, ["开票日期"]) || invoice.issueDate || "",
    "金额*": numericOrString(pickRowValue(exportRow, ["金额", "发票金额（不含税）", "不含税金额", "金额(不含税)"]), invoice.amount),
    "票面税额*": numericOrString(pickRowValue(exportRow, ["税额"]), invoice.tax),
    "有效抵扣税额*": effectiveDeductibleTax ?? numericOrString(pickRowValue(exportRow, ["有效抵扣税额", "税额"]), invoice.tax),
    "购买方识别号*": pickRowValue(exportRow, ["购方识别号", "购买方识别号"]) || invoice.buyerTaxNo || "",
    "销售方纳税人名称": pickRowValue(exportRow, ["销方名称", "销售方纳税人名称"]) || invoice.sellerName || "",
    "销售方纳税人识别号": pickRowValue(exportRow, ["销方识别号", "销售方纳税人识别号"]) || invoice.sellerTaxNo || "",
    "发票来源": pickRowValue(exportRow, ["发票来源"]) || TEMPLATE_DEFAULTS["发票来源"],
    "票种*": pickRowValue(exportRow, ["发票票种", "票种", "发票种类"]) || invoice.invoiceType || "",
    "发票状态": pickRowValue(exportRow, ["发票状态", "状态"]),
    "红字锁定标志": TEMPLATE_DEFAULTS["红字锁定标志"],
    "转内销证明编号": pickRowValue(exportRow, ["转内销证明编号"]),
    "业务类型": pickRowValue(exportRow, ["特定业务类型", "业务类型"]),
    "勾选时间": pickRowValue(exportRow, ["勾选时间"]),
    "发票风险等级": pickRowValue(exportRow, ["发票风险等级", "风险等级"]),
    "风险状态": exportRow["风险状态"] || inferRiskStatus(exportRow),
    "差额扣除标识": pickRowValue(exportRow, ["差额扣除标识"]),
  };

  if (directMap[header] !== undefined) {
    return directMap[header];
  }

  if (exportRow[header] !== undefined) {
    return exportRow[header];
  }

  return "";
}

function inferRiskStatus(exportRow) {
  const riskLevel = pickRowValue(exportRow, ["发票风险等级", "风险等级"]);
  if (!riskLevel) return "";
  return riskLevel === "正常" ? "无风险" : "";
}

function normalizePdfText(text) {
  return text.replace(/\u0000/g, "").replace(/\s+/g, " ").trim();
}

function normalizeUploadFileName(fileName) {
  const text = cleanText(fileName);
  if (!text) return "";

  const decoded = Buffer.from(text, "latin1").toString("utf8");
  return countChineseCharacters(decoded) > countChineseCharacters(text) ? decoded : text;
}

function countChineseCharacters(value) {
  return (String(value || "").match(/\p{Script=Han}/gu) || []).length;
}

function pickBuyerName(text) {
  const buyerTaxNo = pickFirst(text, [/购\s*买\s*方\s*信\s*息.*?统一社会信用代码\/纳税人识别号[:：]?\s*([0-9A-Z]{15,25})/]);
  if (!buyerTaxNo) return "";
  const match = text.match(new RegExp(`开票日期[:：]?.{0,30}?${buyerTaxNo}\\s+(.+?)\\s+销\\s*售\\s*方\\s*信\\s*息`));
  return match ? cleanText(match[1]) : "";
}

function pickSellerName(text) {
  const match = text.match(/销\s*售\s*方\s*信\s*息.*?名称[:：]?\s*([^\s].*?)\s+统一社会信用代码\/纳税人识别号/s);
  if (match) return cleanText(match[1]);

  const candidate = text.match(/杭州安唐文化发展有限公司\s+[0-9A-Z]{18}\s+(.+?)\s+[0-9A-Z]{18}/);
  return candidate ? cleanText(candidate[1]) : "";
}

function pickInvoiceType(text) {
  return pickFirst(text, [
    /(电子\s*发票（[^）]+）)/,
    /(数电\s*发票（[^）]+）)/,
  ]).replace(/\s+/g, "");
}

function pickInvoiceIssueDate(text) {
  const chineseDate = text.match(/开\s*票\s*日\s*期\s*[:：]?\s*([0-9]{4})\s*年\s*([0-9]{1,2})\s*月\s*([0-9]{1,2})\s*日/);
  if (chineseDate) {
    return `${chineseDate[1]}-${String(chineseDate[2]).padStart(2, "0")}-${String(chineseDate[3]).padStart(2, "0")}`;
  }

  return normalizeDate(pickFirst(text, [/([0-9]{4}-[0-9]{2}-[0-9]{2})/]));
}

function pickInvoiceNumber(text) {
  return pickFirst(text, [
    /(?:数\s*电\s*发\s*票\s*号\s*码|发\s*票\s*号\s*码)[:：]?\s*([0-9]{20})(?![0-9A-Z])/,
    /(?:数\s*电\s*发\s*票\s*号\s*码|发\s*票\s*号\s*码)[:：]?[\s\S]{0,80}?(?<![0-9A-Z])([0-9]{20})(?![0-9A-Z])/,
    /(?<![0-9A-Z])([0-9]{20})(?![0-9A-Z])/,
  ]);
}

function pickFirst(text, patterns, useLast = false) {
  for (const pattern of patterns) {
    if (pattern.global) {
      const matches = Array.from(text.matchAll(pattern));
      const target = useLast ? matches[matches.length - 1] : matches[0];
      if (target?.[1]) return cleanText(target[1]);
      continue;
    }
    const match = text.match(pattern);
    if (match?.[1]) return cleanText(match[1]);
  }
  return "";
}

function cleanText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function normalizeDate(value) {
  const text = cleanText(value);
  if (!text) return "";
  const digits = text.replace(/[^\d]/g, "");
  if (digits.length >= 8) {
    return `${digits.slice(0, 4)}-${digits.slice(4, 6)}-${digits.slice(6, 8)}`;
  }
  return text;
}

function normalizeCompactDate(value) {
  const normalized = normalizeDate(value);
  return normalized ? normalized.replace(/[^\d]/g, "").slice(0, 8) : "";
}

function parseMoney(value) {
  if (value == null || value === "") return null;
  const numeric = Number(String(value).replace(/[^\d.-]/g, ""));
  return Number.isFinite(numeric) ? roundMoney(numeric) : null;
}

function normalizeNumber(value) {
  return String(value || "").replace(/[^\dA-Z]/gi, "");
}

function buildFallbackKey({ sellerName, issueDate, grossAmount }) {
  const normalizedName = cleanText(sellerName);
  const normalizedDate = normalizeDate(issueDate);
  const normalizedAmount = parseMoney(grossAmount);
  if (!normalizedName || !normalizedDate || normalizedAmount == null) return "";
  return `${normalizedName}|${normalizedDate}|${normalizedAmount.toFixed(2)}`;
}

function buildUnmatchedReason(invoice) {
  return "未在“全量发票查询导出结果”中找到匹配记录";
}

function buildAbnormalStatusReason(exportRow) {
  const status = cleanText(pickRowValue(exportRow, ["发票状态", "状态"]));
  return status ? `发票状态为“${status}”` : "发票状态不是“正常”";
}

function isOrdinaryInvoice(invoice) {
  return /(普通发票)/.test(invoice.invoiceType || "");
}

function isNormalInvoiceStatus(row) {
  return cleanText(pickRowValue(row, ["发票状态", "状态"])) === "正常";
}

function isSpecialInvoiceRow(row) {
  return /(增值税专用发票)/.test(String(pickRowValue(row, ["发票票种", "票种", "发票种类"]) || ""));
}

function isRailwayElectronicInvoiceRow(row) {
  return /(铁路电子客票|铁路电子发票)/.test(String(pickRowValue(row, ["发票票种", "票种", "发票种类"]) || ""));
}

function isSelectableUploadedInvoiceRow(row) {
  return isSpecialInvoiceRow(row) || isRailwayElectronicInvoiceRow(row);
}

function isPendingImportEligibleRow(row) {
  return isSpecialInvoiceRow(row) && !isRailwayElectronicInvoiceRow(row);
}

function isPreviousMonthInvoice(value, now = new Date()) {
  const date = parseDateValue(value);
  if (!date) return false;

  const targetYear = now.getMonth() === 0 ? now.getFullYear() - 1 : now.getFullYear();
  const targetMonth = now.getMonth() === 0 ? 11 : now.getMonth() - 1;

  return date.getFullYear() === targetYear && date.getMonth() === targetMonth;
}

function parseDateValue(value) {
  const text = cleanText(value);
  if (!text) return null;

  const normalized = text.replace("T", " ").replace(/\//g, "-");
  const match = normalized.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return null;

  const year = Number(match[1]);
  const month = Number(match[2]) - 1;
  const day = Number(match[3]);
  const date = new Date(year, month, day);
  return Number.isNaN(date.getTime()) ? null : date;
}

function buildInvoiceLookupKeys(invoice) {
  const keys = [];
  const digitalNumber = normalizeNumber(invoice.digitalInvoiceNumber);
  if (digitalNumber) {
    keys.push(`digital:${digitalNumber}`);
  }

  const fallbackKey = buildFallbackKey({
    sellerName: invoice.sellerName,
    issueDate: invoice.issueDate,
    grossAmount: invoice.grossAmount,
  });
  if (fallbackKey) {
    keys.push(`fallback:${fallbackKey}`);
  }

  return keys;
}

function buildRowLookupKeys(row) {
  const keys = [];
  const digitalNumber = normalizeNumber(pickRowValue(row, ["数电发票号码", "发票号码", "发票代码"]));
  if (digitalNumber) {
    keys.push(`digital:${digitalNumber}`);
  }

  const fallbackKey = buildFallbackKey({
    sellerName: pickRowValue(row, ["销方名称", "销售方纳税人名称"]),
    issueDate: pickRowValue(row, ["开票日期"]),
    grossAmount: pickRowValue(row, ["价税合计", "价税总额", "合计金额"]),
  });
  if (fallbackKey) {
    keys.push(`fallback:${fallbackKey}`);
  }

  return keys;
}

function buildRedFlushContext(rows) {
  const uniqueRows = mergeRowsByInvoiceNumber(rows);
  const redFlushTotals = buildRedFlushTotalsByOriginalInvoice(uniqueRows);
  const redInvoiceNumbers = buildRedInvoiceNumberSet(uniqueRows);
  const fullyRedFlushedInvoiceNumbers = buildFullyRedFlushedInvoiceNumberSet(uniqueRows, redFlushTotals);

  return {
    redFlushTotals,
    redInvoiceNumbers,
    fullyRedFlushedInvoiceNumbers,
  };
}

function buildFullyRedFlushedInvoiceNumberSet(rows, redFlushTotals) {
  const result = new Set();

  for (const row of rows) {
    if (isRedInvoiceRow(row)) continue;

    const invoiceNumber = normalizeNumber(pickRowValue(row, ["数电发票号码", "发票号码", "发票代码"]));
    if (!invoiceNumber) continue;

    const redFlushTotal = redFlushTotals.get(invoiceNumber);
    if (!redFlushTotal) continue;

    const amount = parseMoney(pickRowValue(row, ["金额", "发票金额（不含税）", "不含税金额", "金额(不含税)"]));
    const tax = parseMoney(pickRowValue(row, ["有效抵扣税额", "有效抵扣税额*", "税额"]));
    const grossAmount = parseMoney(pickRowValue(row, ["价税合计", "价税总额", "合计金额"]));

    const amountCovered = amount != null && redFlushTotal.amount >= Math.abs(amount) - 0.005;
    const taxCovered = tax != null && redFlushTotal.tax >= Math.abs(tax) - 0.005;
    const grossCovered = grossAmount != null && redFlushTotal.grossAmount >= Math.abs(grossAmount) - 0.005;

    if (tax != null ? taxCovered : (amount != null ? amountCovered : grossCovered)) {
      result.add(invoiceNumber);
    }
  }

  return result;
}

function buildRedFlushTotalsByOriginalInvoice(rows) {
  const totals = new Map();

  for (const row of rows) {
    if (!isRedInvoiceRow(row)) continue;

    const originalInvoiceNumber = extractOriginalInvoiceNumber(row);
    if (!originalInvoiceNumber) continue;

    const current = totals.get(originalInvoiceNumber) || {
      amount: 0,
      tax: 0,
      grossAmount: 0,
    };

    current.amount = roundMoney(current.amount + Math.abs(parseMoney(pickRowValue(row, ["金额", "发票金额（不含税）", "不含税金额", "金额(不含税)"])) || 0));
    current.tax = roundMoney(current.tax + Math.abs(parseMoney(pickRowValue(row, ["有效抵扣税额", "有效抵扣税额*", "税额"])) || 0));
    current.grossAmount = roundMoney(current.grossAmount + Math.abs(parseMoney(pickRowValue(row, ["价税合计", "价税总额", "合计金额"])) || 0));
    totals.set(originalInvoiceNumber, current);
  }

  return totals;
}

function buildRedInvoiceNumberSet(rows) {
  const result = new Set();

  for (const row of rows) {
    if (!isRedInvoiceRow(row)) continue;

    const currentInvoiceNumber = normalizeNumber(pickRowValue(row, ["数电发票号码", "发票号码", "发票代码"]));
    if (currentInvoiceNumber) {
      result.add(currentInvoiceNumber);
    }
  }

  return result;
}

function isRedInvoiceRow(row) {
  const positiveFlag = cleanText(row["是否正数发票"]);
  if (positiveFlag) return positiveFlag === "否";

  const amount = parseMoney(pickRowValue(row, ["金额", "发票金额（不含税）", "不含税金额", "金额(不含税)"]));
  const tax = parseMoney(pickRowValue(row, ["税额"]));
  return (amount != null && amount < 0) || (tax != null && tax < 0);
}

function calculateEffectiveDeductibleTax(exportRow, invoice, redFlushContext) {
  const originalTax = parseMoney(pickRowValue(exportRow, ["有效抵扣税额", "有效抵扣税额*", "税额"]))
    ?? parseMoney(invoice.tax);
  if (originalTax == null) return null;

  const invoiceNumber = normalizeNumber(pickRowValue(exportRow, ["数电发票号码", "发票号码", "发票代码"]));
  const redFlushTax = redFlushContext.redFlushTotals.get(invoiceNumber)?.tax || 0;
  return Math.max(0, roundMoney(originalTax - redFlushTax));
}

function extractOriginalInvoiceNumber(row) {
  const remark = cleanText(pickRowValue(row, ["备注"]));
  if (!remark) return "";

  const match = remark.match(/被红冲蓝字(?:数电)?发票号码[:：]\s*([0-9A-Z]+)/i);
  return normalizeNumber(match?.[1] || "");
}

function mergeRowObjects(current, incoming) {
  const merged = { ...current };
  for (const [key, value] of Object.entries(incoming)) {
    if (value !== "" && value != null) {
      merged[key] = value;
    }
  }
  return merged;
}

function pickRowValue(row, keys) {
  for (const key of keys) {
    const value = row[key];
    if (value !== "" && value != null) {
      return value;
    }
  }
  return "";
}

function numericOrString(primary, fallback) {
  if (primary !== "" && primary != null) return normalizeCellValue(primary);
  if (fallback !== "" && fallback != null) return normalizeCellValue(fallback);
  return "";
}

function normalizeCellValue(value) {
  const numeric = Number(value);
  return Number.isFinite(numeric) && String(value).trim() !== "" ? roundMoney(numeric) : value;
}

function roundMoney(value) {
  return Math.round(Number(value) * 100) / 100;
}

function sumBy(rows, key) {
  return rows.reduce((sum, row) => sum + (Number(row[key]) || 0), 0);
}

function uniqueValues(rows, key) {
  return [...new Set(rows.map((row) => row[key]).filter(Boolean))];
}

module.exports = {
  buildPendingDownloadWorkbook,
  processInvoiceBundle,
  __test: {
    buildRedFlushContext,
    calculateEffectiveDeductibleTax,
  },
};
