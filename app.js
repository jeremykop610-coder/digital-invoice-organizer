const processForm = document.getElementById("processForm");
const invoiceInput = document.getElementById("invoiceInput");
const fullExportInput = document.getElementById("fullExportInput");
const pendingDownloadButton = document.getElementById("pendingDownloadButton");
const processButton = document.getElementById("processButton");
const resetButton = document.getElementById("resetButton");
const processStatus = document.getElementById("processStatus");
const pendingMonthLabel = document.getElementById("pendingMonthLabel");
const statusPill = document.getElementById("statusPill");
const invoiceCount = document.getElementById("invoiceCount");
const matchedCount = document.getElementById("matchedCount");
const ordinaryCount = document.getElementById("ordinaryCount");
const errorCount = document.getElementById("errorCount");
const taxTotal = document.getElementById("taxTotal");
const errorList = document.getElementById("errorList");
const ordinaryList = document.getElementById("ordinaryList");
const downloadButton = document.getElementById("downloadButton");

const state = {
  workbookBytes: null,
  outputFileName: "",
  pendingWorkbookBytes: null,
  pendingOutputFileName: "",
  errorFileUrls: [],
};

const apiOrigin = resolveApiOrigin();

pendingMonthLabel.textContent = formatPreviousMonth(new Date());

processForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (!invoiceInput.files.length || !fullExportInput.files.length) {
    window.alert("请先上传 PDF 发票和全量发票查询导出结果。");
    return;
  }

  const formData = new FormData();
  Array.from(invoiceInput.files).forEach((file) => formData.append("invoices", file));
  formData.append("fullExport", fullExportInput.files[0]);

  setProcessing(true, "正在生成最终勾选表...");

  try {
    const payload = await postFormData("/api/process", formData);
    applyResult(payload);
    setProcessing(
      false,
      payload.errors.length
        ? "第二步完成：勾选表已生成，但仍有未匹配发票。"
        : "第二步完成：勾选表已生成，可下载结果。",
    );
  } catch (error) {
    clearResult();
    setStatus("issue", "处理失败");
    processStatus.textContent = error.message;
    setProcessing(false, error.message);
  }
});

resetButton.addEventListener("click", () => {
  processForm.reset();
  processStatus.textContent = "等待上传";
  clearResult();
});

pendingDownloadButton.addEventListener("click", async () => {
  if (!invoiceInput.files.length || !fullExportInput.files.length) {
    window.alert("请先上传 PDF 发票和全量发票查询导出结果。");
    return;
  }

  const formData = new FormData();
  Array.from(invoiceInput.files).forEach((file) => formData.append("invoices", file));
  formData.append("fullExport", fullExportInput.files[0]);

  setBusyState(true, "正在执行第一步，检查上月未上传专票...");

  try {
    const payload = await postFormData("/api/pending-download", formData);

    state.pendingWorkbookBytes = base64ToUint8Array(payload.workbookBase64);
    state.pendingOutputFileName = payload.outputFileName;
    downloadWorkbook(state.pendingWorkbookBytes, state.pendingOutputFileName || "导入清单模板.xlsx");
    processStatus.textContent = payload.pendingCount
      ? `第一步完成：已生成导入清单，共 ${payload.pendingCount} 张上月未上传专票。补齐发票后再生成勾选表。`
      : "第一步完成：没有需补齐的上月专票，可直接生成勾选表。";
  } catch (error) {
    processStatus.textContent = error.message;
  } finally {
    setBusyState(false, processStatus.textContent);
  }
});

downloadButton.addEventListener("click", () => {
  if (!state.workbookBytes) return;
  downloadWorkbook(state.workbookBytes, state.outputFileName || "抵扣勾选增值税发票信息.xlsx");
});

async function postFormData(path, formData) {
  const response = await fetch(buildApiUrl(path), {
      method: "POST",
      body: formData,
    });
  return readJsonResponse(response);
}

function applyResult(payload) {
  state.workbookBytes = base64ToUint8Array(payload.workbookBase64);
  state.outputFileName = payload.outputFileName;

  invoiceCount.textContent = String(payload.summary.invoiceCount);
  matchedCount.textContent = String(payload.summary.matchedCount);
  ordinaryCount.textContent = String(payload.summary.ordinaryCount);
  errorCount.textContent = String(payload.summary.errorCount);
  taxTotal.textContent = formatCurrency(payload.summary.totalTax);
  setStatus(payload.errors.length ? "issue" : "ready", payload.errors.length ? "存在异常" : "已生成");
  renderErrors(payload.errors);
  renderOrdinary(payload.ordinaryInvoices || []);
  downloadButton.disabled = payload.matchedRows.length === 0;
}

function clearResult() {
  releaseErrorFileUrls();
  state.workbookBytes = null;
  state.outputFileName = "";
  state.pendingWorkbookBytes = null;
  state.pendingOutputFileName = "";
  invoiceCount.textContent = "0";
  matchedCount.textContent = "0";
  ordinaryCount.textContent = "0";
  errorCount.textContent = "0";
  taxTotal.textContent = "¥0.00";
  errorList.innerHTML = '<li class="empty-state">当前没有异常。</li>';
  ordinaryList.innerHTML = '<li class="empty-state">当前没有普通发票。</li>';
  setStatus("", "未执行");
  downloadButton.disabled = true;
}

function renderErrors(errors) {
  releaseErrorFileUrls();
  errorList.innerHTML = "";
  if (!errors.length) {
    errorList.innerHTML = '<li class="empty-state">当前没有异常。</li>';
    return;
  }

  errors.forEach((item) => {
    const li = document.createElement("li");
    const fileLink = buildErrorFileLink(item);
    const parts = [];
    if (fileLink) {
      li.appendChild(fileLink);
    } else if (item.fileName) {
      parts.push(`PDF ${item.fileName}`);
    }
    parts.push(item.digitalInvoiceNumber ? `发票号 ${item.digitalInvoiceNumber}` : "发票号未识别");
    if (item.sellerName) parts.push(`销售方 ${item.sellerName}`);
    if (item.issueDate) parts.push(`开票日期 ${item.issueDate}`);
    if (item.grossAmount != null) parts.push(`价税合计 ${formatCurrency(item.grossAmount)}`);
    if (item.reason) parts.push(`原因 ${item.reason}`);
    if (parts.length) {
      if (li.childNodes.length) {
        li.append(`：${parts.join("，")}`);
      } else {
        li.textContent = parts.join("，");
      }
    }
    errorList.appendChild(li);
  });
}

function renderOrdinary(rows) {
  ordinaryList.innerHTML = "";
  if (!rows.length) {
    ordinaryList.innerHTML = '<li class="empty-state">当前没有普通发票。</li>';
    return;
  }

  rows.forEach((item) => {
    const li = document.createElement("li");
    const invoiceNo = item.digitalInvoiceNumber ? `发票号 ${item.digitalInvoiceNumber}` : "发票号未识别";
    const sellerName = item.sellerName ? `，销售方 ${item.sellerName}` : "";
    li.textContent = `${invoiceNo}${sellerName}`;
    ordinaryList.appendChild(li);
  });
}

function setProcessing(isProcessing, message) {
  setBusyState(isProcessing, message);
}

function setStatus(kind, text) {
  statusPill.textContent = text;
  statusPill.className = `status-pill${kind ? ` ${kind}` : ""}`;
}

function setBusyState(isBusy, message) {
  pendingDownloadButton.disabled = isBusy;
  processButton.disabled = isBusy;
  downloadButton.disabled = isBusy || !state.workbookBytes;
  processStatus.textContent = message;
}

function formatPreviousMonth(now) {
  const previousMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  return `${previousMonth.getFullYear()}年${previousMonth.getMonth() + 1}月`;
}

function resolveApiOrigin() {
  const { protocol, hostname, port, origin } = window.location;
  if (protocol === "file:") {
    return "http://localhost:3000";
  }
  if ((hostname === "localhost" || hostname === "127.0.0.1") && port && port !== "3000") {
    return "http://localhost:3000";
  }
  return origin;
}

function buildApiUrl(path) {
  return new URL(path, `${apiOrigin}/`).toString();
}

async function readJsonResponse(response) {
  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("application/json")) {
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "请求失败");
    }
    return payload;
  }

  const bodyText = await response.text();
  if (bodyText.startsWith("<!DOCTYPE") || bodyText.startsWith("<html")) {
    throw new Error(`接口返回了 HTML 页面。请使用 http://localhost:3000 打开网站，并确认 Node 服务正在运行。当前接口地址：${buildApiUrl(response.url ? new URL(response.url).pathname : "/api")}`);
  }

  throw new Error(bodyText || "接口返回了非 JSON 内容。");
}

function formatCurrency(value) {
  return new Intl.NumberFormat("zh-CN", {
    style: "currency",
    currency: "CNY",
  }).format(Number(value || 0));
}

function base64ToUint8Array(base64) {
  const binary = window.atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function buildErrorFileLink(item) {
  const file = Number.isInteger(item.fileIndex) ? invoiceInput.files[item.fileIndex] : null;
  if (!file) return null;

  const url = URL.createObjectURL(file);
  state.errorFileUrls.push(url);

  const link = document.createElement("a");
  link.href = url;
  link.target = "_blank";
  link.rel = "noopener noreferrer";
  link.textContent = item.fileName || file.name || "打开原件";
  return link;
}

function releaseErrorFileUrls() {
  state.errorFileUrls.forEach((url) => URL.revokeObjectURL(url));
  state.errorFileUrls = [];
}

function downloadWorkbook(bytes, fileName) {
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

clearResult();
