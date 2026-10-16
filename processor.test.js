const test = require("node:test");
const assert = require("node:assert/strict");

global.DOMMatrix = class DOMMatrix {};
global.ImageData = class ImageData {};
global.Path2D = class Path2D {};

const { __test } = require("./processor");

const originalInvoice = {
  "数电发票号码": "26332000000000000001",
  "是否正数发票": "是",
  "金额": "1000.00",
  "税额": "130.00",
  "价税合计": "1130.00",
};

function redInvoice(number, amount, tax) {
  return {
    "数电发票号码": number,
    "是否正数发票": "否",
    "金额": String(-amount),
    "税额": String(-tax),
    "价税合计": String(-(amount + tax)),
    "备注": `被红冲蓝字数电发票号码：${originalInvoice["数电发票号码"]}`,
  };
}

test("partial red flush deducts cumulative red tax once", () => {
  const partialRed = redInvoice("26332000000000000002", 400, 52);
  const context = __test.buildRedFlushContext([
    originalInvoice,
    partialRed,
    { ...partialRed, "发票票种": "数电发票（增值税专用发票）" },
  ]);

  assert.equal(context.redFlushTotals.get(originalInvoice["数电发票号码"]).tax, 52);
  assert.equal(context.fullyRedFlushedInvoiceNumbers.has(originalInvoice["数电发票号码"]), false);
  assert.equal(context.redInvoiceNumbers.has(partialRed["数电发票号码"]), true);
  assert.equal(__test.calculateEffectiveDeductibleTax(originalInvoice, { tax: null }, context), 78);
});

test("cumulative full red flush leaves no deductible tax", () => {
  const context = __test.buildRedFlushContext([
    originalInvoice,
    redInvoice("26332000000000000002", 400, 52),
    redInvoice("26332000000000000003", 600, 78),
  ]);

  assert.equal(context.fullyRedFlushedInvoiceNumbers.has(originalInvoice["数电发票号码"]), true);
  assert.equal(__test.calculateEffectiveDeductibleTax(originalInvoice, { tax: null }, context), 0);
});

test("railway XBRL attachment provides invoice fields when PDF text is unreadable", () => {
  const invoice = __test.parseRailwayInvoiceXml(`<?xml version="1.0" encoding="UTF-8"?>
    <xbrl xmlns:rai="http://example.com/railway">
      <rai:TypeOfVoucher>电子发票（铁路电子客票）</rai:TypeOfVoucher>
      <rai:ElectronicInvoiceRailwayETicketNumber>26319132270002140337</rai:ElectronicInvoiceRailwayETicketNumber>
      <rai:DateOfIssue>2026-09-28</rai:DateOfIssue>
      <rai:Fare>86.00</rai:Fare>
      <rai:TotalAmountExcludingTax>78.90</rai:TotalAmountExcludingTax>
      <rai:TaxAmount>7.10</rai:TaxAmount>
      <rai:NameOfPurchaser>上海安唐文化发展有限公司</rai:NameOfPurchaser>
      <rai:UnifiedSocialCreditCodeOfPurchaser>91310112MAKB6BDK24</rai:UnifiedSocialCreditCodeOfPurchaser>
      <rai:NumberOfOriginalInvoice>26319132270002089721</rai:NumberOfOriginalInvoice>
    </xbrl>`);

  assert.deepEqual(invoice, {
    invoiceType: "电子发票（铁路电子客票）",
    digitalInvoiceNumber: "26319132270002140337",
    issueDate: "2026-09-28",
    buyerName: "上海安唐文化发展有限公司",
    buyerTaxNo: "91310112MAKB6BDK24",
    amount: 78.9,
    tax: 7.1,
    grossAmount: 86,
  });
});
