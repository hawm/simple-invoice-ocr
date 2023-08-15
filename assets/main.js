(async () => {
  globalThis.invoiceInput = document.getElementById("invoice-input");
  globalThis.invoiceTable = document.getElementById("invoice-table");
  globalThis.parseButton = document.getElementById("parse-btn");
  globalThis.parseButton.disabled = true;
  globalThis.exportButton = document.getElementById("export-btn");
  globalThis.exportButton.disabled = true;
  globalThis.clearButton = document.getElementById("clear-btn");
  globalThis.progressText = document.getElementById("progress-text");
  globalThis.ocrWorker = await Tesseract.createWorker({
    langPath: "./vendor/tessdata",
    //logger: (m) => console.log(m),
  });
  await ocrWorker.loadLanguage("chi_sim");
  await ocrWorker.initialize("chi_sim");
})();

async function parseInvoiceQRcodeData(canvas) {
  let ctx = canvas.getContext("2d");
  let imgData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  let code = jsQR(imgData.data, canvas.width, canvas.height);
  return code.data.split(",");
}

async function parseInvoiceTextData(canvas) {
  const AMOUNT_REGEX = /[\(（]\s?小\s?写\s?[\)）]\s?\D\s?(\d+(\.\d{2})?)/;
  let res = await globalThis.ocrWorker.recognize(canvas);
  let text = res.data.text;
  let amount = AMOUNT_REGEX.exec(text)[1];

  return {
    amount,
  };
}

async function parseImageInvoiceData(file) {
  let img = await loadImage(file, { canvas: true });
  let canvas = img.image;
  let qrCodeData = await parseInvoiceQRcodeData(canvas);
  let textData = await parseInvoiceTextData(canvas);
  return {
    qrCodeData,
    textData,
  };
}

async function parseInvoiceData(file) {
  const { qrCodeData, textData } = await parseImageInvoiceData(file);
  return {
    code: qrCodeData[2],
    num: qrCodeData[3],
    date: qrCodeData[5],
    amount: textData.amount,
  };
}

async function renderTableRowData(rowData) {
  const tr = document.createElement("tr");
  const header = globalThis.invoiceTable.rows[0];
  Array.from(header.children).forEach((th) => {
    let prop = th.getAttribute("itemprop");
    field = rowData[prop];
    let td = document.createElement("td");
    td.textContent = field;
    tr.appendChild(td);
  });

  globalThis.invoiceTable.appendChild(tr);
}

function clearTable() {
  globalThis.invoiceTable.innerHTML = "";
}

async function renderTableHeader(headerData) {
  // headerData should be an two dimensional array
  // [[name, itemprop], ...]
  clearTable();
  let thead = document.createElement("thead");
  let tr = document.createElement("tr");
  headerData.forEach((field) => {
    let th = document.createElement("th");
    th.setAttribute("itemprop", field[1]);
    th.textContent = field[0];
    tr.appendChild(th);
  });
  thead.appendChild(tr);
  globalThis.invoiceTable.appendChild(thead);
}

function updateProgressText(text) {
  globalThis.progressText.textContent = text;
}

async function parse() {
  const invoiceFiles = globalThis.invoiceInput.files;

  renderTableHeader([
    ["发票代码", "code"],
    ["发票号码", "num"],
    ["开盘日期", "date"],
    ["价税合计", "amount"],
    ["文件名", "file"],
  ]);

  for (let i = 0; i < invoiceFiles.length; i++) {
    let f = invoiceFiles[i];
    updateProgressText(
      `正在识别第${i + 1}张发票...， 还剩下${invoiceFiles.length - (i + 1)}张`
    );
    console.log(`Parsing ${f.name}`);
    let data = await parseInvoiceData(f);
    data["order"] = i + 1;
    data["file"] = f.name;
    renderTableRowData(data);
  }
  updateProgressText(`识别完成！共识别${invoiceFiles.length}张发票`);
  console.log("All done!");
}

function exportInvoiceXLSX() {
  let worksheet = XLSX.utils.table_to_book(
    document.getElementById("invoice-table")
  );
  time = new Date()
    .toISOString()
    .split(".")[0]
    .replaceAll("-", "")
    .replaceAll(":", "")
    .replace("T", "_");
  XLSX.writeFile(worksheet, `invoice_${time}.xlsx`);
}

function updateParseButtonState() {
  if (globalThis.invoiceInput.files.length == 0) {
    globalThis.parseButton.disabled = true;
  } else {
    globalThis.parseButton.disabled = false;
  }
}

globalThis.invoiceInput.addEventListener("change", (e) => {
  updateParseButtonState();
  updateProgressText(`已选择${e.target.files.length}张发票`);
});

globalThis.parseButton.addEventListener("click", async (e) => {
  try {
    e.target.disabled = true;
    e.target.textContent = "识别中...";
    globalThis.invoiceInput.disabled = true;
    globalThis.clearButton.disabled = true;
    globalThis.exportButton.disabled = true;
    await parse();
    globalThis.exportButton.disabled = false;
  } finally {
    e.target.textContent = "识别";
    updateParseButtonState();
    globalThis.clearButton.disabled = false;
    globalThis.invoiceInput.disabled = false;
  }
});

globalThis.exportButton.addEventListener("click", (e) => {
  try {
    e.target.disabled = true;
    globalThis.invoiceInput.disabled = true;
    globalThis.parseButton.disabled = true;
    globalThis.clearButton.disabled = true;
    updateProgressText("正在导出...");
    exportInvoiceXLSX();
  } finally {
    updateProgressText("导出完成！");
    e.target.disabled = false;
    globalThis.parseButton.disabled = false;
    globalThis.clearButton.disabled = false;
    globalThis.invoiceInput.disabled = false;
  }
});

globalThis.clearButton.addEventListener("click", (e) => {
  clearTable();
  globalThis.invoiceInput.value = null;
  globalThis.invoiceInput.dispatchEvent(new Event("change"));
});
