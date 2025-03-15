let data = [];
let standardizedData = [];
let fileName = "";
let outputFileName = "";

const KNOWN_CURRENCIES = [
  "USD",
  "EUR",
  "GBP",
  "SGD",
  "AUD",
  "CAD",
  "JPY",
  "CNY",
  "HKD",
  "AED",
  "QAR",
  "SAR",
  "OMR",
  "BHD",
  "KWD",
  "MYR",
  "THB",
  "IDR",
  "PHP",
  "CHF",
  "SEK",
  "DKK",
  "NOK",
  "ZAR",
  "NZD",
  "RUB",
  "KRW",
  "BRL",
  "EGP",
  "PLN",
  "MXN",
  "TRY",
  "ARS",
  "CLP",
  "COP",
  "CZK",
  "HUF",
  "RON",
  "INR",
];

function parseExcelDate(serial) {
  if (typeof serial === "number") {
    const dateObj = XLSX.SSF.parse_date_code(serial);
    if (dateObj) {
      const dd = String(dateObj.d).padStart(2, "0");
      const mm = String(dateObj.m).padStart(2, "0");
      const yyyy = String(dateObj.y);
      return `${dd}/${mm}/${yyyy}`;
    }
  }
  return "";
}

function convertDateString(dateStr) {
  if (!dateStr) return "";
  const parts = dateStr.split(/[-/]/);
  if (parts.length === 3) {
    let [dd, mm, yyyy] = parts;
    if (dd.length === 1) dd = "0" + dd;
    if (mm.length === 1) mm = "0" + mm;
    if (yyyy.length === 2) yyyy = "20" + yyyy;
    return `${dd}/${mm}/${yyyy}`;
  }
  return dateStr;
}

function isEmpty(val) {
  return !val || val.toString().trim() === "";
}

function detectFormat(inputData) {
  for (let row of inputData) {
    if (row && row.length > 0) {
      const rowStr = row.join(" ").toLowerCase();

      if (row[0]) {
        const colALower = row[0].toString().toLowerCase();
        if (colALower.includes("transaction details")) {
          return "fourth";
        }
      }

      if (row[3]) {
        const colDLower = row[3].toString().toLowerCase();
        if (
          colDLower.includes("transaction detail") ||
          colDLower.includes("transaction details")
        ) {
          return "third";
        }
      }

      if (rowStr.includes("amount") && rowStr.includes("transaction")) {
        return "first";
      }

      if (rowStr.includes("debit") && rowStr.includes("credit")) {
        return "second";
      }
    }
  }
  return "first";
}

function isNameRow(row) {
  if (!row[1]) return false;
  const candidate = row[1].toString().trim().toLowerCase();
  if (
    candidate.includes("domestic") ||
    candidate.includes("international") ||
    candidate.includes("transaction")
  ) {
    return false;
  }
  for (let c = 0; c < row.length; c++) {
    if (c === 1) continue;
    const val = row[c];
    if (val && val.toString().trim() !== "") {
      return false;
    }
  }
  return true;
}

function standardizeFirstFormat(inputData) {
  const outputData = [
    [
      "Date",
      "Transaction Description",
      "Debit",
      "Credit",
      "Currency",
      "CardName",
      "Transaction",
      "Location",
    ],
  ];
  let currentTransactionType = "";
  let currentCardName = "";
  let inDataSection = false;

  for (let i = 0; i < inputData.length; i++) {
    const row = inputData[i];
    if (!row || row.length === 0) continue;

    if (row[1] && typeof row[1] === "string") {
      const lower = row[1].toLowerCase();
      if (lower.includes("domestic")) {
        currentTransactionType = "Domestic";
        continue;
      }
      if (lower.includes("international")) {
        currentTransactionType = "International";
        continue;
      }
    }

    if (
      row[0] &&
      row[0].toString().toLowerCase().includes("date") &&
      row[1] &&
      row[1].toString().toLowerCase().includes("transaction") &&
      row[2] &&
      row[2].toString().toLowerCase().includes("amount")
    ) {
      inDataSection = true;
      continue;
    }

    if (isNameRow(row)) {
      currentCardName = row[1].trim();
      continue;
    }

    if (inDataSection) {
      let finalDate = "";
      if (typeof row[0] === "number") {
        finalDate = parseExcelDate(row[0]);
      } else {
        finalDate = convertDateString(row[0]?.toString().trim());
      }

      const rawDesc = row[1] ? row[1].toString().trim() : "";
      const rawAmount = row[2] ? row[2].toString().trim() : "";
      const isCredit = rawAmount.toLowerCase().includes("cr");
      const numericAmount = parseFloat(rawAmount.replace(/[^0-9.]+/g, "")) || 0;

      const debitVal = isCredit ? 0 : numericAmount;
      const creditVal = isCredit ? numericAmount : 0;

      let currency = "INR";
      let location = "";
      const words = rawDesc.split(" ").filter(Boolean);

      if (currentTransactionType.toLowerCase() === "domestic") {
        if (words.length > 0) {
          location = words.pop();
        }
      } else if (currentTransactionType.toLowerCase() === "international") {
        if (words.length > 1) {
          currency = words.pop();
          location = words.pop();
        }
      }
      const finalDesc = words.join(" ");

      outputData.push([
        finalDate,
        finalDesc,
        debitVal,
        creditVal,
        currency,
        currentCardName,
        currentTransactionType,
        location,
      ]);
    }
  }
  return outputData;
}

function standardizeSecondFormat(inputData) {
  const outputData = [
    [
      "Date",
      "Transaction Description",
      "Debit",
      "Credit",
      "Currency",
      "CardName",
      "Transaction",
      "Location",
    ],
  ];
  let currentTransactionType = "Domestic";
  let currentCardName = "";

  for (let i = 0; i < inputData.length; i++) {
    const row = inputData[i];
    if (!row || row.length === 0) continue;

    const colA = row[0] || "";
    const colB = row[1] || "";
    const colC = row[2] || "";
    const colD = row[3] || "";
    const colE = row[4] || "";

    const rowStr = `${colA} ${colB} ${colC} ${colD} ${colE}`.toLowerCase();

    if (
      rowStr.includes("date") &&
      rowStr.includes("transaction") &&
      rowStr.includes("debit") &&
      rowStr.includes("credit")
    ) {
      continue;
    }

    const colCLower = colC.toString().toLowerCase();
    if (colCLower.includes("domestic")) {
      currentTransactionType = "Domestic";
      continue;
    }
    if (colCLower.includes("international")) {
      currentTransactionType = "International";
      continue;
    }

    if (
      !isEmpty(colC) &&
      isEmpty(colA) &&
      isEmpty(colB) &&
      isEmpty(colD) &&
      isEmpty(colE)
    ) {
      currentCardName = colC.toString().trim();
      continue;
    }

    if (isEmpty(colA) && isEmpty(colB)) continue;

    let finalDate = "";
    if (typeof colA === "number") {
      finalDate = parseExcelDate(colA);
    } else {
      finalDate = convertDateString(colA.toString().trim());
    }

    const rawDesc = colB.toString().trim();
    let debitVal = parseFloat(String(colC).replace(/[^0-9.]+/g, "")) || 0;
    let creditVal = parseFloat(String(colD).replace(/[^0-9.]+/g, "")) || 0;

    const words = rawDesc.split(" ").filter(Boolean);
    let currency = "INR";
    let location = "";

    if (currentTransactionType === "Domestic") {
      if (words.length > 0) {
        location = words.pop();
      }
    } else {
      if (words.length > 1) {
        currency = words.pop().toUpperCase();
        location = words.pop();
      }
    }
    const finalDesc = words.join(" ");

    outputData.push([
      finalDate,
      finalDesc,
      debitVal,
      creditVal,
      currency,
      currentCardName,
      currentTransactionType,
      location,
    ]);
  }
  return outputData;
}

function standardizeThirdFormat(inputData) {
  const outputData = [
    [
      "Date",
      "Transaction Description",
      "Debit",
      "Credit",
      "Currency",
      "CardName",
      "Transaction",
      "Location",
    ],
  ];
  let currentTransactionType = "Domestic";
  let currentCardName = "";

  for (let i = 0; i < inputData.length; i++) {
    const row = inputData[i];
    if (!row || row.length === 0) continue;

    const colA = row[0] ? row[0].toString().trim() : "";
    const colB = row[1] ? row[1].toString().trim() : "";
    const colC = row[2] ? row[2].toString().trim() : "";
    const colD = row[3] ? row[3].toString().trim() : "";

    const rowStr = (colA + " " + colB + " " + colC + " " + colD).toLowerCase();
    if (
      rowStr.includes("date") &&
      rowStr.includes("debit") &&
      rowStr.includes("credit") &&
      rowStr.includes("transaction")
    ) {
      continue;
    }

    if (colC.toLowerCase().includes("domestic")) {
      currentTransactionType = "Domestic";
      continue;
    }
    if (colC.toLowerCase().includes("international")) {
      currentTransactionType = "International";
      continue;
    }

    if (!isEmpty(colC) && isEmpty(colA) && isEmpty(colB) && isEmpty(colD)) {
      currentCardName = colC.toString().trim();
      continue;
    }

    let finalDate = "";
    if (!isEmpty(colA)) {
      if (!isNaN(colA)) {
        finalDate = parseExcelDate(Number(colA));
      } else {
        finalDate = convertDateString(colA);
      }
    }

    let debitVal = parseFloat(colB.replace(/[^0-9.]+/g, "")) || 0;
    let creditVal = parseFloat(colC.replace(/[^0-9.]+/g, "")) || 0;

    const words = colD.split(" ").filter(Boolean);
    let currency = "INR";
    let location = "";

    if (words.length > 0) {
      const lastToken = words[words.length - 1].toUpperCase();
      if (KNOWN_CURRENCIES.includes(lastToken) && lastToken !== "INR") {
        currentTransactionType = "International";
        currency = lastToken;
        words.pop();
        if (words.length > 0) {
          location = words.pop();
        }
      } else if (currentTransactionType === "Domestic") {
        location = words.pop() || "";
      } else {
        if (words.length > 0) {
          location = words.pop();
        }
      }
    }

    const finalDesc = words.join(" ");

    outputData.push([
      finalDate,
      finalDesc,
      debitVal,
      creditVal,
      currency,
      currentCardName,
      currentTransactionType,
      location,
    ]);
  }
  return outputData;
}

function standardizeFourthFormat(inputData) {
  const outputData = [
    [
      "Date",
      "Transaction Description",
      "Debit",
      "Credit",
      "Currency",
      "CardName",
      "Transaction",
      "Location",
    ],
  ];

  let currentTransactionType = "Domestic";
  let currentCardName = "";

  for (let i = 0; i < inputData.length; i++) {
    const row = inputData[i];

    if (!row || row.length < 1) {
      continue;
    }

    const colA = row[0] ? row[0].toString().trim() : "";
    const colB = row[1] ? row[1].toString().trim() : "";
    const colC = row[2] ? row[2].toString().trim() : "";
    const colD = row[3] ? row[3].toString().trim() : "";

    const colBLower = colB.toLowerCase();

    if (colBLower.includes("domestic transactions")) {
      currentTransactionType = "Domestic";
      continue;
    }

    if (colBLower.includes("international transactions")) {
      currentTransactionType = "International";
      continue;
    }

    if (!isEmpty(colB)) {
      if (colBLower.includes("rahul")) {
        currentCardName = "Rahul";
        continue;
      }
      if (colBLower.includes("rajat")) {
        currentCardName = "Rajat";
        continue;
      }
    }

    if (
      colBLower.includes("date") &&
      colBLower.includes("transaction") &&
      colBLower.includes("amount")
    ) {
      continue;
    }
    if (colBLower === "transaction details") {
      continue;
    }

    let finalDate = "";
    if (!isEmpty(colB)) {
      if (!isNaN(colB)) {
        finalDate = parseExcelDate(Number(colB));
      } else {
        finalDate = convertDateString(colB);
      }
    }

    let debitVal = 0;
    let creditVal = 0;
    if (!isEmpty(colC)) {
      const rawAmount = colC.toLowerCase();
      const isCredit = rawAmount.includes("cr");
      const numericAmount = parseFloat(rawAmount.replace(/[^0-9.]+/g, "")) || 0;
      if (isCredit) {
        creditVal = numericAmount;
      } else {
        debitVal = numericAmount;
      }
    }

    let words = colA.split(" ").filter(Boolean);
    let currency = "INR";
    let location = "";
    if (words.length > 0) {
      const lastToken = words[words.length - 1].toUpperCase();
      if (KNOWN_CURRENCIES.includes(lastToken) && lastToken !== "INR") {
        currentTransactionType = "International";
        currency = lastToken;
        words.pop();

        if (words.length > 0) {
          location = words.pop();
        }
      } else if (currentTransactionType === "Domestic") {
        location = words.pop() || "";
      } else {
        if (words.length > 0) {
          location = words.pop();
        }
      }
    }
    const finalDesc = words.join(" ");

    outputData.push([
      finalDate,
      finalDesc,
      debitVal,
      creditVal,
      currency,
      currentCardName,
      currentTransactionType,
      location,
    ]);
  }

  return outputData;
}

function handleFileChange(e) {
  const file = e.target.files[0];
  if (!file) return;

  fileName = file.name;
  document.getElementById(
    "selectedFile"
  ).textContent = `Selected File: ${fileName}`;

  const reader = new FileReader();
  reader.onload = (evt) => {
    const binaryStr = evt.target.result;

    const workbook = XLSX.read(binaryStr, { type: "binary" });
    const wsname = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[wsname];

    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    data = jsonData;

    const format = detectFormat(jsonData);

    if (format === "first") {
      standardizedData = standardizeFirstFormat(jsonData);
    } else if (format === "second") {
      standardizedData = standardizeSecondFormat(jsonData);
    } else if (format === "third") {
      standardizedData = standardizeThirdFormat(jsonData);
    } else if (format === "fourth") {
      standardizedData = standardizeFourthFormat(jsonData);
    } else {
      standardizedData = standardizeFirstFormat(jsonData);
    }

    outputFileName = file.name.replace("Input", "Output");

    const downloadBtn = document.getElementById("downloadBtn");
    downloadBtn.style.display = "inline-block";
    downloadBtn.textContent = `Download ${outputFileName}`;
  };
  reader.readAsBinaryString(file);
}

function downloadCSV() {
  if (!standardizedData || standardizedData.length === 0) return;

  const ws = XLSX.utils.aoa_to_sheet(standardizedData);

  const csvContent = XLSX.utils.sheet_to_csv(ws);

  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = outputFileName || "output.csv";
  link.click();
}

window.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("fileInput");
  fileInput.addEventListener("change", handleFileChange);

  const downloadBtn = document.getElementById("downloadBtn");
  downloadBtn.addEventListener("click", downloadCSV);
});
