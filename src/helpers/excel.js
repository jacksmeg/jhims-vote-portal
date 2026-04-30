const fs = require("node:fs");
const path = require("node:path");
const ExcelJS = require("exceljs");

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "_");
}

function getCellText(cell) {
  if (!cell) {
    return "";
  }

  if (typeof cell.text === "string" && cell.text.trim()) {
    return cell.text.trim();
  }

  if (cell.value === null || cell.value === undefined) {
    return "";
  }

  if (typeof cell.value === "object") {
    if ("text" in cell.value && cell.value.text) {
      return String(cell.value.text).trim();
    }

    if ("result" in cell.value && cell.value.result) {
      return String(cell.value.result).trim();
    }
  }

  return String(cell.value).trim();
}

function readWorksheetRows(worksheet) {
  const rows = [];

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    rows.push({ row, rowNumber });
  });

  return rows;
}

async function parseVoterWorkbook(filePath) {
  const workbook = new ExcelJS.Workbook();
  const extension = path.extname(filePath).toLowerCase();

  if (extension === ".csv") {
    await workbook.csv.readFile(filePath);
  } else {
    await workbook.xlsx.readFile(filePath);
  }

  const worksheet = workbook.worksheets[0];

  if (!worksheet) {
    throw new Error("The uploaded workbook does not contain any sheets.");
  }

  const worksheetRows = readWorksheetRows(worksheet);

  if (worksheetRows.length === 0) {
    throw new Error("The uploaded workbook is empty.");
  }

  const headerRow = worksheetRows[0].row;
  const headers = [];

  headerRow.eachCell({ includeEmpty: true }, (cell, columnNumber) => {
    headers[columnNumber - 1] = normalizeHeader(getCellText(cell));
  });

  const requiredHeaders = ["staff_id", "phone_number"];
  const missingHeaders = requiredHeaders.filter((header) => !headers.includes(header));

  if (missingHeaders.length > 0) {
    throw new Error(
      `The upload is missing required columns: ${missingHeaders.join(", ")}.`,
    );
  }

  const parsedRows = [];

  for (const rowInfo of worksheetRows.slice(1)) {
    const rowValues = [];

    rowInfo.row.eachCell({ includeEmpty: true }, (cell, columnNumber) => {
      rowValues[columnNumber - 1] = getCellText(cell);
    });

    const record = {};
    let hasAnyValue = false;

    headers.forEach((header, index) => {
      const value = rowValues[index] ? String(rowValues[index]).trim() : "";
      record[header] = value;
      if (value) {
        hasAnyValue = true;
      }
    });

    if (!hasAnyValue) {
      continue;
    }

    record.__rowNumber = rowInfo.rowNumber;
    parsedRows.push(record);
  }

  return parsedRows;
}

async function ensureVoterTemplate(templatePath) {
  if (fs.existsSync(templatePath)) {
    return;
  }

  fs.mkdirSync(path.dirname(templatePath), { recursive: true });

  const workbook = new ExcelJS.Workbook();
  const instructions = workbook.addWorksheet("Instructions");
  const worksheet = workbook.addWorksheet("Voters");

  instructions.columns = [
    { header: "Step", key: "step", width: 18 },
    { header: "Guidance", key: "guidance", width: 68 },
  ];
  instructions.addRows([
    {
      step: "1",
      guidance: "Enter each voter's unique staff_id in the exact format used by your organization.",
    },
    {
      step: "2",
      guidance: "Enter the phone_number that should be used during voter login.",
    },
    {
      step: "3",
      guidance: "Optional columns full_name and department can be filled to make admin records easier to review.",
    },
    {
      step: "4",
      guidance: "Do not repeat the same staff_id in more than one row.",
    },
  ]);
  instructions.getRow(1).font = { bold: true };
  instructions.views = [{ state: "frozen", ySplit: 1 }];

  worksheet.columns = [
    { header: "staff_id", key: "staff_id", width: 18 },
    { header: "phone_number", key: "phone_number", width: 18 },
    { header: "full_name", key: "full_name", width: 28 },
    { header: "department", key: "department", width: 22 },
  ];

  worksheet.addRow({
    staff_id: "STF0234",
    phone_number: "0240000000",
    full_name: "Sample Staff Member",
    department: "Operations",
  });

  worksheet.getRow(1).font = { bold: true };
  worksheet.views = [{ state: "frozen", ySplit: 1 }];

  await workbook.xlsx.writeFile(templatePath);
}

module.exports = {
  ensureVoterTemplate,
  parseVoterWorkbook,
};
