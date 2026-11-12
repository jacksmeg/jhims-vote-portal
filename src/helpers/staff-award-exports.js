const ExcelJS = require("exceljs");
const PDFDocument = require("pdfkit");

function nominationFields(staffFields, criteria) {
  return [
    ["Nomination ID", row => row.id],
    ["Submitted UTC", row => row.submitted_at],
    ["Award category", (_, values) => values.category],
    ["Nomination type", (_, values) => values.nominationType === "self" ? "Self nomination" : "Peer nomination"],
    ...["nominee", "nominator"].flatMap(prefix => staffFields.map(([key, label]) => [
      (prefix === "nominee" ? "Nominee " : "Nominator ") + label.toLowerCase(),
      (_, values) => prefix === "nominator" && values.nominationType === "self" ? "" : values[prefix + "_" + key],
    ])),
    ["Years worked", (_, values) => values.years],
    ["Reasons for nomination", (_, values) => values.reasons],
    ...criteria.map(([key, label]) => [label + " (1-3)", (_, values) => values[key]]),
    ["Innovation example", (_, values) => values.idea],
    ["Implementation", (_, values) => values.implementation],
    ["Review status", (row, _, review, labels) => labels[row.review_status]],
    ["Biometric attendance findings", (_, values, review) => review.attendance],
    ["2025 appraisal findings", (_, values, review) => review.appraisal],
    ["Committee decision notes", (_, values, review) => review.notes],
    ["Reviewed by", (_, values, review) => review.reviewer],
    ["Reviewed UTC", (_, values, review) => review.reviewed_at],
  ];
}

async function exportNominations(res, format, data) {
  const { hospital, q, category, rows, staffFields, criteria, reviewStatuses } = data;
  const generated = new Date().toISOString();
  const filters = "Award category: " + (category || "All") + " | Search: " + (q || "All");
  const fields = nominationFields(staffFields, criteria);
  const records = rows.map(row => {
    const values = JSON.parse(row.submission);
    const review = row.committee_review ? JSON.parse(row.committee_review) : {};
    return fields.map(([, read]) => read(row, values, review, reviewStatuses) ?? "");
  });
  const filename = "staff-award-nominations-" + generated.slice(0, 10);

  if (format === "excel") {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = hospital;
    workbook.created = new Date(generated);
    const info = workbook.addWorksheet("Export information");
    info.addRows([
      ["Institution", hospital], ["Report", "2026 Staff Award Nominations"],
      ["Generated UTC", generated], ["Award category", category || "All"],
      ["Search", q || "All"], ["Nomination count", rows.length],
      ["Rating scale", "1 to 3 per competency"],
    ]);
    info.getColumn(1).width = 24;
    info.getColumn(2).width = 70;
    info.getColumn(1).font = { bold: true };
    const sheet = workbook.addWorksheet("Nominations", { views: [{ state: "frozen", ySplit: 1 }] });
    sheet.columns = fields.map(([header]) => ({ header, width: /findings|notes|Reasons|example|Implementation/.test(header) ? 55 : 25 }));
    sheet.addRows(records);
    sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: Math.max(1, sheet.rowCount), column: fields.length } };
    sheet.getRow(1).height = 44;
    sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFA7193D" } };
    sheet.eachRow(row => { row.alignment = { vertical: "top", wrapText: true }; });
    res.set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.set("Content-Disposition", 'attachment; filename="' + filename + '.xlsx"');
    await workbook.xlsx.write(res);
    res.end();
    return;
  }

  const document = new PDFDocument({
    size: "A4", margin: 45, bufferPages: true,
    info: { Title: "2026 Staff Award Nominations", Author: hospital },
  });
  res.set("Content-Type", "application/pdf");
  res.set("Content-Disposition", 'attachment; filename="' + filename + '.pdf"');
  document.pipe(res);
  document.on("error", error => res.destroy(error));
  document.font("Helvetica-Bold").fontSize(17).fillColor("#151a23").text(hospital);
  document.moveDown(0.5).fontSize(14).text("2026 Staff Award Nominations");
  document.moveDown(0.4).font("Helvetica").fontSize(9).fillColor("#505966")
    .text("Generated UTC: " + generated).text(filters).text("Nominations: " + rows.length);
  if (!records.length) document.moveDown(2).fontSize(11).text("No nominations match these filters.");
  records.forEach((record, index) => {
    document.addPage();
    document.font("Helvetica-Bold").fontSize(15).fillColor("#a7193d")
      .text("Nomination #" + rows[index].id);
    document.moveDown(0.6);
    fields.forEach(([label], column) => {
      const value = record[column];
      if (value === "") return;
      if (document.y > document.page.height - 110) document.addPage();
      document.font("Helvetica-Bold").fontSize(9).fillColor("#151a23").text(label);
      document.font("Helvetica").fontSize(10).text(String(value), { lineGap: 2 });
      document.moveDown(0.5);
    });
  });
  const pages = document.bufferedPageRange();
  for (let page = pages.start; page < pages.start + pages.count; page++) {
    document.switchToPage(page);
    document.font("Helvetica").fontSize(8).fillColor("#505966")
      .text("2026 Staff Awards | Page " + (page + 1) + " of " + pages.count, 45, document.page.height - 30, { lineBreak: false });
  }
  document.end();
}

module.exports = { exportNominations };
