const fs = require("node:fs");
const path = require("node:path");
const PDFDocument = require("pdfkit");

const outputPath = path.join(
  process.cwd(),
  "public",
  "downloads",
  "organization-vote-portal-user-guide.pdf",
);

function ensureSpace(doc, height = 80) {
  const bottomLimit = doc.page.height - doc.page.margins.bottom;
  if (doc.y + height > bottomLimit) {
    doc.addPage();
  }
}

function sectionTitle(doc, text) {
  ensureSpace(doc, 48);
  doc
    .moveDown(0.8)
    .font("Helvetica-Bold")
    .fontSize(16)
    .fillColor("#102338")
    .text(text);
  doc.moveDown(0.2);
}

function paragraph(doc, text) {
  ensureSpace(doc, 48);
  doc
    .font("Helvetica")
    .fontSize(11)
    .fillColor("#102338")
    .text(text, {
      lineGap: 4,
    });
  doc.moveDown(0.45);
}

function bulletList(doc, items) {
  for (const item of items) {
    ensureSpace(doc, 34);
    doc
      .font("Helvetica")
      .fontSize(11)
      .fillColor("#102338")
      .text(`• ${item}`, {
        indent: 12,
        lineGap: 4,
      });
  }
  doc.moveDown(0.35);
}

function numberedList(doc, items) {
  items.forEach((item, index) => {
    ensureSpace(doc, 36);
    doc
      .font("Helvetica")
      .fontSize(11)
      .fillColor("#102338")
      .text(`${index + 1}. ${item}`, {
        indent: 12,
        lineGap: 4,
      });
  });
  doc.moveDown(0.35);
}

function addPageNumbers(doc) {
  const range = doc.bufferedPageRange();
  for (let index = 0; index < range.count; index += 1) {
    doc.switchToPage(index);
    doc
      .font("Helvetica")
      .fontSize(9)
      .fillColor("#5d6d80")
      .text(
        `Page ${index + 1} of ${range.count}`,
        doc.page.margins.left,
        doc.page.height - doc.page.margins.bottom + 16,
        {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          align: "right",
        },
      );
  }
}

function buildGuide() {
  const currentYear = new Date().getFullYear();
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });

  const doc = new PDFDocument({
    size: "A4",
    margin: 50,
    bufferPages: true,
    info: {
      Title: "Organization Vote Portal User Guide",
      Author: "JACKSTUDIOS",
      Subject: "User manual for the election software",
    },
  });

  const stream = fs.createWriteStream(outputPath);
  doc.pipe(stream);

  doc
    .font("Helvetica-Bold")
    .fontSize(24)
    .fillColor("#102338")
    .text("Organization Vote Portal", { align: "left" });

  doc
    .moveDown(0.25)
    .font("Helvetica")
    .fontSize(13)
    .fillColor("#5d6d80")
    .text("Software User Guide", { align: "left" })
    .text("For administrators and voters", { align: "left" });

  doc
    .moveDown(0.8)
    .roundedRect(50, doc.y, 495, 98, 18)
    .fillAndStroke("#f8f3e7", "#d8c197");

  doc
    .fillColor("#102338")
    .font("Helvetica-Bold")
    .fontSize(12)
    .text("Live Access Links", 68, doc.y - 82)
    .font("Helvetica")
    .fontSize(11)
    .text("Voter login: https://vote.jhimssoftware.com/vote/login", 68, doc.y + 4)
    .text("Admin login: https://vote.jhimssoftware.com/admin/login", 68, doc.y + 6)
    .text("Guide download: https://vote.jhimssoftware.com/downloads/organization-vote-portal-user-guide.pdf", 68, doc.y + 6);

  doc.y += 36;

  sectionTitle(doc, "1. What This Software Does");
  paragraph(
    doc,
    "The Organization Vote Portal helps an organization run secure internal elections. Voters use a staff ID and phone number to log in, vote once, and submit a ballot position by position. Administrators manage voter registration, candidates, positions, election timing, live monitoring, final reporting, archives, and system reset for the next election.",
  );

  sectionTitle(doc, "2. Voter Process");
  numberedList(doc, [
    "Open the voter link: https://vote.jhimssoftware.com/vote/login",
    "Enter the registered staff ID and phone number.",
    "The system checks that the staff ID exists, the phone number matches, and the voter has not already voted.",
    "If OTP verification is enabled, the system sends a one-time SMS code to the registered phone number and the voter must enter that code before the ballot opens.",
    "The voter sees one position at a time and chooses one candidate before moving to the next position.",
    "At the confirmation page, the voter reviews all selections and submits the ballot.",
    "After submission, the vote is saved, the voter is marked as voted, and the ballot cannot be changed.",
  ]);

  sectionTitle(doc, "3. Administrator Process");
  paragraph(
    doc,
    "Administrators sign in through the admin login page and get the full control dashboard. The public voter pages do not show the admin navigation. Administrators can set up the election, manage voters and candidates, monitor progress, close the election, export results, archive completed elections, and prepare the next cycle.",
  );

  sectionTitle(doc, "4. Dashboard Setup");
  bulletList(doc, [
    "Set the election name.",
    "Set the voting opening date and time.",
    "Set the voting closing date and time.",
    "Upload the organization logo.",
    "Choose one of the software themes: Heritage Gold, Emerald Pulse, or Midnight Blue.",
    "Create a database backup before major changes.",
  ]);

  sectionTitle(doc, "5. Managing Voters");
  bulletList(doc, [
    "Download the staff login Excel template from the Voters page.",
    "Import voters from Excel or CSV using staff_id and phone_number columns.",
    "Add a voter manually inside the admin page.",
    "Edit a voter record during setup before voting opens.",
    "Clear the imported voter list during setup if you want to start the voter list again.",
    "Each staff ID must be unique.",
  ]);

  sectionTitle(doc, "6. Positions and Candidates");
  bulletList(doc, [
    "Add election positions such as President, Secretary, Treasurer, or any custom title.",
    "Add candidates under each position.",
    "Upload a candidate photo.",
    "Edit candidate details during setup before voting opens.",
    "Set candidate order on the ballot.",
  ]);

  sectionTitle(doc, "7. Opening and Running an Election");
  numberedList(doc, [
    "Confirm that voters, positions, and candidates are ready.",
    "Open voting from the dashboard.",
    "Once opened, structural setup changes are locked to protect ballot integrity.",
    "Administrators can monitor turnout and live provisional statistics from the Results page while the election is running.",
    "The Results page auto-refreshes every 30 seconds during live voting.",
  ]);

  sectionTitle(doc, "8. Results, PDF Export, and Printing");
  bulletList(doc, [
    "The admin results page shows turnout and candidate-by-candidate totals.",
    "While voting is open, the totals are provisional and for admin monitoring.",
    "After the election is closed, results become final.",
    "Final results can be exported as PDF.",
    "Final results can also be opened in a print-friendly page for printing.",
  ]);

  sectionTitle(doc, "9. Archiving and Resetting for Another Election");
  paragraph(
    doc,
    "After the election is closed, the admin can archive the completed election and reset the system. This action saves the finished election into the archive history and clears the working voter list, positions, candidates, ballots, and results so the next election can be prepared from a clean setup state.",
  );
  bulletList(doc, [
    "Open archived elections from the Archives page.",
    "Review archived turnout and final result details.",
    "Delete an archive if it is no longer needed.",
  ]);

  sectionTitle(doc, "10. Security and Controls");
  bulletList(doc, [
    "One staff member can vote only once.",
    "Voter login requires both staff ID and phone number match.",
    "Optional OTP SMS verification can be enabled so a voter must confirm ownership of the registered phone before the ballot opens.",
    "Admin pages require administrator login.",
    "Audit logs track major admin and voter actions.",
    "Results can be monitored by admins during live voting and finalized after closure.",
    "Backups can be created from the admin dashboard.",
  ]);

  sectionTitle(doc, "11. Updating the Software");
  paragraph(
    doc,
    "Software updates are made in the project code, pushed to GitHub, and then automatically redeployed by Render. Content changes like voters, candidates, logo, themes, election name, and election dates are done inside the admin dashboard and do not need code changes.",
  );

  sectionTitle(doc, "12. Support Contact");
  bulletList(doc, [
    "Powered by JACKSTUDIOS",
    "Phone: 0592934612",
    "Email: jacksmeg99@gmail.com",
    `Copyright © ${currentYear} IT DEPARTMENT DUNKWA MUNICIPAL HOSPITAL. All rights reserved.`,
  ]);

  addPageNumbers(doc);
  doc.end();

  return new Promise((resolve, reject) => {
    stream.on("finish", resolve);
    stream.on("error", reject);
  });
}

buildGuide()
  .then(() => {
    console.log(`Created user guide PDF: ${outputPath}`);
  })
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });
