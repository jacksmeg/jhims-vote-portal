const path = require("node:path");
const { ensureVoterTemplate } = require("../src/helpers/excel");

async function main() {
  const templatePaths = [
    path.join(
      process.cwd(),
      "public",
      "templates",
      "voter-import-template.xlsx",
    ),
    path.join(
      process.cwd(),
      "public",
      "templates",
      "staff-login-template.xlsx",
    ),
  ];

  for (const templatePath of templatePaths) {
    await ensureVoterTemplate(templatePath);
    console.log(`Created template: ${templatePath}`);
  }
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
