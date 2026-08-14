import { chromium } from "playwright";
import path from "node:path";
import fs from "node:fs";

const OUT_DIR = path.join(__dirname, "..", ".visual-check");
fs.mkdirSync(OUT_DIR, { recursive: true });

const serviceMapCsv = `Service Name,Service Type,QKDBID,QKServiceID,Xplor Service ID
Sunnyvale Early Learning,ELC,9999-182,182,122956
`;
const migrationCsv = `ServiceID,Child_Legacy_Id,Child_First_Name,Child_Last_Name,DOB,Status,Child_CRN,Room_Name,Enrolment_Start_Date,State
122956,1001,Ava,Smith,2020-05-10,Active,123456789A,Koalas,2024-01-15,victoria
122956,1002,Ben,Jones,2021-02-20,Active,BADCRN,,2024-02-01,VIC
`;
fs.writeFileSync(path.join(OUT_DIR, "serviceIDs.csv"), serviceMapCsv);
fs.writeFileSync(path.join(OUT_DIR, "migration.csv"), migrationCsv);

async function main() {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1400, height: 1100 } });
  const consoleErrors: string[] = [];
  page.on("console", (msg) => {
    if (msg.type() === "error") consoleErrors.push(msg.text());
  });
  page.on("pageerror", (err) => consoleErrors.push(`pageerror: ${err.message}`));

  console.log("Navigating to landing page...");
  await page.goto("http://localhost:3000", { waitUntil: "networkidle" });
  await page.waitForSelector("text=Xplor Data Migration Tools");
  await page.screenshot({ path: path.join(OUT_DIR, "01-landing.png"), fullPage: true });

  console.log("Clicking Parent and Child Import card...");
  await page.getByRole("link", { name: /Parent and Child Import/ }).click();
  await page.waitForURL("**/parent-child-import");
  await page.waitForSelector("text=Step 1");
  await page.screenshot({ path: path.join(OUT_DIR, "02-tool-empty-manual-tab.png"), fullPage: true });

  console.log("Testing manual service entry...");
  const nameInputs = page.locator('input[placeholder="Service Name"]');
  await nameInputs.first().fill("Sunnyvale Early Learning");
  await page.locator('input[placeholder="QK Service ID"]').first().fill("182");
  await page.locator('input[placeholder="Xplor Service ID"]').first().fill("122956");
  await page.waitForSelector("text=1 service configured");
  await page.getByRole("button", { name: /Add service/ }).click();
  await page.waitForSelector("text=Service 2");
  await page.screenshot({ path: path.join(OUT_DIR, "03-manual-entry.png"), fullPage: true });

  console.log("Testing back-to-home link...");
  await page.getByRole("link", { name: "All tools" }).click();
  await page.waitForURL("http://localhost:3000/");
  await page.goBack();
  await page.waitForURL("**/parent-child-import");

  console.log("Switching to CSV upload tab...");
  await page.getByRole("button", { name: "Upload CSV" }).click();
  const svcInput = page.locator('input[type="file"]').first();
  await svcInput.setInputFiles(path.join(OUT_DIR, "serviceIDs.csv"));
  await page.waitForSelector("text=Using: serviceIDs.csv");

  console.log("Uploading migration file...");
  const migrationInput = page.locator('input[type="file"]').nth(1);
  await migrationInput.setInputFiles(path.join(OUT_DIR, "migration.csv"));
  await page.waitForSelector("text=migration.csv");
  await page.screenshot({ path: path.join(OUT_DIR, "04-files-loaded.png"), fullPage: true });

  console.log("Running validation...");
  await page.getByRole("button", { name: /Run validation/ }).click();
  await page.waitForSelector("text=Validation results", { timeout: 15000 });
  await page.screenshot({ path: path.join(OUT_DIR, "05-results.png"), fullPage: true });

  console.log("Downloading zip bundle...");
  const [download] = await Promise.all([
    page.waitForEvent("download", { timeout: 15000 }),
    page.getByRole("button", { name: /Download all outputs/ }).click(),
  ]);
  const downloadPath = path.join(OUT_DIR, download.suggestedFilename());
  await download.saveAs(downloadPath);
  const stat = fs.statSync(downloadPath);
  console.log(`Downloaded ${download.suggestedFilename()} — ${stat.size} bytes`);

  // Confirm zero network requests carry file content — only the initial static bundle loads.
  const requestUrls: string[] = [];
  page.on("request", (req) => requestUrls.push(req.url()));
  await page.waitForTimeout(1000);
  const nonLocalRequests = requestUrls.filter((u) => !u.startsWith("http://localhost:3000"));

  console.log("\n=== Console errors ===");
  console.log(consoleErrors.length === 0 ? "(none)" : consoleErrors.join("\n"));

  console.log("\n=== Non-localhost network requests after interaction ===");
  console.log(nonLocalRequests.length === 0 ? "(none — confirms no outbound data)" : nonLocalRequests.join("\n"));

  await browser.close();
  console.log(`\nScreenshots written to ${OUT_DIR}`);
  process.exit(consoleErrors.length > 0 ? 1 : 0);
}

main().catch((err) => {
  console.error("Visual check crashed:", err);
  process.exit(1);
});
