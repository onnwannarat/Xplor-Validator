/**
 * One-off smoke test: runs the client-side validation pipeline against synthetic
 * fixture data (no real client data) and prints a summary for manual spot-checking
 * against the documented rules. Not part of the app bundle — dev-only verification.
 */
import { ServiceMapping } from "../src/lib/validator/serviceMapping";
import { runValidation } from "../src/lib/validator/pipeline";
import { buildSplitCsvs } from "../src/lib/validator/reports/splitCsv";
import { buildExcelReport } from "../src/lib/validator/reports/excelReport";
import { buildDuplicateParentsReport } from "../src/lib/validator/reports/duplicateParentsReport";

const serviceMapCsv = `Service Name,Service Type,QKDBID,QKServiceID,Xplor Service ID
Sunnyvale Early Learning,ELC,9999-182,182,122956
`;

const HEADER = [
  "ServiceID",
  "Service_Name",
  "Child_Legacy_Id",
  "Child_First_Name",
  "Child_Last_Name",
  "DOB",
  "Status",
  "Child_CRN",
  "Room_Name",
  "Enrolment_Start_Date",
  "State",
  "PostCode",
  "Parent1_CRN",
  "Parent1_Legacy_Account_ID",
  "Parent1_DOB",
  "Parent1_Email",
  "Parent1_First_Name",
  "Parent1_Last_Name",
  "Parent1_Contact_Mobile",
].join(",");

const rows = [
  // 1. Clean row (direct Xplor ID)
  [
    "122956", "Sunnyvale Early Learning", "1001", "Ava", "Smith", "2020-05-10", "Active",
    "123456789A", "Koalas", "2024-01-15", "VIC", "3000",
    "987654321B", "P-1001", "1990-01-01", "ava.parent@example.com", "Beth", "Smith", "0412345678",
  ],
  // 2. Missing mandatory field (Room_Name blank)
  [
    "122956", "Sunnyvale Early Learning", "1002", "Ben", "Jones", "2021-02-20", "Active",
    "223456789A", "", "2024-02-01", "VIC", "3000",
    "", "", "", "", "", "", "",
  ],
  // 3. Bad CRN / invalid date / invalid email
  [
    "122956", "Sunnyvale Early Learning", "1003", "Cara", "Lee", "2020-13-40", "Active",
    "12345", "Possums", "2024-01-20", "VIC", "3000",
    "987654322C", "P-1003", "1988-06-15", "not-an-email", "Dana", "Lee", "0412345000",
  ],
  // 4. State normalisation (lowercase full name -> VIC)
  [
    "122956", "Sunnyvale Early Learning", "1004", "Evan", "Wright", "2019-09-09", "Active",
    "423456789A", "Koalas", "2024-01-10", "victoria", "3000",
    "", "", "", "", "", "", "",
  ],
  // 5 & 6. Duplicate parent across two child rows (same name+DOB+contact, different legacy ID)
  [
    "122956", "Sunnyvale Early Learning", "1005", "Finn", "Taylor", "2018-03-03", "Active",
    "523456789A", "Possums", "2024-01-05", "VIC", "3000",
    "987654323D", "P-1005", "1985-04-04", "grace.taylor@example.com", "Grace", "Taylor", "0498765432",
  ],
  [
    "122956", "Sunnyvale Early Learning", "1006", "Grace", "Taylor Jr", "2016-07-07", "Active",
    "623456789A", "Koalas", "2024-01-06", "VIC", "3000",
    "987654399Z", "P-9999", "1985-04-04", "grace.taylor@example.com", "Grace", "Taylor", "0498765432",
  ],
  // 7. QK legacy service ID mapping (182 -> 122956)
  [
    "182", "", "1007", "Holly", "Adams", "2020-01-01", "Active",
    "723456789A", "Koalas", "2024-01-08", "VIC", "3000",
    "", "", "", "", "", "", "",
  ],
];

const csvText = [HEADER, ...rows.map((r) => r.join(","))].join("\n");

async function main() {
  const serviceMap = await ServiceMapping.fromCsvText(serviceMapCsv);
  console.log("Service map loaded:", serviceMap.isLoaded);

  const file = new File([csvText], "synthetic_migration.csv", { type: "text/csv" });
  const result = await runValidation(file, serviceMap, { includeWaitlist: true });

  const { recorder, rows: allRows } = result;
  console.log("\n=== Summary ===");
  console.log("Total issues:", recorder.issues.length);
  console.log("Errors:", recorder.errorCount());
  console.log("Warnings:", recorder.warningCount());
  console.log("Fixed:", recorder.fixedCount());
  console.log("Rows surviving filter:", allRows.length);

  console.log("\n=== Expectations check ===");
  const has = (pred: (i: (typeof recorder.issues)[number]) => boolean) => recorder.issues.some(pred);

  const checks: [string, boolean][] = [
    [
      "Row 2 missing Room_Name -> ERROR",
      has((i) => i.Row === 3 && i.Field === "Room_Name" && i.Severity_Level === "ERROR"),
    ],
    [
      "Row 3 bad CRN -> ERROR",
      has((i) => i.Row === 4 && i.Field === "Child_CRN" && i.Severity_Level === "ERROR"),
    ],
    [
      "Row 3 bad DOB -> ERROR",
      has((i) => i.Row === 4 && i.Field === "DOB" && i.Severity_Level === "ERROR"),
    ],
    [
      "Row 3 bad email -> ERROR",
      has((i) => i.Row === 4 && i.Field === "Parent1_Email" && i.Severity_Level === "ERROR"),
    ],
    [
      "Row 4 state normalised 'victoria' -> 'VIC' (FIXED)",
      has((i) => i.Row === 5 && i.Field === "State" && i.Severity_Level === "FIXED"),
    ],
    [
      "Duplicate parent (rows 6 & 7) detected as FIXED intra_file_duplicate_parent",
      has((i) => i._tag === "intra_file_duplicate_parent" && (i.Row === 6 || i.Row === 7)),
    ],
    [
      "Row 8 QK ServiceID '182' mapped to Xplor ID '122956' (FIXED)",
      has((i) => i.Row === 8 && i.Field === "ServiceID" && i.Severity_Level === "FIXED"),
    ],
  ];

  let allPass = true;
  for (const [label, pass] of checks) {
    console.log(`${pass ? "PASS" : "FAIL"} — ${label}`);
    if (!pass) allPass = false;
  }

  console.log("\n=== Report generation smoke test ===");
  const { files: csvFiles, rowNumMap } = buildSplitCsvs(allRows, result.fieldnames, serviceMap);
  console.log(
    "Split CSVs:",
    csvFiles.map((f) => `${f.filename} (${f.content.split("\n").length - 1} lines)`),
  );

  const excelBuffer = await buildExcelReport(recorder, allRows, serviceMap, rowNumMap);
  console.log("Excel audit report bytes:", excelBuffer.byteLength);

  const dupReport = await buildDuplicateParentsReport(recorder, serviceMap, allRows);
  console.log(
    "Duplicate parents report bytes:",
    dupReport.buffer.byteLength,
    "cross:",
    dupReport.crossServiceCount,
    "intra groups:",
    dupReport.intraFileGroupCount,
  );

  console.log(allPass ? "\nALL CHECKS PASSED" : "\nSOME CHECKS FAILED");
  process.exit(allPass ? 0 : 1);
}

main().catch((err) => {
  console.error("Smoke test crashed:", err);
  process.exit(1);
});
