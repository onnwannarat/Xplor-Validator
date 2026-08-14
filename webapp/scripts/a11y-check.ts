import { chromium } from "playwright";
import path from "node:path";

const OUT_DIR = path.join(__dirname, "..", ".visual-check");

async function main() {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });

  await page.goto("http://localhost:3000/validator", { waitUntil: "networkidle" });

  // Tab through the interactive controls on the page and record what receives focus.
  const focusedLabels: string[] = [];
  for (let i = 0; i < 12; i++) {
    await page.keyboard.press("Tab");
    const label = await page.evaluate(() => {
      const el = document.activeElement as HTMLElement | null;
      if (!el) return "(none)";
      return (
        el.getAttribute("aria-label") ||
        el.textContent?.trim().slice(0, 40) ||
        el.tagName.toLowerCase()
      );
    });
    focusedLabels.push(label);
  }
  console.log("Tab order (first 12 stops):");
  focusedLabels.forEach((l, i) => console.log(`  ${i + 1}. ${l}`));

  // Activate the "Active + Waitlist" scope toggle via keyboard only.
  await page.goto("http://localhost:3000/validator", { waitUntil: "networkidle" });
  const scopeButton = page.getByRole("button", { name: "Active + Waitlist" });
  await scopeButton.focus();
  await page.keyboard.press("Enter");
  const pressed = await scopeButton.getAttribute("aria-pressed");
  console.log(`\n"Active + Waitlist" toggled via keyboard, aria-pressed=${pressed}`);

  await page.screenshot({ path: path.join(OUT_DIR, "05-keyboard-focus.png") });

  await browser.close();
  process.exit(pressed === "true" ? 0 : 1);
}

main().catch((err) => {
  console.error("A11y check crashed:", err);
  process.exit(1);
});
