const assert = require("node:assert/strict");
const { chromium } = require("playwright");

const baseUrl = process.argv[2] || "http://127.0.0.1:4173/index.html";

async function textList(locator) {
  return (await locator.allTextContents()).map((value) => value.replace(/\s+/g, " ").trim());
}

async function run() {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();

  try {
    await page.goto(baseUrl, { waitUntil: "networkidle" });

    const navLabels = await textList(page.locator("#topicNav .topic-text"));
    assert.deepEqual(navLabels.slice(0, 6), [
      "Overview",
      "Part I: Planning the feature",
      "Part II: Building the core",
      "Part III: Enhancing the experience",
      "Part IV: Stabilizing the app",
      "Part V: Adding appendicular value",
    ]);

    await page.getByRole("button", { name: /Part I: Planning the feature/i }).click();
    await page.waitForURL(/chapter=part-i-planning/);
    await expectHeading(page, "Part I: Planning the feature");
    await expectBodyContains(page, "Planned vs Dispensed planning prompt");
    assert.ok((await page.locator("#outlineNav").innerText()).includes("Example 1: Planned vs Dispensed planning prompt"));

    await page.getByRole("button", { name: /Part III: Enhancing the experience/i }).click();
    await page.waitForURL(/chapter=part-iii-enhancing-experience/);
    await expectHeading(page, "Part III: Enhancing the experience");
    await expectBodyContains(page, "Nested View metric evolution");

    await page.goBack({ waitUntil: "networkidle" });
    await page.waitForURL(/chapter=part-i-planning/);
    await expectHeading(page, "Part I: Planning the feature");

    await page.getByRole("button", { name: /Overview/i }).click();
    await page.waitForURL(/chapter=overview/);
    await expectHeading(page, "Overview");
    await expectBodyContains(page, "Five-Part Learning Flow");
    assert.ok((await page.locator("#outlineNav").innerText()).includes("Part I: Planning the feature"));

    await page.goto(`${baseUrl}?chapter=part-v-adding-appendicular-value`, { waitUntil: "networkidle" });
    await expectHeading(page, "Part V: Adding appendicular value");
    await expectBodyContains(page, "Skills as workflow scaling");

    await page.goto(`${baseUrl}?doc=01-overall-development-process.md`, { waitUntil: "networkidle" });
    await expectHeading(page, "Overall Process");
    await expectBodyContains(page, "greater game plan");
  } finally {
    await browser.close();
  }
}

async function expectHeading(page, value) {
  await page.locator("#articleHeading").waitFor();
  assert.equal((await page.locator("#articleHeading").innerText()).trim(), value);
}

async function expectBodyContains(page, value) {
  const bodyText = await page.locator("#contentBody").innerText();
  assert.ok(bodyText.includes(value), `Expected content body to contain: ${value}`);
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
