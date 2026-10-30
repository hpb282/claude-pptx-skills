#!/usr/bin/env node
/**
 * canvas_to_gif.js — chụp từng khung của một hoạt họa HTML Canvas rồi ghép thành GIF.
 *
 *   node canvas_to_gif.js page.html out.gif [--fps 25] [--seconds 3] [--width 960] [--height 540]
 *
 * Trang HTML phải có <canvas id="stage"> và hàm toàn cục window.renderFrame(t)
 * vẽ khung tại thời điểm t ∈ [0, 1). Không dùng requestAnimationFrame hay
 * Date.now(): mỗi khung phải xác định (deterministic) để GIF không giật.
 *
 * Cần: playwright (npm i -D playwright && npx playwright install chromium)
 *      python3 + Pillow (bước ghép GIF dùng render_gif.py preset "frames").
 */
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");
const { chromium } = require("playwright");

function arg(name, def) {
  const i = process.argv.indexOf(`--${name}`);
  return i > -1 ? Number(process.argv[i + 1]) : def;
}

(async () => {
  const [html, out] = process.argv.slice(2).filter((a, i, all) => !a.startsWith("--") && !all[i - 1]?.startsWith("--"));
  if (!html || !out) {
    console.error("Cách dùng: node canvas_to_gif.js page.html out.gif [--fps 25] [--seconds 3]");
    process.exit(1);
  }
  const fps = arg("fps", 25);
  const seconds = arg("seconds", 3);
  const width = arg("width", 960);
  const height = arg("height", 540);
  const n = Math.round(fps * seconds);
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "frames-"));

  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width, height }, deviceScaleFactor: 1 });
  await page.goto("file://" + path.resolve(html));
  await page.waitForFunction(() => typeof window.renderFrame === "function");
  const canvas = page.locator("#stage");

  for (let i = 0; i < n; i++) {
    await page.evaluate((t) => window.renderFrame(t), i / n);
    await canvas.screenshot({ path: path.join(dir, `f${String(i).padStart(4, "0")}.png`) });
  }
  await browser.close();

  execFileSync("python3", [path.join(__dirname, "render_gif.py"), "frames", "--input", dir, "--out", out, "--fps", String(fps)], {
    stdio: "inherit",
  });
  fs.rmSync(dir, { recursive: true, force: true });
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
