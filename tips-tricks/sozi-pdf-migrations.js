#!/usr/bin/env node
/*
 * sozi-to-pdf.js — Export Sozi (.sozi.html) presentations to PDF.
 *
 * Drives the embedded Sozi player directly (window.sozi.player), so it works
 * on presentations where `sozi-export` / `sozi-to-pdf` fails with
 * "ReferenceError: Can't find variable: sozi".
 *
 * Each Sozi frame becomes one landscape page. Inter-frame animations are
 * skipped (instant jumps) so export is fast.
 *
 * Usage:
 *   node sozi-to-pdf.js "presentation.sozi.html"
 *   node sozi-to-pdf.js *.sozi.html          # batch
 *
 * Output: same name with .pdf instead of .sozi.html, next to the input file.
 *
 * Requires Puppeteer. If you installed DeckTape you already have it; otherwise:
 *   npm install -g puppeteer
 */

const path = require("path");
const fs = require("fs");
const { pathToFileURL } = require("url");

let puppeteer;
try {
  puppeteer = require("puppeteer");
} catch (e) {
  // Fall back to puppeteer-core bundled with a global decktape install.
  try {
    puppeteer = require("puppeteer-core");
  } catch (e2) {
    console.error(
      "Could not load puppeteer. Install it with:  npm install -g puppeteer"
    );
    process.exit(1);
  }
}

const inputs = process.argv.slice(2);
if (inputs.length === 0) {
  console.error('Usage: node sozi-to-pdf.js "file.sozi.html" [more files...]');
  process.exit(1);
}

// How long to wait for the Sozi player to initialise, in ms.
const PLAYER_TIMEOUT = 30000;

async function exportFile(browser, inputPath) {
  const abs = path.resolve(inputPath);
  if (!fs.existsSync(abs)) {
    console.error(`  ! Skipping (not found): ${inputPath}`);
    return false;
  }
  const outPath = abs.replace(/\.sozi\.html$/i, "") + ".pdf";
  const url = pathToFileURL(abs).href;

  const page = await browser.newPage();
  // Surface page-side errors to help diagnose any odd file.
  page.on("pageerror", (err) => console.error("    [page error]", err.message));

  try {
    await page.goto(url, { waitUntil: "load", timeout: PLAYER_TIMEOUT });

    // Wait until the Sozi player object is ready and frames are known.
    await page.waitForFunction(
      () =>
        window.sozi &&
        window.sozi.player &&
        window.sozi.presentation &&
        window.sozi.presentation.frames &&
        window.sozi.presentation.frames.length > 0,
      { timeout: PLAYER_TIMEOUT }
    );

    // Disable per-frame auto-advance timeouts and jump to the first frame.
    const frameCount = await page.evaluate(() => {
      const p = window.sozi.presentation;
      for (const f of p.frames) {
        f.timeoutEnable = false;
      }
      window.sozi.player.moveToFrame(p.frames[0]);
      return p.frames.length;
    });

    console.log(`  ${path.basename(inputPath)} -> ${frameCount} frames`);

    const pdfBuffers = [];
    for (let i = 0; i < frameCount; i++) {
      if (i > 0) {
        // Jump straight to frame i (no animation wait).
        await page.evaluate((idx) => {
          window.sozi.player.moveToFrame(window.sozi.presentation.frames[idx]);
        }, i);
      }
      // Small settle for rendering/repaint.
      await new Promise((r) => setTimeout(r, 150));

      const buf = await page.pdf({
        printBackground: true,
        landscape: true,
        // A4 landscape; tweak if you want a different aspect ratio.
        format: "A4",
        margin: { top: 0, bottom: 0, left: 0, right: 0 },
      });
      pdfBuffers.push(buf);
      process.stdout.write(`\r    frame ${i + 1}/${frameCount}`);
    }
    process.stdout.write("\n");

    await mergePdfs(pdfBuffers, outPath);
    console.log(`  ✓ wrote ${path.basename(outPath)}`);
    return true;
  } catch (err) {
    console.error(`  ✗ failed: ${err.message}`);
    return false;
  } finally {
    await page.close();
  }
}

// Merge an array of single-page PDF buffers into one file using pdf-lib.
// (pdf-lib is pure JS — no pdfjam / TeX needed.)
async function mergePdfs(buffers, outPath) {
  const { PDFDocument } = require("pdf-lib");
  const merged = await PDFDocument.create();
  for (const buf of buffers) {
    const src = await PDFDocument.load(buf);
    const pages = await merged.copyPages(src, src.getPageIndices());
    pages.forEach((pg) => merged.addPage(pg));
  }
  const bytes = await merged.save();
  fs.writeFileSync(outPath, bytes);
}

(async () => {
  // Ensure pdf-lib is available before doing heavy work.
  try {
    require("pdf-lib");
  } catch (e) {
    console.error("Missing dependency pdf-lib. Install it with:  npm install -g pdf-lib");
    process.exit(1);
  }

  const browser = await puppeteer.launch({
    headless: "new",
    args: ["--allow-file-access-from-files", "--no-sandbox"],
  });

  let ok = 0;
  for (const input of inputs) {
    console.log(`Converting ${input} ...`);
    if (await exportFile(browser, input)) ok++;
  }
  await browser.close();
  console.log(`\nDone: ${ok}/${inputs.length} converted.`);
})();
