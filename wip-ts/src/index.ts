import fs from "node:fs";
import path from "node:path";
import pc from "picocolors";
import { config, missingConfig } from "./config";
import { jobs } from "./jobs.js";
import { collectData, createBrowser, getTanggalInput, login } from "./scraper";
import { saveToExcel } from "./excel";

async function main() {
  const jobName = process.argv[2];
  const job = jobs[jobName];
  if (!job) {
    console.log(`Usage: npm run <${Object.keys(jobs).join("|")}>`);
    process.exit(1);
  }

  const missing = missingConfig();
  if (missing.length) {
    console.log(`${pc.red("✗")} .env belum lengkap: ${missing.join(", ")}`);
    process.exit(1);
  }

  console.log(
    `\n${pc.bold(pc.white(job.label))} ${pc.dim("· data scraper")}\n`,
  );
  const argDate = process.argv[3];
  const tanggal =
    argDate && /^\d{4}-\d{2}-\d{2}$/.test(argDate)
      ? argDate
      : await getTanggalInput();

  const browser = await createBrowser();
  try {
    await login(browser);
    const data = await collectData(browser, job, tanggal);

    if (!data.length) {
      console.log(pc.yellow("✗ No data found"));
      console.log(pc.dim("  · date may have no records"));
      console.log(pc.dim("  · page structure may have changed"));
      console.log(pc.dim("  · connection timeout or failed to load\n"));
      return;
    }

    fs.mkdirSync(config.outputDir, { recursive: true });
    const filename = path.join(
      config.outputDir,
      `${job.filePrefix}-${tanggal}.xlsx`,
    );
    await saveToExcel(data, filename, job.numericStartCol);

    if (fs.existsSync(filename)) {
      const kb = fs.statSync(filename).size / 1024;
      console.log(pc.green("✓") + " Saved");
      console.log(`${pc.dim("  file")}  ${path.basename(filename)}`);
      console.log(`${pc.dim("  rows")}  ${data.length}`);
      console.log(`${pc.dim("  size")}  ${kb.toFixed(1)} KB\n`);
    } else {
      console.log(
        `${pc.red("✗")} Failed to save file\n${pc.dim(`  ${filename}`)}\n`,
      );
    }
  } catch (e) {
    process.exitCode = 1;
    console.log(`${pc.red("✗")} Something went wrong`);
    console.log(pc.dim(`  ${(e as Error).message}\n`));
    console.error(e);
  } finally {
    if (config.headless) {
      await browser.close();
    } else {
      console.log(pc.dim("Browser left open for manual inspection."));
      browser.disconnect();
    }
  }
}

main();
