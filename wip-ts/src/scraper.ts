import puppeteer, { Browser, Page } from "puppeteer-core";
import pc from "picocolors";
import { createInterface } from "node:readline/promises";
import { config } from "./config";
import type { JobConfig } from "./jobs.js";

export type Row = string[];

export async function createBrowser(): Promise<Browser> {
  return puppeteer.launch({
    executablePath: config.chromePath,
    headless: config.headless,
    defaultViewport: config.headless ? { width: 1366, height: 900 } : null,
    args: config.headless ? [] : ["--start-maximized"],
  });
}

export async function login(browser: Browser): Promise<Page> {
  console.log(pc.dim("› Logging in..."));
  const [page] = await browser.pages();
  try {
    await page.goto(config.webUrl, { waitUntil: "domcontentloaded" });
    await page.waitForSelector('[name="email"]', { timeout: 15000 });
    await page.type('[name="email"]', config.email);
    await page.type('[name="password"]', config.password);
    await Promise.all([
      page
        .waitForNavigation({ waitUntil: "domcontentloaded", timeout: 15000 })
        .catch(() => {}),
      page.keyboard.press("Enter"),
    ]);
    console.log(`${pc.green("✓")} Logged in\n`);
  } catch (e) {
    console.log(`${pc.red("✗")} Login failed`);
    console.log(pc.dim(`  ${(e as Error).message}`));
    throw e;
  }
  return page;
}

const TABLE_SELECTORS = [
  "td.judul > table > tbody",
  "table > tbody",
  ".judul table tbody",
  "table tbody",
];

/** Ambil isi tabel (tanpa header & footer). Highlight baris hanya jika browser tampil. */
async function extractTableData(
  page: Page,
  highlight: boolean,
): Promise<Row[]> {
  return page.evaluate(
    (selectors, hl) => {
      let tbody: Element | null = null;
      for (const s of selectors) {
        tbody = document.querySelector(s);
        if (tbody) break;
      }
      if (!tbody) return [];

      const trs = Array.from(tbody.getElementsByTagName("tr"));
      if (trs.length <= 2) return [];

      const dataRows = trs.slice(1, -1);

      if (hl) {
        for (const tr of dataRows) {
          tr.style.border = "2px solid red";
          tr.style.backgroundColor = "yellow";
        }
        setTimeout(() => {
          for (const tr of dataRows) {
            tr.style.border = "";
            tr.style.backgroundColor = "";
          }
        }, 1500);
      }

      const out: string[][] = [];
      for (const tr of dataRows) {
        const tds = Array.from(tr.getElementsByTagName("td"));
        if (tds.length) out.push(tds.map((td) => (td.innerText ?? "").trim()));
      }
      return out;
    },
    TABLE_SELECTORS,
    highlight,
  );
}

type SectionResult =
  | { status: "ok"; rows: Row[] }
  | { status: "empty" }
  | { status: "error"; message: string };

async function scrapeSection(
  browser: Browser,
  url: string,
  job: JobConfig,
  bagian: string,
): Promise<SectionResult> {
  const page = await browser.newPage();
  try {
    const res = await page.goto(url, {
      waitUntil: "domcontentloaded",
      timeout: 30000,
    });
    await page.waitForSelector("table", { timeout: 10000 }).catch(() => {});

    const title = await page.title();
    const html = (await page.content()).toLowerCase();
    if (
      (res && res.status() >= 400) ||
      html.includes("error") ||
      title.includes("404")
    ) {
      return { status: "error", message: "page error" };
    }

    const raw = await extractTableData(page, !config.headless);
    if (!raw.length) return { status: "empty" };

    const rows = raw.map((r) => {
      const cut = job.dropLastCols > 0 ? r.slice(0, -job.dropLastCols) : r;
      return job.prependSection ? [bagian, ...cut] : cut;
    });
    return { status: "ok", rows };
  } catch (e) {
    return { status: "error", message: "failed to load" };
  } finally {
    await page.close().catch(() => {});
  }
}

export function generateUrls(
  job: JobConfig,
  tanggal: string,
): Array<[string, string]> {
  return Object.entries(job.sections).map(([bagian, id]) => [
    bagian,
    `${config.webUrl}/${job.basePath}?d=${tanggal}&s=&b=${id}&m=all`,
  ]);
}

export async function collectData(
  browser: Browser,
  job: JobConfig,
  tanggal: string,
): Promise<Row[]> {
  const entries = generateUrls(job, tanggal);
  console.log(
    `${pc.bold(pc.white(job.label))} ${pc.dim(`· ${entries.length} sections`)}\n`,
  );

  const results: SectionResult[] = new Array(entries.length);
  let next = 0;

  const worker = async () => {
    while (true) {
      const i = next++;
      if (i >= entries.length) return;
      const [bagian, url] = entries[i];
      const r = await scrapeSection(browser, url, job, bagian);
      results[i] = r;

      const name = bagian.padEnd(20);
      if (r.status === "ok")
        console.log(
          `  ${pc.green("✓")} ${name} ${pc.dim(`${r.rows.length} rows`)}`,
        );
      else if (r.status === "empty")
        console.log(`  ${pc.dim("·")} ${pc.dim(`${name} no data`)}`);
      else console.log(`  ${pc.yellow("‼")} ${name} ${pc.dim(r.message)}`);
    }
  };

  await Promise.all(
    Array.from(
      { length: Math.min(config.concurrency, entries.length) },
      worker,
    ),
  );

  // Susun sesuai urutan bagian asli, bukan urutan selesai
  const all: Row[] = [];
  for (const r of results) if (r.status === "ok") all.push(...r.rows);

  console.log();
  if (all.length)
    console.log(`${pc.dim("total")}  ${pc.bold(String(all.length))} rows`);
  else console.log(pc.dim(`no data collected for ${job.label}`));
  console.log();

  return all;
}

export async function getTanggalInput(): Promise<string> {
  const rl = createInterface({ input: process.stdin, output: process.stdout });
  try {
    while (true) {
      const ans = (
        await rl.question("date (YYYY-MM-DD, default: today) › ")
      ).trim();
      if (ans === "") {
        const d = new Date();
        const pad = (n: number) => String(n).padStart(2, "0");
        return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
      }
      if (/^\d{4}-\d{2}-\d{2}$/.test(ans)) return ans;
      console.log(pc.yellow("  format harus YYYY-MM-DD"));
    }
  } finally {
    rl.close();
  }
}
