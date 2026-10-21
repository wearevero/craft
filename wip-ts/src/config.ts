import "dotenv/config";

export const config = {
  webUrl: (process.env.WEB_URL ?? "").replace(/\/+$/, ""),
  email: process.env.EMAIL ?? "",
  password: process.env.PASSWORD ?? "",
  chromePath: process.env.CHROME_PATH ?? "",
  outputDir:
    process.env.OUTPUT_DIR ?? "D:\\Work\\CORE\\WIP\\KRIDO\\WIP TEMPLATE",
  headless: (process.env.HEADLESS ?? "true").toLowerCase() !== "false",
  concurrency: Math.max(1, Number(process.env.CONCURRENCY ?? 3) || 3),
};

export function missingConfig(): string[] {
  const required: Record<string, string> = {
    WEB_URL: config.webUrl,
    EMAIL: config.email,
    PASSWORD: config.password,
    CHROME_PATH: config.chromePath,
  };
  return Object.entries(required)
    .filter(([, v]) => !v)
    .map(([k]) => k);
}
