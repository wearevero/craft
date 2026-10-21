import ExcelJS from "exceljs";

const NUMERIC_RE = /^-?\d+\.?\d*$/;

export function isNumericValue(s: string): boolean {
  return NUMERIC_RE.test(s.replace(/,/g, ""));
}

/** Simpan data ke Excel; angka dikonversi & format desimal dipertahankan seperti di web. */
export async function saveToExcel(
  data: string[][],
  filename: string,
  numericStartCol: number,
): Promise<void> {
  if (!data.length) return;

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");
  ws.addRows(data);

  ws.eachRow((row) => {
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (colNumber < numericStartCol) return;

      const raw = String(cell.value ?? "").trim();
      if (!raw || !isNumericValue(raw)) return;

      const num = Number(raw.replace(/,/g, ""));
      if (Number.isNaN(num)) return;

      cell.value = num;

      if (raw.includes(".")) {
        const dec = raw.split(".").pop()!.length;
        cell.numFmt = dec === 0 ? "0" : "0." + "0".repeat(dec);
      } else {
        cell.numFmt = Number.isInteger(num) ? "0" : "0.00";
      }
    });
  });

  await wb.xlsx.writeFile(filename);
}
