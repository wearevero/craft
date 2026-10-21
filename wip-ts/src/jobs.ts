export interface JobConfig {
  /** Label di console */
  label: string;
  /** Path laporan di web */
  basePath: string;
  /** Nama bagian -> id bagian di sistem */
  sections: Record<string, number>;
  /** Prefix nama file output, mis. "loss" -> loss-2025-01-31.xlsx */
  filePrefix: string;
  /** Tambahkan nama bagian sebagai kolom pertama tiap baris */
  prependSection: boolean;
  /** Buang N kolom terakhir dari tiap baris */
  dropLastCols: number;
  /** Kolom (1-based) pertama yang dicek untuk dikonversi jadi angka */
  numericStartCol: number;
}

const SECTIONS_ALL: Record<string, number> = {
  "CUTTING 2": 7,
  "TAMBAH PART": 125,
  "RE-CASTING": 101,
  "REPAIR PART": 126,
  ILCA: 103,
  STRIPING: 9,
  PENDING: 15,
  PERBAIKAN: 16,
  "RANGKAI 1": 12,
  "SEGONG REPAIR": 2,
  FILLLING1: 13,
  "FILLING 2": 17,
  "POLISHING 1": 100,
  "POLISHING 2": 11,
  "POLISHING CVD": 122,
};

export const jobs: Record<string, JobConfig> = {
  filling: {
    label: "FILLING TEM",
    basePath: "laporan/loss_bagian_cetak",
    sections: { "FILLING TEM": 8 },
    filePrefix: "filling",
    prependSection: true,
    dropLastCols: 0,
    numericStartCol: 2,
  },
  komponen: {
    label: "KOMPONEN",
    basePath: "laporan/komponen_cetak",
    sections: SECTIONS_ALL,
    filePrefix: "komponen",
    prependSection: false,
    dropLastCols: 2, // buang kolom BAGIAN & TANGGAL
    numericStartCol: 1,
  },
  loss: {
    label: "LOSS",
    basePath: "laporan/loss_bagian_cetak",
    sections: SECTIONS_ALL,
    filePrefix: "loss",
    prependSection: true,
    dropLastCols: 0,
    numericStartCol: 2,
  },
};
