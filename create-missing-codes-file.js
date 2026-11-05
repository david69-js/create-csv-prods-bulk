import fs from 'node:fs';
import path from 'node:path';
import { parse } from 'csv-parse';
import { stringify } from 'csv-stringify/sync';
import ExcelJS from 'exceljs';

// Lists every variant that ended up without UPC and NS-PID in Shopify, to be filled in manually:
// - products without any UPC, which were left out of the upload groups
// - products that Shopify rejected during the import because of an invalid barcode
const INPUT_FILES = ['products_export_1.csv', 'products_export_2.csv'];
const OUTPUT_DIR = 'variants-without-upc-and-ns-pid';
const OUTPUT_NAME = 'variants_without_upc_and_ns_pid';

// Handles from the Shopify import reports (groups 02, 28-30, 33-38 and 40)
const FAILED_IMPORT_HANDLES = new Set([
  'olympic-lichen-shirt-jacket',
  'gramicci-x-parks-project-reversible-bucket-hat',
  'guppyfriend-washing-bag',
  'igloo-x-parks-project-recycled-playmate-elite-stickers',
  'lets-play-outdoors',
  'hk-x-parks-project-utah-roundup-kerchief',
  'parks-book',
  'yosemite-stoneware-mug',
  'state-parks-of-the-lost-coast-tee',
  'state-parks-of-california-tee',
  'texas-all-parks-roundup-long-sleeve-tee',
  'leave-it-better-stamped-tee',
  'parks-for-lovers-woodstock-womens-tee',
  'americas-national-parks',
  'the-outsiders',
  'national-parks-are-for-lovers-ringer',
]);

const COLUMNS = [
  'Handle',
  'Title',
  'Option1 Name',
  'Option1 Value',
  'Option2 Name',
  'Option2 Value',
  'Option3 Name',
  'Option3 Value',
  'Variant SKU',
  'Variant Barcodes',
];

const BARCODES = COLUMNS.indexOf('Variant Barcodes');

const WHITE = 'FFFFFFFF';
const GRAY = 'FFD9D9D9';
const RED = 'FFFF9999';

// Shopify only accepts UPCs with these lengths
const UPC_LENGTHS = [8, 12];

// Shopify exports numbers with a leading apostrophe to keep the zeros
const cleanBarcode = (barcodes) => barcodes.replace(/^'/, '').trim();

// Image-only rows have no option value, SKU or barcode
const isVariantRow = (record) =>
  record['Option1 Value'] || record['Variant SKU'] || record['Variant Barcodes'];

// Same rule as create-official-groups.js: a numeric barcode counts as a UPC
const hasUpc = (record) => /^\d+$/.test(cleanBarcode(record['Variant Barcodes']));

// GS1 check digit, used by UPC, EAN and GTIN-14
function isValidCheckDigit(code) {
  const digits = [...code].map(Number);
  const checkDigit = digits.pop();
  const sum = digits
    .reverse()
    .reduce((total, digit, i) => total + digit * (i % 2 === 0 ? 3 : 1), 0);
  return (10 - (sum % 10)) % 10 === checkDigit;
}

// A barcode Shopify rejects as UPC, highlighted in red in the xlsx
function isInvalidBarcode(row) {
  const code = cleanBarcode(row[BARCODES]);
  if (!code) return false;
  if (!/^\d+$/.test(code)) return true;
  return !UPC_LENGTHS.includes(code.length) || !isValidCheckDigit(code);
}

// Read every input file in order and group consecutive variant rows that share a handle
async function readVariantsByProduct() {
  const products = [];
  let current = null;

  for (const file of INPUT_FILES) {
    const parser = fs
      .createReadStream(path.resolve(file))
      .pipe(parse({ columns: true, bom: true, relax_column_count: true }));

    for await (const record of parser) {
      const handle = record['Handle'];
      if (!handle || !isVariantRow(record)) continue;

      if (current?.handle !== handle) {
        current = { handle, variants: [] };
        products.push(current);
      }
      current.variants.push(record);
    }
  }

  return products;
}

async function writeXlsx(filePath, products) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Products');

  sheet.columns = COLUMNS.map((header) => ({ header, key: header, width: 24 }));
  sheet.getRow(1).font = { bold: true };

  // Zebra fill per product: every row of a handle shares the same color,
  // rows with an invalid barcode are filled in red
  products.forEach((rows, index) => {
    const zebra = index % 2 === 0 ? GRAY : WHITE;
    for (const values of rows) {
      const color = isInvalidBarcode(values) ? RED : zebra;
      const row = sheet.addRow(values);
      for (let col = 1; col <= COLUMNS.length; col++) {
        row.getCell(col).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: color },
        };
      }
    }
  });

  await workbook.xlsx.writeFile(filePath);
}

async function main() {
  const selected = (await readVariantsByProduct()).filter(
    (product) => !product.variants.some(hasUpc) || FAILED_IMPORT_HANDLES.has(product.handle)
  );

  const products = selected.map((product) =>
    product.variants.map((record) => COLUMNS.map((column) => record[column] ?? ''))
  );

  const csvDir = path.join(OUTPUT_DIR, 'csv');
  const xlsxDir = path.join(OUTPUT_DIR, 'xlsx');
  fs.mkdirSync(csvDir, { recursive: true });
  fs.mkdirSync(xlsxDir, { recursive: true });

  fs.writeFileSync(path.join(csvDir, `${OUTPUT_NAME}.csv`), stringify([COLUMNS, ...products.flat()]));
  await writeXlsx(path.join(xlsxDir, `${OUTPUT_NAME}.xlsx`), products);

  const failed = selected.filter((product) => FAILED_IMPORT_HANDLES.has(product.handle)).length;
  console.log(
    `${products.flat().length} variants from ${selected.length} products ` +
      `(${selected.length - failed} without UPC, ${failed} failed import) -> ${OUTPUT_DIR}`
  );
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
