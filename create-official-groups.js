import fs from 'node:fs';
import path from 'node:path';
import { parse } from 'csv-parse';
import { stringify } from 'csv-stringify/sync';
import ExcelJS from 'exceljs';

const INPUT_FILES = ['products_export_1.csv', 'products_export_2.csv'];
const BACKUP_DIR = 'bulk-backup-products';
const UPDATE_DIR = 'bulk-update-products';
const PRODUCTS_WITHOUT_UPC_DIR = 'products-without-upc';
const VARIANTS_WITHOUT_UPC_DIR = 'variants-without-upc';
const PRODUCTS_PER_FILE = 100;

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

const SKU = COLUMNS.indexOf('Variant SKU');
const BARCODES = COLUMNS.indexOf('Variant Barcodes');
const OPTION1_VALUE = COLUMNS.indexOf('Option1 Value');

const WHITE = 'FFFFFFFF';
const GRAY = 'FFD9D9D9';
const RED = 'FFFF9999';

// Barcode types as Shopify writes them in the "Variant Barcodes" column
const UPC_TYPE = 'upc';
const GTIN_TYPE = 'gtin';
const NS_PID_TYPE = 'ns-pid';

// Shopify reads ":" and ";" as separators, so literal ones must be escaped with "\"
const escapeValue = (value) => value.replace(/([\\:;])/g, '\\$1');

// Shopify exports numbers with a leading apostrophe to keep the zeros
const cleanBarcode = (barcodes) => barcodes.replace(/^'/, '').trim();

// Image-only rows have no option value, SKU or barcode
const isVariantRow = (row) => row[OPTION1_VALUE] || row[SKU] || row[BARCODES];

// Numeric barcodes decide which products go in the groups, so the groups don't move
const hasUpc = (row) => /^\d+$/.test(cleanBarcode(row[BARCODES]));

// GS1 check digit, used by UPC, EAN and GTIN-14
function isValidCheckDigit(code) {
  const digits = [...code].map(Number);
  const checkDigit = digits.pop();
  const sum = digits
    .reverse()
    .reduce((total, digit, i) => total + digit * (i % 2 === 0 ? 3 : 1), 0);
  return (10 - (sum % 10)) % 10 === checkDigit;
}

// Returns "upc:..." or "gtin:..." for a valid code, or null if Shopify would reject it
function typedBarcode(barcodes) {
  let code = cleanBarcode(barcodes);
  if (!/^\d+$/.test(code)) return null;

  // UPCs that lost their leading zero, or were exported with extra zeros
  if (code.length === 11) code = `0${code}`;
  if (code.length === 14 && code.startsWith('00')) code = code.slice(2);
  if (code.length === 13 && code.startsWith('0')) code = code.slice(1);

  if (![12, 13, 14].includes(code.length) || !isValidCheckDigit(code)) return null;

  return code.length === 12 ? `${UPC_TYPE}:${code}` : `${GTIN_TYPE}:${code}`;
}

const hasValidBarcode = (row) => typedBarcode(row[BARCODES]) !== null;

// A barcode that can't be used (e.g. "ROUTEIN" or a wrong check digit) is dropped from the update
const hasInvalidBarcode = (row) => cleanBarcode(row[BARCODES]) !== '' && !hasValidBarcode(row);

// Builds "upc:841949245967; ns-pid:PP016002DPOLVMedium" from the UPC and the full SKU,
// barcodes that are not valid are left out
function buildBarcodes(barcodes, sku) {
  const barcode = typedBarcode(barcodes);
  const values = [];

  if (barcode) values.push(barcode);

  if (sku.trim()) values.push(`${NS_PID_TYPE}:${escapeValue(sku.trim())}`);

  return values.join('; ');
}

// Read every input file in order and group consecutive rows that share a handle
async function readProducts() {
  const products = [];
  let current = null;

  for (const file of INPUT_FILES) {
    const parser = fs
      .createReadStream(path.resolve(file))
      .pipe(parse({ columns: true, bom: true, relax_column_count: true }));

    for await (const record of parser) {
      const handle = record['Handle'];
      if (!handle) continue;

      const row = COLUMNS.map((column) => record[column] ?? '');
      if (current?.handle !== handle) {
        current = { handle, rows: [] };
        products.push(current);
      }
      current.rows.push(row);
    }
  }

  return products.map((product) => product.rows);
}

function chunk(items, size) {
  const chunks = [];
  for (let i = 0; i < items.length; i += size) {
    chunks.push(items.slice(i, i + size));
  }
  return chunks;
}

const withUpdatedBarcodes = (products) =>
  products.map((rows) =>
    rows.map((row) => row.with(BARCODES, buildBarcodes(row[BARCODES], row[SKU])))
  );

async function writeXlsx(filePath, products, highlight = () => false) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Products');

  sheet.columns = COLUMNS.map((header) => ({ header, key: header, width: 24 }));
  sheet.getRow(1).font = { bold: true };

  // Zebra fill per product: every row of a handle shares the same color,
  // highlighted rows are filled in red
  products.forEach((rows, index) => {
    const zebra = index % 2 === 0 ? GRAY : WHITE;
    for (const values of rows) {
      const color = highlight(values) ? RED : zebra;
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

// Writes the products as name.csv and name.xlsx inside dir/csv and dir/xlsx
async function writeFiles(dir, name, products, highlight) {
  const csvDir = path.join(dir, 'csv');
  const xlsxDir = path.join(dir, 'xlsx');
  fs.mkdirSync(csvDir, { recursive: true });
  fs.mkdirSync(xlsxDir, { recursive: true });

  fs.writeFileSync(path.join(csvDir, `${name}.csv`), stringify([COLUMNS, ...products.flat()]));
  await writeXlsx(path.join(xlsxDir, `${name}.xlsx`), products, highlight);
}

// Removes the generated files so groups from a previous run don't stay behind
function cleanDir(dir) {
  for (const sub of ['csv', 'xlsx']) {
    fs.rmSync(path.join(dir, sub), { recursive: true, force: true });
  }
}

async function main() {
  const products = await readProducts();

  const withUpc = [];
  const withoutUpc = [];
  const variantsWithoutUpc = [];

  for (const rows of products) {
    const variants = rows.filter(isVariantRow);

    if (!variants.some(hasUpc)) {
      withoutUpc.push(rows);
      continue;
    }

    withUpc.push(rows);
    const missing = variants.filter((row) => !hasValidBarcode(row));
    if (missing.length) variantsWithoutUpc.push(missing);
  }

  const groups = chunk(withUpc, PRODUCTS_PER_FILE);
  const pad = String(groups.length).length;

  for (const dir of [BACKUP_DIR, UPDATE_DIR, PRODUCTS_WITHOUT_UPC_DIR, VARIANTS_WITHOUT_UPC_DIR]) {
    cleanDir(dir);
  }

  for (const [index, group] of groups.entries()) {
    const name = `products_group_${String(index + 1).padStart(pad, '0')}`;
    await writeFiles(BACKUP_DIR, name, group);
    await writeFiles(UPDATE_DIR, name, withUpdatedBarcodes(group));
  }

  await writeFiles(PRODUCTS_WITHOUT_UPC_DIR, 'products_without_upc', withoutUpc);
  await writeFiles(VARIANTS_WITHOUT_UPC_DIR, 'variants_without_upc', variantsWithoutUpc, hasInvalidBarcode);

  const countRows = (list) => list.reduce((total, rows) => total + rows.length, 0);
  console.log(`${withUpc.length} products (${countRows(withUpc)} rows) -> ${groups.length} groups in ${BACKUP_DIR}, ${UPDATE_DIR}`);
  console.log(`${withoutUpc.length} products (${countRows(withoutUpc)} rows) -> ${PRODUCTS_WITHOUT_UPC_DIR}`);
  console.log(`${countRows(variantsWithoutUpc)} variants from ${variantsWithoutUpc.length} products -> ${VARIANTS_WITHOUT_UPC_DIR}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
