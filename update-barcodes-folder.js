import fs from 'node:fs';
import path from 'node:path';
import { parse } from 'csv-parse/sync';
import { stringify } from 'csv-stringify/sync';
import ExcelJS from 'exceljs';

// Updates every csv and xlsx file in bulk-update-products in place
const UPDATE_DIR = 'bulk-update-products';

// Barcode types as Shopify writes them in the "Variant Barcodes" column
const UPC_TYPE = 'upc';
const NS_PID_TYPE = 'ns-pid';

// Shopify reads ":" and ";" as separators, so literal ones must be escaped with "\"
const escapeValue = (value) => value.replace(/([\\:;])/g, '\\$1');

// Builds "upc:841949245967; ns-pid:PP016002DPOLVMedium" from the UPC and the full SKU
function buildBarcodes(barcodes, sku) {
  // Already in the typed format, leave it as it is
  if (barcodes.includes(':')) return barcodes;

  // Shopify exports numbers with a leading apostrophe to keep the zeros
  const upc = barcodes.replace(/^'/, '').trim();
  const values = [];

  if (/^\d+$/.test(upc)) values.push(`${UPC_TYPE}:${upc}`);
  else if (upc) values.push(escapeValue(upc));

  if (sku.trim()) values.push(`${NS_PID_TYPE}:${escapeValue(sku.trim())}`);

  return values.join('; ');
}

function updateCsv(inputFile, outputFile) {
  const rows = parse(fs.readFileSync(inputFile), { columns: true, bom: true });
  let updated = 0;

  for (const row of rows) {
    const barcodes = buildBarcodes(row['Variant Barcodes'], row['Variant SKU']);
    if (barcodes !== row['Variant Barcodes']) updated++;
    row['Variant Barcodes'] = barcodes;
  }

  fs.writeFileSync(outputFile, stringify(rows, { header: true }));
  return updated;
}

// Only the barcode cells are rewritten, so the zebra fill is kept
async function updateXlsx(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.getWorksheet(1);

  const headers = sheet.getRow(1).values;
  const skuCol = headers.indexOf('Variant SKU');
  const barcodesCol = headers.indexOf('Variant Barcodes');
  let updated = 0;

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const current = String(row.getCell(barcodesCol).value ?? '');
    const sku = String(row.getCell(skuCol).value ?? '');
    const barcodes = buildBarcodes(current, sku);

    if (barcodes !== current) {
      row.getCell(barcodesCol).value = barcodes;
      updated++;
    }
  });

  await workbook.xlsx.writeFile(filePath);
  return updated;
}

const listFiles = (dir, ext) =>
  fs.readdirSync(dir).filter((file) => file.endsWith(ext)).sort();

async function main() {
  const csvDir = path.join(UPDATE_DIR, 'csv');
  const xlsxDir = path.join(UPDATE_DIR, 'xlsx');
  let csvUpdated = 0;
  let xlsxUpdated = 0;

  for (const file of listFiles(csvDir, '.csv')) {
    const filePath = path.join(csvDir, file);
    csvUpdated += updateCsv(filePath, filePath);
  }

  for (const file of listFiles(xlsxDir, '.xlsx')) {
    xlsxUpdated += await updateXlsx(path.join(xlsxDir, file));
  }

  console.log(`${csvUpdated} csv rows and ${xlsxUpdated} xlsx rows updated in ${UPDATE_DIR}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
