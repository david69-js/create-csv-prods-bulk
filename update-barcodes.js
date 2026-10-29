import fs from 'node:fs';
import path from 'node:path';
import { parse } from 'csv-parse/sync';
import { stringify } from 'csv-stringify/sync';

const INPUT_FILE = process.argv[2] ?? 'bulk-update-products/csv/products_group_01.csv';
const OUTPUT_DIR = 'test-bulk-update-products';

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

function main() {
  const rows = parse(fs.readFileSync(INPUT_FILE), { columns: true, bom: true });
  let updated = 0;

  for (const row of rows) {
    const barcodes = buildBarcodes(row['Variant Barcodes'], row['Variant SKU']);
    if (barcodes !== row['Variant Barcodes']) updated++;
    row['Variant Barcodes'] = barcodes;
  }

  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  const outputFile = path.join(OUTPUT_DIR, path.basename(INPUT_FILE));
  fs.writeFileSync(outputFile, stringify(rows, { header: true }));

  console.log(`${updated} of ${rows.length} rows updated -> ${outputFile}`);
}

main();
