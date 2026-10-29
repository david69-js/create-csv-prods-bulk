import fs from 'node:fs';
import path from 'node:path';
import { parse } from 'csv-parse/sync';
import { stringify } from 'csv-stringify/sync';

const PRODUCTS = 20;
const FILE = 'products_group_01.csv';
const OUTPUT_DIR = 'test-20-products';

const SOURCES = {
  update: path.join('bulk-update-products', 'csv', FILE),
  backup: path.join('bulk-backup-products', 'csv', FILE),
};

// Keeps every row of the first N handles, in the original order
function firstProducts(rows, count) {
  const handles = new Set();
  return rows.filter((row) => {
    if (!handles.has(row['Handle'])) {
      if (handles.size === count) return false;
      handles.add(row['Handle']);
    }
    return true;
  });
}

fs.mkdirSync(OUTPUT_DIR, { recursive: true });

for (const [name, source] of Object.entries(SOURCES)) {
  const rows = parse(fs.readFileSync(source), { columns: true, bom: true });
  const sample = firstProducts(rows, PRODUCTS);
  const outputFile = path.join(OUTPUT_DIR, `${name}_products_${PRODUCTS}.csv`);

  fs.writeFileSync(outputFile, stringify(sample, { header: true }));
  console.log(`${sample.length} rows (${PRODUCTS} products) -> ${outputFile}`);
}
