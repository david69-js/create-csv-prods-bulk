import fs from 'node:fs';
import path from 'node:path';
import { parse } from 'csv-parse';
import { stringify } from 'csv-stringify/sync';
import ExcelJS from 'exceljs';

const INPUT_FILES = ['products_export_1.csv', 'products_export_2.csv'];
const OUTPUT_DIRS = ['bulk-backup-products', 'bulk-update-products'];
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

const WHITE = 'FFFFFFFF';
const GRAY = 'FFD9D9D9';

// Read every input file in order and group consecutive rows that share a handle,
// so the original row order is never changed
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

async function writeXlsx(filePath, group) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Products');

  sheet.columns = COLUMNS.map((header) => ({ header, key: header, width: 24 }));
  sheet.getRow(1).font = { bold: true };

  // Zebra fill per product: every row of a handle shares the same color
  group.forEach((rows, index) => {
    const color = index % 2 === 0 ? GRAY : WHITE;
    for (const values of rows) {
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
  const products = await readProducts();
  const groups = chunk(products, PRODUCTS_PER_FILE);
  const pad = String(groups.length).length;

  for (const dir of OUTPUT_DIRS) {
    const csvDir = path.join(dir, 'csv');
    const xlsxDir = path.join(dir, 'xlsx');
    fs.mkdirSync(csvDir, { recursive: true });
    fs.mkdirSync(xlsxDir, { recursive: true });

    for (const [index, group] of groups.entries()) {
      const name = `products_group_${String(index + 1).padStart(pad, '0')}`;
      const csv = stringify([COLUMNS, ...group.flat()]);

      fs.writeFileSync(path.join(csvDir, `${name}.csv`), csv);
      await writeXlsx(path.join(xlsxDir, `${name}.xlsx`), group);
    }
  }

  const rowCount = products.reduce((total, rows) => total + rows.length, 0);
  console.log(
    `${products.length} products (${rowCount} rows) -> ${groups.length} groups written to ${OUTPUT_DIRS.join(', ')}`
  );
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
