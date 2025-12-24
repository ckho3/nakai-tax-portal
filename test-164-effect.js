const ExcelJS = require('exceljs');

async function test164Effect() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('初期状態:');
  console.log(`  Row 146: ${sheet.getRow(146).getCell(3).value || ''}`);
  console.log(`  Row 147: ${sheet.getRow(147).getCell(3).value || ''}`);
  console.log('');

  console.log('duplicateRow(164, 1, true) を1回実行:');
  sheet.duplicateRow(164, 1, true);

  console.log(`  Row 146: ${sheet.getRow(146).getCell(3).value || ''}`);
  console.log(`  Row 147: ${sheet.getRow(147).getCell(3).value || ''}`);
  console.log(`  Row 148: ${sheet.getRow(148).getCell(3).value || ''}`);
  console.log('');

  console.log('→ Row 147に変化がありますか？');
}

test164Effect().catch(console.error);
