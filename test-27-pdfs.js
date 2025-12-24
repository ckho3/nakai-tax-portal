const ExcelJS = require('exceljs');

async function test27PDFs() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('========================================');
  console.log('27件PDF（extraRows = 7）の行追加テスト');
  console.log('========================================\n');

  const extraRows = 7;

  console.log('実行前: Row 147 =', sheet.getRow(147).getCell(3).value);
  console.log('');

  for (let i = 0; i < extraRows; i++) {
    sheet.duplicateRow(164, 1, true);
    sheet.duplicateRow(143, 1, true);
    sheet.duplicateRow(120, 1, true);
    sheet.duplicateRow(97, 1, true);
    sheet.duplicateRow(73, 1, true);
  }

  console.log('実行後、【E】ヘッダーの位置を検索:');
  for (let row = 147; row <= 185; row++) {
    const cCell = sheet.getRow(row).getCell(3);
    const cValue = cCell.value || '';
    if (cValue.toString().includes('【E】')) {
      console.log(`  【E】ヘッダー: Row ${row}`);
      console.log(`  計算式1: 147 + ${extraRows} * 4 = ${147 + extraRows * 4}`);
      console.log(`  計算式2: 147 + ${extraRows} * 4 + 1 = ${147 + extraRows * 4 + 1}`);
      console.log(`  実際との差: ${row - (147 + extraRows * 4)} 行`);
      break;
    }
  }
}

test27PDFs().catch(console.error);
