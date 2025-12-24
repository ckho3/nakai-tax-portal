const ExcelJS = require('exceljs');

async function testDuplicateRow() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('========================================');
  console.log('duplicateRow 実行前の【E】セクション:');
  console.log('========================================\n');

  for (let row = 147; row <= 170; row++) {
    const cCell = sheet.getRow(row).getCell(3);
    const hCell = sheet.getRow(row).getCell(8);
    console.log(`Row ${row}: C="${cCell.value || ''}" H="${hCell.formula || hCell.value || '(空)'}"`);
  }

  console.log('\n========================================');
  console.log('duplicateRow(164, 1, true) を6回実行');
  console.log('========================================\n');

  const extraRows = 6;
  for (let i = 0; i < extraRows; i++) {
    sheet.duplicateRow(164, 1, true);
    console.log(`${i + 1}回目: Row 164を複製`);
  }

  console.log('\n========================================');
  console.log('duplicateRow 実行後の【E】セクション周辺:');
  console.log('========================================\n');

  for (let row = 147; row <= 180; row++) {
    const cCell = sheet.getRow(row).getCell(3);
    const hCell = sheet.getRow(row).getCell(8);
    console.log(`Row ${row}: C="${cCell.value || ''}" H="${hCell.formula || hCell.value || '(空)'}"`);
  }

  console.log('\n========================================');
  console.log('【E】ヘッダーの位置を検索:');
  console.log('========================================\n');

  for (let row = 147; row <= 180; row++) {
    const cCell = sheet.getRow(row).getCell(3);
    if (cCell.value && cCell.value.toString().includes('【E】')) {
      console.log(`【E】ヘッダー発見: Row ${row}`);
      console.log(`  期待値: 147 + extraRows * 4 = 147 + 24 = 171`);
      console.log(`  実際: ${row}`);
      console.log(`  差分: ${row - 171}`);
      break;
    }
  }
}

testDuplicateRow().catch(console.error);
