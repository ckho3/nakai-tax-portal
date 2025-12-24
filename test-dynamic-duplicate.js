const ExcelJS = require('exceljs');

async function testDynamicDuplicate() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('========================================');
  console.log('動的な行番号で複製（修正後のロジック）');
  console.log('========================================\n');

  const extraRows = 7; // 27件PDF

  console.log(`27件PDF（extraRows = ${extraRows}）で実行:`);
  console.log('');

  for (let i = 0; i < extraRows; i++) {
    const eRowToDuplicate = 164 + i * 4;
    const dRowToDuplicate = 143 + i * 3;
    const cRowToDuplicate = 120 + i * 2;
    const bRowToDuplicate = 97 + i;

    sheet.duplicateRow(eRowToDuplicate, 1, true);
    sheet.duplicateRow(dRowToDuplicate, 1, true);
    sheet.duplicateRow(cRowToDuplicate, 1, true);
    sheet.duplicateRow(bRowToDuplicate, 1, true);
    sheet.duplicateRow(73, 1, true);

    if (i < 3 || i === extraRows - 1) {
      console.log(`${i + 1}回目: E=${eRowToDuplicate}, D=${dRowToDuplicate}, C=${cRowToDuplicate}, B=${bRowToDuplicate}, A=73`);
    } else if (i === 3) {
      console.log('...');
    }
  }

  console.log('');
  console.log('【E】ヘッダーの位置を検索:');
  for (let row = 147; row <= 185; row++) {
    const cCell = sheet.getRow(row).getCell(3);
    const cValue = cCell.value || '';
    if (cValue.toString().includes('【E】')) {
      console.log(`  実際の位置: Row ${row}`);
      console.log(`  計算式（旧）: 147 + ${extraRows} * 4 + 1 = ${147 + extraRows * 4 + 1}`);
      console.log(`  計算式（新）: 147 + ${extraRows} * 4 = ${147 + extraRows * 4}`);
      console.log(`  一致: ${row === 147 + extraRows * 4 ? '✅' : '❌'}`);
      break;
    }
  }
}

testDynamicDuplicate().catch(console.error);
