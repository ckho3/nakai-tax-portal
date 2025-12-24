const ExcelJS = require('exceljs');

async function testDuplicateOrder() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('初期位置: Row 147のC列 =', sheet.getRow(147).getCell(3).value);
  console.log('');

  // 1回だけ実行
  console.log('---1回目の複製開始---');

  console.log('1. duplicateRow(164)');
  sheet.duplicateRow(164, 1, true);
  let e147 = findEHeader(sheet);
  console.log(`   → 【E】ヘッダー: Row ${e147}`);

  console.log('2. duplicateRow(143)');
  sheet.duplicateRow(143, 1, true);
  e147 = findEHeader(sheet);
  console.log(`   → 【E】ヘッダー: Row ${e147}`);

  console.log('3. duplicateRow(120)');
  sheet.duplicateRow(120, 1, true);
  e147 = findEHeader(sheet);
  console.log(`   → 【E】ヘッダー: Row ${e147}`);

  console.log('4. duplicateRow(97)');
  sheet.duplicateRow(97, 1, true);
  e147 = findEHeader(sheet);
  console.log(`   → 【E】ヘッダー: Row ${e147}`);

  console.log('5. duplicateRow(73)');
  sheet.duplicateRow(73, 1, true);
  e147 = findEHeader(sheet);
  console.log(`   → 【E】ヘッダー: Row ${e147}`);

  console.log('');
  console.log('1回目終了後: Row 147 → Row', e147, `(+${e147 - 147}行)`);
}

function findEHeader(sheet) {
  for (let row = 147; row <= 200; row++) {
    const cCell = sheet.getRow(row).getCell(3);
    if (cCell.value && cCell.value.toString().includes('【E】')) {
      return row;
    }
  }
  return -1;
}

testDuplicateOrder().catch(console.error);
