const ExcelJS = require('exceljs');

async function testAllDuplicate() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('========================================');
  console.log('全セクションの行追加を実行（26件PDF想定）');
  console.log('========================================\n');

  const extraRows = 6;

  console.log('実行前の主要行:');
  console.log(`  Row 73: C="${sheet.getRow(73).getCell(3).value || ''}"`);
  console.log(`  Row 97: C="${sheet.getRow(97).getCell(3).value || ''}"`);
  console.log(`  Row 120: C="${sheet.getRow(120).getCell(3).value || ''}"`);
  console.log(`  Row 143: C="${sheet.getRow(143).getCell(3).value || ''}"`);
  console.log(`  Row 147: C="${sheet.getRow(147).getCell(3).value || ''}"`);
  console.log(`  Row 164: C="${sheet.getRow(164).getCell(3).value || ''}" H="${sheet.getRow(164).getCell(8).value || sheet.getRow(164).getCell(8).formula || ''}"`);
  console.log('');

  for (let i = 0; i < extraRows; i++) {
    console.log(`--- ${i + 1}回目の複製 ---`);
    sheet.duplicateRow(164, 1, true);
    console.log(`  【E】Row 164を複製`);

    sheet.duplicateRow(143, 1, true);
    console.log(`  【D】Row 143を複製`);

    sheet.duplicateRow(120, 1, true);
    console.log(`  【C】Row 120を複製`);

    sheet.duplicateRow(97, 1, true);
    console.log(`  【B】Row 97を複製`);

    sheet.duplicateRow(73, 1, true);
    console.log(`  【A】Row 73を複製`);
  }

  console.log('\n========================================');
  console.log('実行後の【E】セクション周辺を検索:');
  console.log('========================================\n');

  for (let row = 165; row <= 180; row++) {
    const cCell = sheet.getRow(row).getCell(3);
    const hCell = sheet.getRow(row).getCell(8);
    const cValue = cCell.value || '';
    const hValue = hCell.formula || hCell.value || '(空)';

    if (cValue.toString().includes('【E】') || row >= 170) {
      console.log(`Row ${row}: C="${cValue}" H="${hValue}"`);
    }
  }

  console.log('\n========================================');
  console.log('【E】ヘッダーの実際の位置:');
  console.log('========================================\n');

  for (let row = 147; row <= 180; row++) {
    const cCell = sheet.getRow(row).getCell(3);
    if (cCell.value && cCell.value.toString().includes('【E】')) {
      console.log(`【E】ヘッダー: Row ${row}`);
      break;
    }
  }
}

testAllDuplicate().catch(console.error);
