const ExcelJS = require('exceljs');

async function inspectGoal() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Downloads/goal.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  if (!sheet) {
    console.log('シート【不】①不動産収入が見つかりません');
    return;
  }

  console.log('========================================');
  console.log('goal.xlsxの構造分析');
  console.log('========================================\n');

  // 【A】セクションを確認
  console.log('【A】収入セクション:');
  console.log('Row 55:', sheet.getRow(55).getCell(3).value); // C列のヘッダー
  for (let i = 55; i <= 80; i++) {
    const gValue = sheet.getRow(i).getCell(7).value;
    const hValue = sheet.getRow(i).getCell(8).value;
    if (gValue || hValue) {
      console.log(`  Row ${i}: G="${gValue}" H="${hValue}"`);
    }
  }
  console.log('');

  // 【B】セクションを確認
  console.log('【B】管理手数料セクション:');
  for (let i = 78; i <= 110; i++) {
    const cValue = sheet.getRow(i).getCell(3).value;
    const gValue = sheet.getRow(i).getCell(7).value;
    const hValue = sheet.getRow(i).getCell(8).value;
    if (cValue || gValue || hValue) {
      console.log(`  Row ${i}: C="${cValue}" G="${gValue}" H="${hValue}"`);
    }
  }
  console.log('');

  // 【C】セクションを確認
  console.log('【C】広告費セクション:');
  for (let i = 104; i <= 135; i++) {
    const cValue = sheet.getRow(i).getCell(3).value;
    const gValue = sheet.getRow(i).getCell(7).value;
    const hValue = sheet.getRow(i).getCell(8).value;
    if (cValue || gValue || hValue) {
      console.log(`  Row ${i}: C="${cValue}" G="${gValue}" H="${hValue}"`);
    }
  }
  console.log('');

  // 【D】セクションを確認
  console.log('【D】修繕費セクション:');
  for (let i = 130; i <= 165; i++) {
    const cValue = sheet.getRow(i).getCell(3).value;
    const gValue = sheet.getRow(i).getCell(7).value;
    const hValue = sheet.getRow(i).getCell(8).value;
    if (cValue || gValue || hValue) {
      console.log(`  Row ${i}: C="${cValue}" G="${gValue}" H="${hValue}"`);
    }
  }
  console.log('');
}

inspectGoal().catch(console.error);
