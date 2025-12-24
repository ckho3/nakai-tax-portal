const ExcelJS = require('exceljs');

async function inspectGoalDetailed() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Downloads/goal.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  if (!sheet) {
    console.log('シート【不】①不動産収入が見つかりません');
    return;
  }

  console.log('========================================');
  console.log('goal.xlsxの詳細分析（行追加の確認）');
  console.log('========================================\n');

  // 原本と比較するために、原本も読み込む
  const originalWorkbook = new ExcelJS.Workbook();
  await originalWorkbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');
  const originalSheet = originalWorkbook.getWorksheet('【不】①不動産収入');

  console.log('【A】収入セクションの行数比較:');
  console.log('原本:');
  for (let i = 55; i <= 78; i++) {
    const cValue = originalSheet.getRow(i).getCell(3).value;
    const gValue = originalSheet.getRow(i).getCell(7).value;
    if (cValue || gValue || i === 55 || i === 78) {
      console.log(`  Row ${i}: C="${cValue}" G="${gValue}"`);
    }
  }
  console.log('\ngoal.xlsx:');
  for (let i = 55; i <= 81; i++) {
    const cValue = sheet.getRow(i).getCell(3).value;
    const gValue = sheet.getRow(i).getCell(7).value;
    if (cValue || gValue || i === 55 || i === 78) {
      console.log(`  Row ${i}: C="${cValue}" G="${gValue}"`);
    }
  }
  console.log('');

  console.log('【B】管理手数料セクションの行数比較:');
  console.log('原本:');
  for (let i = 78; i <= 105; i++) {
    const cValue = originalSheet.getRow(i).getCell(3).value;
    if (cValue && cValue.toString().includes('【B】')) {
      console.log(`  Row ${i}: C="${cValue}"`);
      break;
    }
  }
  console.log('\ngoal.xlsx:');
  for (let i = 78; i <= 110; i++) {
    const cValue = sheet.getRow(i).getCell(3).value;
    if (cValue && (cValue.toString().includes('【B】') || cValue.toString().includes('【C】'))) {
      console.log(`  Row ${i}: C="${cValue}"`);
    }
  }
  console.log('');

  console.log('========================================');
  console.log('行数のカウント');
  console.log('========================================\n');

  // 【A】セクションのデータ行をカウント
  let aCount = 0;
  for (let i = 55; i <= 100; i++) {
    const gValue = sheet.getRow(i).getCell(7).value;
    if (gValue && gValue === '収入合計①') {
      aCount++;
    }
  }
  console.log(`【A】収入のデータ行数: ${aCount}行`);

  // 【B】セクションのデータ行をカウント
  let bCount = 0;
  for (let i = 78; i <= 120; i++) {
    const gValue = sheet.getRow(i).getCell(7).value;
    if (gValue && gValue === '管理手数料') {
      bCount++;
    }
  }
  console.log(`【B】管理手数料のデータ行数: ${bCount}行`);

  // 【C】セクションのヘッダー位置を確認
  for (let i = 100; i <= 120; i++) {
    const cValue = sheet.getRow(i).getCell(3).value;
    if (cValue && cValue.toString().includes('【C】')) {
      console.log(`【C】広告費のヘッダー位置: Row ${i}`);
      break;
    }
  }

  // 【D】セクションのヘッダー位置を確認
  for (let i = 130; i <= 140; i++) {
    const cValue = sheet.getRow(i).getCell(3).value;
    if (cValue && cValue.toString().includes('【D】')) {
      console.log(`【D】修繕費のヘッダー位置: Row ${i}`);
      break;
    }
  }

  console.log('');
  console.log('========================================');
  console.log('結論');
  console.log('========================================\n');
  console.log('原本は各セクション20行ですが、');
  console.log(`goal.xlsxは${aCount}行のデータがあります。`);
  console.log(`つまり、${aCount - 20}行が追加されています。`);
}

inspectGoalDetailed().catch(console.error);
