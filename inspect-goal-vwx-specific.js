const ExcelJS = require('exceljs');

async function inspectGoalVWXSpecific() {
  const goalWorkbook = new ExcelJS.Workbook();
  await goalWorkbook.xlsx.readFile('/Users/user/Downloads/goal.xlsx');
  const goalSheet = goalWorkbook.getWorksheet('【不】①不動産収入');

  console.log('========================================');
  console.log('goal.xlsx（23件）のRow 130, 132を確認');
  console.log('========================================\n');

  for (let row = 129; row <= 133; row++) {
    const vCell = goalSheet.getRow(row).getCell(22);
    const wCell = goalSheet.getRow(row).getCell(23);
    const xCell = goalSheet.getRow(row).getCell(24);
    const gCell = goalSheet.getRow(row).getCell(7);

    console.log(`Row ${row}: G="${gCell.value || ''}"`);
    console.log(`  V: ${vCell.formula || vCell.value || '(空)'}`);
    console.log(`  W: ${wCell.formula || wCell.value || '(空)'}`);
    console.log(`  X: ${xCell.formula || xCell.value || '(空)'}`);
    console.log('');
  }

  console.log('========================================');
  console.log('goal.xlsx（23件）のRow 158-182を確認');
  console.log('========================================\n');

  for (let row = 158; row <= 182; row++) {
    const vCell = goalSheet.getRow(row).getCell(22);
    const wCell = goalSheet.getRow(row).getCell(23);
    const xCell = goalSheet.getRow(row).getCell(24);
    const gCell = goalSheet.getRow(row).getCell(7);
    const cCell = goalSheet.getRow(row).getCell(3);

    console.log(`Row ${row}: C="${cCell.value || ''}" G="${gCell.value || ''}"`);
    console.log(`  V: ${vCell.formula || vCell.value || '(空)'}`);
    console.log(`  W: ${wCell.formula || wCell.value || '(空)'}`);
    console.log(`  X: ${xCell.formula || xCell.value || '(空)'}`);
    console.log('');
  }

  console.log('========================================');
  console.log('原本の対応する行を確認');
  console.log('========================================\n');

  const originalWorkbook = new ExcelJS.Workbook();
  await originalWorkbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');
  const originalSheet = originalWorkbook.getWorksheet('【不】①不動産収入');

  console.log('原本のRow 123（【C】セクションの合計行）:');
  const v123 = originalSheet.getRow(123).getCell(22);
  const w123 = originalSheet.getRow(123).getCell(23);
  const x123 = originalSheet.getRow(123).getCell(24);
  console.log(`  V123: ${v123.formula || v123.value || '(空)'}`);
  console.log(`  W123: ${w123.formula || w123.value || '(空)'}`);
  console.log(`  X123: ${x123.formula || x123.value || '(空)'}`);
  console.log('');

  console.log('原本のRow 146（【D】セクションの後）:');
  for (let row = 146; row <= 170; row++) {
    const vCell = originalSheet.getRow(row).getCell(22);
    const cCell = originalSheet.getRow(row).getCell(3);
    const gCell = originalSheet.getRow(row).getCell(7);

    if (vCell.formula || vCell.value) {
      console.log(`Row ${row}: C="${cCell.value || ''}" G="${gCell.value || ''}"`);
      console.log(`  V: ${vCell.formula || vCell.value}`);
    }
  }
  console.log('');

  console.log('========================================');
  console.log('【E】セクションの確認');
  console.log('========================================\n');

  console.log('原本の【E】セクション（Row 147-166）:');
  for (let row = 147; row <= 150; row++) {
    const vCell = originalSheet.getRow(row).getCell(22);
    const wCell = originalSheet.getRow(row).getCell(23);
    const xCell = originalSheet.getRow(row).getCell(24);
    const cCell = originalSheet.getRow(row).getCell(3);

    console.log(`Row ${row}: C="${cCell.value || ''}"`);
    console.log(`  V: ${vCell.formula || vCell.value || '(空)'}`);
    console.log(`  W: ${wCell.formula || wCell.value || '(空)'}`);
    console.log(`  X: ${xCell.formula || xCell.value || '(空)'}`);
    console.log('');
  }

  console.log('goal.xlsxの【E】セクション（Row 159-182）:');
  for (let row = 159; row <= 165; row++) {
    const vCell = goalSheet.getRow(row).getCell(22);
    const wCell = goalSheet.getRow(row).getCell(23);
    const xCell = goalSheet.getRow(row).getCell(24);
    const cCell = goalSheet.getRow(row).getCell(3);

    console.log(`Row ${row}: C="${cCell.value || ''}"`);
    console.log(`  V: ${vCell.formula || vCell.value || '(空)'}`);
    console.log(`  W: ${wCell.formula || wCell.value || '(空)'}`);
    console.log(`  X: ${xCell.formula || xCell.value || '(空)'}`);
    console.log('');
  }
}

inspectGoalVWXSpecific().catch(console.error);
