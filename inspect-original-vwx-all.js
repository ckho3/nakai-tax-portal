const ExcelJS = require('exceljs');

async function inspectOriginalVWXAll() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('========================================');
  console.log('原本のRow 130, 132周辺を確認');
  console.log('========================================\n');

  console.log('【C】セクション終わり（Row 120-125）:');
  for (let row = 120; row <= 125; row++) {
    const vCell = sheet.getRow(row).getCell(22);
    const wCell = sheet.getRow(row).getCell(23);
    const xCell = sheet.getRow(row).getCell(24);
    const gCell = sheet.getRow(row).getCell(7);
    const cCell = sheet.getRow(row).getCell(3);

    console.log(`Row ${row}: C="${cCell.value || ''}" G="${gCell.value || ''}"`);
    console.log(`  V: ${vCell.formula || vCell.value || '(空)'}`);
    console.log(`  W: ${wCell.formula || wCell.value || '(空)'}`);
    console.log(`  X: ${xCell.formula || xCell.value || '(空)'}`);
    console.log('');
  }

  console.log('========================================');
  console.log('【D】セクション終わり（Row 143-148）:');
  console.log('========================================\n');

  for (let row = 143; row <= 148; row++) {
    const vCell = sheet.getRow(row).getCell(22);
    const wCell = sheet.getRow(row).getCell(23);
    const xCell = sheet.getRow(row).getCell(24);
    const gCell = sheet.getRow(row).getCell(7);
    const cCell = sheet.getRow(row).getCell(3);

    console.log(`Row ${row}: C="${cCell.value || ''}" G="${gCell.value || ''}"`);
    console.log(`  V: ${vCell.formula || vCell.value || '(空)'}`);
    console.log(`  W: ${wCell.formula || wCell.value || '(空)'}`);
    console.log(`  X: ${xCell.formula || xCell.value || '(空)'}`);
    console.log('');
  }

  console.log('========================================');
  console.log('【E】セクション全体（Row 147-170）:');
  console.log('========================================\n');

  for (let row = 147; row <= 170; row++) {
    const vCell = sheet.getRow(row).getCell(22);
    const wCell = sheet.getRow(row).getCell(23);
    const xCell = sheet.getRow(row).getCell(24);
    const gCell = sheet.getRow(row).getCell(7);
    const cCell = sheet.getRow(row).getCell(3);

    console.log(`Row ${row}: C="${cCell.value || ''}" G="${gCell.value || ''}"`);
    console.log(`  V: ${vCell.formula || vCell.value || '(空)'}`);
    console.log(`  W: ${wCell.formula || wCell.value || '(空)'}`);
    console.log(`  X: ${xCell.formula || xCell.value || '(空)'}`);
    console.log('');
  }

  console.log('========================================');
  console.log('まとめ: V-X列に数式がある行');
  console.log('========================================\n');

  console.log('【E】セクション:');
  for (let row = 147; row <= 170; row++) {
    const vCell = sheet.getRow(row).getCell(22);
    const hasFormula = vCell.formula ? '✅' : '❌';
    const cValue = sheet.getRow(row).getCell(3).value;
    console.log(`  Row ${row} ${hasFormula}: ${cValue || '(空)'}`);
  }
}

inspectOriginalVWXAll().catch(console.error);
