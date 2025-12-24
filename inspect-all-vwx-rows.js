const ExcelJS = require('exceljs');

async function inspectAllVWXRows() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('========================================');
  console.log('【A】セクション（Row 55-76）のV-X列を全て確認');
  console.log('========================================\n');

  for (let row = 55; row <= 76; row++) {
    const vCell = sheet.getRow(row).getCell(22);
    const wCell = sheet.getRow(row).getCell(23);
    const xCell = sheet.getRow(row).getCell(24);
    const gCell = sheet.getRow(row).getCell(7);

    const vValue = vCell.formula || vCell.value || '(空)';
    const wValue = wCell.formula || wCell.value || '(空)';
    const xValue = xCell.formula || xCell.value || '(空)';

    console.log(`Row ${row}: G="${gCell.value || ''}"`);
    console.log(`  V: ${vValue}`);
    console.log(`  W: ${wValue}`);
    console.log(`  X: ${xValue}`);
    console.log('');
  }

  console.log('========================================');
  console.log('【B】セクション（Row 78-100）のV-X列を全て確認');
  console.log('========================================\n');

  for (let row = 78; row <= 100; row++) {
    const vCell = sheet.getRow(row).getCell(22);
    const wCell = sheet.getRow(row).getCell(23);
    const xCell = sheet.getRow(row).getCell(24);
    const gCell = sheet.getRow(row).getCell(7);

    const vValue = vCell.formula || vCell.value || '(空)';
    const wValue = wCell.formula || wCell.value || '(空)';
    const xValue = xCell.formula || xCell.value || '(空)';

    console.log(`Row ${row}: G="${gCell.value || ''}"`);
    console.log(`  V: ${vValue}`);
    console.log(`  W: ${wValue}`);
    console.log(`  X: ${xValue}`);
    console.log('');
  }

  console.log('========================================');
  console.log('【C】セクション（Row 101-123）のV-X列を全て確認');
  console.log('========================================\n');

  for (let row = 101; row <= 123; row++) {
    const vCell = sheet.getRow(row).getCell(22);
    const wCell = sheet.getRow(row).getCell(23);
    const xCell = sheet.getRow(row).getCell(24);
    const gCell = sheet.getRow(row).getCell(7);

    const vValue = vCell.formula || vCell.value || '(空)';
    const wValue = wCell.formula || wCell.value || '(空)';
    const xValue = xCell.formula || xCell.value || '(空)';

    console.log(`Row ${row}: G="${gCell.value || ''}"`);
    console.log(`  V: ${vValue}`);
    console.log(`  W: ${wValue}`);
    console.log(`  X: ${xValue}`);
    console.log('');
  }

  console.log('========================================');
  console.log('【D】セクション（Row 124-146）のV-X列を全て確認');
  console.log('========================================\n');

  for (let row = 124; row <= 146; row++) {
    const vCell = sheet.getRow(row).getCell(22);
    const wCell = sheet.getRow(row).getCell(23);
    const xCell = sheet.getRow(row).getCell(24);
    const gCell = sheet.getRow(row).getCell(7);

    const vValue = vCell.formula || vCell.value || '(空)';
    const wValue = wCell.formula || wCell.value || '(空)';
    const xValue = xCell.formula || xCell.value || '(空)';

    console.log(`Row ${row}: G="${gCell.value || ''}"`);
    console.log(`  V: ${vValue}`);
    console.log(`  W: ${wValue}`);
    console.log(`  X: ${xValue}`);
    console.log('');
  }

  console.log('========================================');
  console.log('まとめ: 数式がある行とない行');
  console.log('========================================\n');

  console.log('【A】セクション:');
  for (let row = 55; row <= 76; row++) {
    const vCell = sheet.getRow(row).getCell(22);
    const hasFormula = vCell.formula ? '✅' : '❌';
    const gValue = sheet.getRow(row).getCell(7).value;
    console.log(`  Row ${row} ${hasFormula}: ${gValue || '(空)'}`);
  }
  console.log('');

  console.log('【B】セクション:');
  for (let row = 78; row <= 100; row++) {
    const vCell = sheet.getRow(row).getCell(22);
    const hasFormula = vCell.formula ? '✅' : '❌';
    const gValue = sheet.getRow(row).getCell(7).value;
    console.log(`  Row ${row} ${hasFormula}: ${gValue || '(空)'}`);
  }
  console.log('');

  console.log('【C】セクション:');
  for (let row = 101; row <= 123; row++) {
    const vCell = sheet.getRow(row).getCell(22);
    const hasFormula = vCell.formula ? '✅' : '❌';
    const gValue = sheet.getRow(row).getCell(7).value;
    console.log(`  Row ${row} ${hasFormula}: ${gValue || '(空)'}`);
  }
  console.log('');

  console.log('【D】セクション:');
  for (let row = 124; row <= 146; row++) {
    const vCell = sheet.getRow(row).getCell(22);
    const hasFormula = vCell.formula ? '✅' : '❌';
    const gValue = sheet.getRow(row).getCell(7).value;
    console.log(`  Row ${row} ${hasFormula}: ${gValue || '(空)'}`);
  }
}

inspectAllVWXRows().catch(console.error);
