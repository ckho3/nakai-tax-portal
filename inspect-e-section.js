const ExcelJS = require('exceljs');

async function inspectESection() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('========================================');
  console.log('【E】セクション（Row 147-170）の全列を確認');
  console.log('========================================\n');

  for (let row = 147; row <= 170; row++) {
    const cCell = sheet.getRow(row).getCell(3);  // C列
    const gCell = sheet.getRow(row).getCell(7);  // G列
    const hCell = sheet.getRow(row).getCell(8);  // H列
    const iCell = sheet.getRow(row).getCell(9);  // I列
    const tCell = sheet.getRow(row).getCell(20); // T列
    const uCell = sheet.getRow(row).getCell(21); // U列
    const vCell = sheet.getRow(row).getCell(22); // V列

    console.log(`Row ${row}:`);
    console.log(`  C: "${cCell.value || ''}"`);
    console.log(`  G: "${gCell.value || ''}"`);
    console.log(`  H: ${hCell.formula || hCell.value || '(空)'}`);
    console.log(`  I: ${iCell.formula || iCell.value || '(空)'}`);
    console.log(`  T: ${tCell.formula || tCell.value || '(空)'}`);
    console.log(`  U: ${uCell.formula || uCell.value || '(空)'}`);
    console.log(`  V: ${vCell.formula || vCell.value || '(空)'}`);
    console.log('');
  }

  console.log('========================================');
  console.log('【E】セクションの理解');
  console.log('========================================\n');
  console.log('【E】セクションは計算専用セクション？');
  console.log('- H列: 【D】セクションへの参照');
  console.log('- I-T列: サブリース計算式');
  console.log('- U列: 合計');
  console.log('- V-X列: SUMIF集計');
  console.log('');
  console.log('→ PDFからデータを直接書き込む必要はない？');
}

inspectESection().catch(console.error);
