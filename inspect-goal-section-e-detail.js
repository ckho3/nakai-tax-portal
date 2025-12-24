const ExcelJS = require('exceljs');

async function inspectSectionEDetail() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Downloads/goal.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('========================================');
  console.log('【E】セクション Row 159-182 を詳細確認');
  console.log('========================================\n');

  for (let i = 159; i <= 182; i++) {
    const cValue = sheet.getRow(i).getCell(3).value;
    const hValue = sheet.getRow(i).getCell(8).value;
    const hFormula = sheet.getRow(i).getCell(8).formula;

    console.log(`Row ${i}:`);
    console.log(`  C列: "${cValue}"`);
    if (hFormula) {
      console.log(`  H列: 数式=${hFormula}`);
    } else if (hValue) {
      console.log(`  H列: 値=${hValue}`);
    } else {
      console.log(`  H列: (空)`);
    }
    console.log('');
  }
}

inspectSectionEDetail().catch(console.error);
