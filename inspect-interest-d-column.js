const ExcelJS = require('exceljs');

async function inspectInterestSheet() {
  const excelPath = '/Users/user/Downloads/1766418694555-647937845-【原本】R7確定申告フォーマット（不動産所得・太陽光事業　共用）_updated_2025-12-22T15-51-35_newprop_2025-12-22T15-51-35_depreciation_2025-12-22T15-51-36.xlsx';

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    const interestSheet = workbook.getWorksheet('【不】⑤利息');

    console.log('\n===== 【不】⑤利息シート D46-D60 =====');
    for (let rowNum = 46; rowNum <= 60; rowNum++) {
      const row = interestSheet.getRow(rowNum);
      const dCell = row.getCell(4); // D列

      console.log(`\nRow ${rowNum}:`);
      console.log(`  type: ${dCell.type}`);
      console.log(`  value:`, dCell.value);

      // 数式オブジェクトの場合
      if (dCell.value && typeof dCell.value === 'object') {
        console.log(`  formula: ${dCell.value.formula || dCell.value.sharedFormula || '(なし)'}`);
        console.log(`  result: ${dCell.value.result !== undefined ? dCell.value.result : '(なし)'}`);
      }
    }

  } catch (error) {
    console.error('エラー:', error.message);
  }
}

inspectInterestSheet();
