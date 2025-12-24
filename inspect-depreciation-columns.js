const ExcelJS = require('exceljs');

async function inspectDepreciationSheet() {
  const excelPath = '/Users/user/Downloads/1766418694555-647937845-【原本】R7確定申告フォーマット（不動産所得・太陽光事業　共用）_updated_2025-12-22T15-51-35_newprop_2025-12-22T15-51-35_depreciation_2025-12-22T15-51-36.xlsx';

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    const depreciationSheet = workbook.getWorksheet('【不】③減価償却（JDLよりエクスポート）');

    console.log('\n===== 【不】③減価償却（JDLよりエクスポート）シート Row 4の全列 =====');
    const row4 = depreciationSheet.getRow(4);

    // A-M列まで確認
    for (let col = 1; col <= 13; col++) {
      const cell = row4.getCell(col);
      const colLetter = String.fromCharCode(64 + col); // A=65
      console.log(`${colLetter}4: type=${cell.type}, value=`, cell.value);
    }

  } catch (error) {
    console.error('エラー:', error.message);
  }
}

inspectDepreciationSheet();
