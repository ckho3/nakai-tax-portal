const ExcelJS = require('exceljs');

async function inspectExcel() {
  const excelPath = '/Users/user/Downloads/1766418694555-647937845-【原本】R7確定申告フォーマット（不動産所得・太陽光事業　共用）_updated_2025-12-22T15-51-35_newprop_2025-12-22T15-51-35_depreciation_2025-12-22T15-51-36.xlsx';

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    // 【不】③減価償却（JDLよりエクスポート）シートを確認
    const depreciationSheet = workbook.getWorksheet('【不】③減価償却（JDLよりエクスポート）');

    console.log('\n===== 【不】③減価償却（JDLよりエクスポート）シート E4-I13 =====');
    for (let rowNum = 4; rowNum <= 13; rowNum++) {
      const row = depreciationSheet.getRow(rowNum);
      const eValue = row.getCell(5).value; // E列
      const gValue = row.getCell(7).value; // G列
      const iValue = row.getCell(9).value; // I列

      console.log(`Row ${rowNum}: E="${eValue}", G="${gValue}", I="${iValue}"`);
    }

    // 【不】④耐用年数シート Row 51-60 のC, G, L列を確認
    const usefulLifeSheet = workbook.getWorksheet('【不】④耐用年数');

    console.log('\n===== 【不】④耐用年数シート Row 51-60 (C, G, L列) =====');
    for (let rowNum = 51; rowNum <= 60; rowNum++) {
      const row = usefulLifeSheet.getRow(rowNum);
      const cCell = row.getCell(3); // C列
      const gCell = row.getCell(7); // G列
      const lCell = row.getCell(12); // L列

      console.log(`Row ${rowNum}:`);
      console.log(`  C列 type=${cCell.type}, value=`, cCell.value);
      console.log(`  G列 type=${gCell.type}, value=`, gCell.value);
      console.log(`  L列 type=${lCell.type}, value=`, lCell.value);
    }

    // 【不】①不動産収入シート Row 55-64 のH列を確認
    const incomeSheet = workbook.getWorksheet('【不】①不動産収入');

    console.log('\n===== 【不】①不動産収入シート Row 55-64 (H列) =====');
    for (let rowNum = 55; rowNum <= 64; rowNum++) {
      const row = incomeSheet.getRow(rowNum);
      const hCell = row.getCell(8); // H列

      console.log(`Row ${rowNum}: H列 type=${hCell.type}, value=`, hCell.value);
    }

  } catch (error) {
    console.error('エラー:', error.message);
  }
}

inspectExcel();
