const ExcelJS = require('exceljs');
const path = require('path');

async function debugStructure() {
  try {
    const workbook = new ExcelJS.Workbook();
    const excelPath = path.join(__dirname, '【原本】R7確定申告フォーマット.xlsx');
    await workbook.xlsx.readFile(excelPath);

    const sheet = workbook.getWorksheet('【不】①不動産収入');

    console.log('========================================');
    console.log('【A】収入セクションの詳細確認（Row 55-60）');
    console.log('========================================\n');

    for (let rowNum = 55; rowNum <= 60; rowNum++) {
      const row = sheet.getRow(rowNum);
      const cellC = row.getCell(3).value;
      const cellG = row.getCell(7).value;
      const cellH = row.getCell(8).value;

      const formatCell = (val) => {
        if (!val) return '(空)';
        if (typeof val === 'object' && val.formula) return '[Formula]';
        return val.toString().substring(0, 40);
      };

      console.log(`Row ${rowNum}:`);
      console.log(`  C="${formatCell(cellC)}"`);
      console.log(`  G="${formatCell(cellG)}"`);
      console.log(`  H="${formatCell(cellH)}"`);
      console.log('');
    }

    console.log('========================================');
    console.log('【B】管理手数料セクションの詳細確認（Row 78-83）');
    console.log('========================================\n');

    for (let rowNum = 78; rowNum <= 83; rowNum++) {
      const row = sheet.getRow(rowNum);
      const cellC = row.getCell(3).value;
      const cellG = row.getCell(7).value;
      const cellH = row.getCell(8).value;

      const formatCell = (val) => {
        if (!val) return '(空)';
        if (typeof val === 'object' && val.formula) return '[Formula]';
        return val.toString().substring(0, 40);
      };

      console.log(`Row ${rowNum}:`);
      console.log(`  C="${formatCell(cellC)}"`);
      console.log(`  G="${formatCell(cellG)}"`);
      console.log(`  H="${formatCell(cellH)}"`);
      console.log('');
    }

    console.log('========================================');
    console.log('結論：データ開始行はどこか？');
    console.log('========================================\n');

    console.log('【A】収入セクション:');
    console.log('  Row 55: C列にヘッダー「【A】収入...」');
    console.log('  Row 56〜: データ行の開始？');
    console.log('');

    console.log('【B】管理手数料セクション:');
    console.log('  Row 78: C列にヘッダー「【B】支払手数料...」');
    console.log('  Row 79〜: データ行の開始？');

  } catch (error) {
    console.error('エラー:', error);
  }
}

debugStructure();
