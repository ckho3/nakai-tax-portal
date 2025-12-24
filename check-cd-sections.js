const ExcelJS = require('exceljs');
const path = require('path');

async function checkSections() {
  try {
    const workbook = new ExcelJS.Workbook();
    const excelPath = path.join(__dirname, '【原本】R7確定申告フォーマット.xlsx');
    await workbook.xlsx.readFile(excelPath);

    const sheet = workbook.getWorksheet('【不】①不動産収入');

    const formatCell = (val) => {
      if (!val) return '(空)';
      if (typeof val === 'object' && val.formula) return '[Formula]';
      return val.toString().substring(0, 50);
    };

    console.log('========================================');
    console.log('【C】広告費セクション（Row 101-106）');
    console.log('========================================\n');

    for (let rowNum = 101; rowNum <= 106; rowNum++) {
      const row = sheet.getRow(rowNum);
      const cellC = row.getCell(3).value;
      const cellG = row.getCell(7).value;
      const cellH = row.getCell(8).value;

      console.log(`Row ${rowNum}:`);
      console.log(`  C="${formatCell(cellC)}"`);
      console.log(`  G="${formatCell(cellG)}"`);
      console.log(`  H="${formatCell(cellH)}"`);
      console.log('');
    }

    console.log('========================================');
    console.log('【D】修繕費セクション（Row 124-129）');
    console.log('========================================\n');

    for (let rowNum = 124; rowNum <= 129; rowNum++) {
      const row = sheet.getRow(rowNum);
      const cellC = row.getCell(3).value;
      const cellG = row.getCell(7).value;
      const cellH = row.getCell(8).value;

      console.log(`Row ${rowNum}:`);
      console.log(`  C="${formatCell(cellC)}"`);
      console.log(`  G="${formatCell(cellG)}"`);
      console.log(`  H="${formatCell(cellH)}"`);
      console.log('');
    }

    console.log('========================================');
    console.log('結論：各セクションのデータ開始行');
    console.log('========================================\n');

    console.log('【A】収入:');
    console.log('  Row 55: ヘッダー');
    console.log('  Row 56: データ開始');
    console.log('');

    console.log('【B】管理手数料:');
    console.log('  Row 78: ヘッダー');
    console.log('  Row 79: 注釈（→サブリースの場合...）');
    console.log('  Row 80: データ開始');
    console.log('');

    console.log('【C】広告費:');
    console.log('  Row 101: ヘッダー');
    console.log('  Row 102 or 103?: データ開始（要確認）');
    console.log('');

    console.log('【D】修繕費:');
    console.log('  Row 124: ヘッダー');
    console.log('  Row 125 or 126?: データ開始（要確認）');

  } catch (error) {
    console.error('エラー:', error);
  }
}

checkSections();
