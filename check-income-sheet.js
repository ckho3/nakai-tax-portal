const ExcelJS = require('exceljs');

async function checkIncomeSheet() {
  const fs = require('fs');
  const path = require('path');

  const outputDir = path.join(__dirname, 'output');
  const files = fs.readdirSync(outputDir)
    .filter(f => f.endsWith('.xlsx'))
    .map(f => ({
      name: f,
      path: path.join(outputDir, f),
      time: fs.statSync(path.join(outputDir, f)).mtime.getTime()
    }))
    .sort((a, b) => b.time - a.time);

  const excelPath = files[0].path;

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    const incomeSheet = workbook.getWorksheet('【不】①不動産収入');
    
    console.log('===== 【不】①不動産収入 Row 4 (物件情報テーブル1件目) =====\n');
    
    const row4 = incomeSheet.getRow(4);
    
    for (let col = 1; col <= 10; col++) {
      const cell = row4.getCell(col);
      const colLetter = String.fromCharCode(64 + col);
      console.log(colLetter + '4:', cell.value);
    }
    
    console.log('\n===== 【不】①不動産収入 Row 55 (収入セクション1件目) =====\n');
    
    const row55 = incomeSheet.getRow(55);
    
    for (let col = 1; col <= 10; col++) {
      const cell = row55.getCell(col);
      const colLetter = String.fromCharCode(64 + col);
      console.log(colLetter + '55:', cell.value);
    }

  } catch (error) {
    console.error('エラー:', error.message);
  }
}

checkIncomeSheet();
