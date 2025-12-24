const ExcelJS = require('exceljs');

async function checkFormulaPattern() {
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

    const interestSheet = workbook.getWorksheet('【不】⑤利息');

    console.log('===== Row 95以降のD列数式パターン =====\n');

    for (let rowNum = 95; rowNum <= 110; rowNum++) {
      const row = interestSheet.getRow(rowNum);
      const cCell = row.getCell(3);
      const dCell = row.getCell(4);

      const dValue = dCell.value;
      const cValue = cCell.value;

      if (dValue && typeof dValue === 'object' && 'formula' in dValue) {
        const formula = dValue.formula;
        
        // IFERROR(INT(G...*C...),0) パターンをチェック
        const pattern1 = /IFERROR\(INT\(G\d+\*C\d+\),0\)/;
        const pattern2 = /IFERROR\(G\d+-D\d+,0\)/;
        
        const isPattern1 = pattern1.test(formula);
        const isPattern2 = pattern2.test(formula);
        
        console.log('Row ' + rowNum + ':');
        console.log('  C列:', cValue);
        console.log('  D列:', formula);
        console.log('  パターン1(INT):', isPattern1 ? '✅' : '❌');
        console.log('  パターン2(引き算):', isPattern2 ? '✅' : '❌');
        console.log('');
      }
    }

  } catch (error) {
    console.error('エラー:', error.message);
  }
}

checkFormulaPattern();
