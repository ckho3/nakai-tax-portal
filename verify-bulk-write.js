const ExcelJS = require('exceljs');

async function verifyBulkWrite() {
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
  console.log('使用ファイル:', files[0].name.substring(0, 100), '...\n');

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    const interestSheet = workbook.getWorksheet('【不】⑤利息');

    console.log('===== 【不】⑤利息シート Row 95-104の確認 =====\n');

    let successCount = 0;
    let failCount = 0;

    for (let rowNum = 95; rowNum <= 104; rowNum++) {
      const row = interestSheet.getRow(rowNum);
      const cCell = row.getCell(3);
      const dCell = row.getCell(4);

      const cValue = cCell.value;
      const dValue = dCell.value;

      let dFormula = '';
      if (dValue && typeof dValue === 'object' && 'formula' in dValue) {
        dFormula = dValue.formula;
      }

      const pattern = /IFERROR\(INT\(G\d+\*C\d+\),0\)/;
      const isPattern = pattern.test(dFormula);

      if (isPattern) {
        const isCorrect = cValue === 0.8;
        console.log('Row ' + rowNum + ':');
        console.log('  C列:', cValue, isCorrect ? '✅' : '❌');
        console.log('  D列:', dFormula);
        console.log('');

        if (isCorrect) {
          successCount++;
        } else {
          failCount++;
        }
      }
    }

    console.log('結果:');
    console.log('  ✅ 正しく書き込まれた行:', successCount);
    console.log('  ❌ 失敗した行:', failCount);

  } catch (error) {
    console.error('エラー:', error.message);
  }
}

verifyBulkWrite();
