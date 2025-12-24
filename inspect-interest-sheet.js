const ExcelJS = require('exceljs');

async function inspectInterestSheet() {
  // 最新のExcelファイルを使用
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

  if (files.length === 0) {
    console.log('出力ファイルが見つかりません');
    return;
  }

  const excelPath = files[0].path;
  console.log(`使用ファイル: ${files[0].name}\n`);

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    const interestSheet = workbook.getWorksheet('【不】⑤利息');

    if (!interestSheet) {
      console.log('【不】⑤利息シートが見つかりません');
      return;
    }

    console.log('===== 【不】⑤利息シート D列 Row 40-70 =====\n');

    for (let rowNum = 40; rowNum <= 70; rowNum++) {
      const row = interestSheet.getRow(rowNum);
      const dCell = row.getCell(4); // D列
      const cellValue = dCell.value;

      let displayValue = '';
      let actualValue = null;

      if (cellValue === null || cellValue === undefined) {
        displayValue = '(null)';
      } else if (typeof cellValue === 'object') {
        if ('formula' in cellValue || 'sharedFormula' in cellValue) {
          const formula = cellValue.formula || cellValue.sharedFormula;
          const result = cellValue.result;
          displayValue = `数式: ${formula}, result: ${result}`;
          actualValue = result;
        } else {
          displayValue = `オブジェクト: ${Object.keys(cellValue).join(', ')}`;
        }
      } else {
        displayValue = `値: ${cellValue}`;
        actualValue = cellValue;
      }

      const isValid = typeof actualValue === 'number' && actualValue >= 0;
      const marker = isValid ? ' ✅' : '';

      console.log(`Row ${rowNum}: ${displayValue}${marker}`);
    }

  } catch (error) {
    console.error('エラー:', error.message);
  }
}

inspectInterestSheet();
