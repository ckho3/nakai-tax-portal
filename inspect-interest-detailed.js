const ExcelJS = require('exceljs');

async function inspectInterestDetailed() {
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
  console.log('使用ファイル:', files[0].name, '\n');

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    const interestSheet = workbook.getWorksheet('【不】⑤利息');

    console.log('===== 【不】⑤利息シート 詳細 =====\n');
    
    // Row 49の詳細
    console.log('--- Row 49 (D49) ---');
    const row49 = interestSheet.getRow(49);
    const d49 = row49.getCell(4);
    console.log('D49 cellValue:', d49.value);
    console.log('D49 type:', d49.type);
    
    // 参照されているセルも確認
    console.log('\n--- 参照セル G10, C49 ---');
    const g10 = interestSheet.getRow(10).getCell(7);
    const c49 = row49.getCell(3);
    console.log('G10:', g10.value);
    console.log('C49:', c49.value);
    
    // Row 50以降も確認
    console.log('\n--- Row 50-60 (D列) ---');
    for (let rowNum = 50; rowNum <= 60; rowNum++) {
      const row = interestSheet.getRow(rowNum);
      const dCell = row.getCell(4);
      const cellValue = dCell.value;
      
      let display = '';
      if (cellValue && typeof cellValue === 'object' && 'formula' in cellValue) {
        display = 'formula: ' + cellValue.formula + ', result: ' + cellValue.result;
      } else {
        display = 'value: ' + cellValue;
      }
      
      console.log('Row ' + rowNum + ': ' + display);
    }

  } catch (error) {
    console.error('エラー:', error.message);
  }
}

inspectInterestDetailed();
