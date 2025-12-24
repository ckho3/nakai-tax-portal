const ExcelJS = require('exceljs');

async function analyzeInterestStructure() {
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
    const incomeSheet = workbook.getWorksheet('【不】①不動産収入');

    console.log('===== 【不】⑤利息シートの構造解析 =====\n');
    
    // Row 10から物件情報テーブルを確認
    console.log('--- 物件情報テーブル（Row 10付近） ---\n');
    
    for (let rowNum = 10; rowNum <= 20; rowNum++) {
      const row = interestSheet.getRow(rowNum);
      const bCell = row.getCell(2); // B列（物件名）
      const gCell = row.getCell(7); // G列
      
      if (bCell.value) {
        const bValue = bCell.value;
        const bDisplay = typeof bValue === 'object' && 'formula' in bValue 
          ? 'formula: ' + bValue.formula 
          : bValue;
        
        const gValue = gCell.value;
        const gDisplay = typeof gValue === 'object' && 'formula' in gValue 
          ? 'formula: ' + gValue.formula 
          : gValue;
        
        console.log('Row ' + rowNum + ':');
        console.log('  B列: ' + bDisplay);
        console.log('  G列: ' + gDisplay);
      }
    }
    
    console.log('\n--- ヘッダー検索（Row 40-50） ---\n');
    
    for (let rowNum = 40; rowNum <= 50; rowNum++) {
      const row = interestSheet.getRow(rowNum);
      const dCell = row.getCell(4);
      
      if (dCell.value && typeof dCell.value === 'string') {
        console.log('Row ' + rowNum + ':', dCell.value);
      }
    }
    
    console.log('\n--- 最初のデータ行（Row 49-52） ---\n');
    
    for (let rowNum = 49; rowNum <= 52; rowNum++) {
      const row = interestSheet.getRow(rowNum);
      const bCell = row.getCell(2);
      const dCell = row.getCell(4);
      
      const bValue = bCell.value;
      const bDisplay = typeof bValue === 'object' && 'formula' in bValue 
        ? 'formula: ' + bValue.formula 
        : bValue;
      
      const dValue = dCell.value;
      const dDisplay = typeof dValue === 'object' && 'formula' in dValue 
        ? 'formula: ' + dValue.formula 
        : dValue;
      
      console.log('Row ' + rowNum + ':');
      console.log('  B列: ' + bDisplay);
      console.log('  D列: ' + dDisplay);
      
      // B列の数式が参照している行を抽出
      if (typeof bValue === 'object' && 'formula' in bValue) {
        const match = bValue.formula.match(/G(\d+)/);
        if (match) {
          const refRow = parseInt(match[1]);
          const incomeRow = incomeSheet.getRow(refRow);
          const propertyName = incomeRow.getCell(7).value;
          console.log('  → 参照: 【不】①不動産収入 G' + refRow + ' = ' + propertyName);
        }
      }
      console.log('');
    }

  } catch (error) {
    console.error('エラー:', error.message);
  }
}

analyzeInterestStructure();
