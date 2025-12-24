const ExcelJS = require('exceljs');

async function checkRowMapping() {
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

    const usefulLifeSheet = workbook.getWorksheet('【不】④耐用年数');
    const incomeSheet = workbook.getWorksheet('【不】①不動産収入');

    console.log('===== 【不】④耐用年数 → 【不】①不動産収入 マッピング =====\n');
    
    const usefulLifeStartRow = 51;
    
    for (let rowNum = usefulLifeStartRow; rowNum <= usefulLifeStartRow + 10; rowNum += 2) {
      const setIndex = Math.floor((rowNum - usefulLifeStartRow) / 2);
      const propertyInfoRowNum = 4 + setIndex;
      
      const usefulLifeRow = usefulLifeSheet.getRow(rowNum);
      const eCell = usefulLifeRow.getCell(5);
      
      if (!eCell.value) continue;
      
      const propertyInfoRow = incomeSheet.getRow(propertyInfoRowNum);
      const propertyNameCell = propertyInfoRow.getCell(7);
      
      console.log('【不】④耐用年数 Row', rowNum, '-', rowNum + 1, '(setIndex=' + setIndex + ')');
      console.log('  → 【不】①不動産収入 Row', propertyInfoRowNum, 'G列:',propertyNameCell.value);
      console.log('');
    }

  } catch (error) {
    console.error('エラー:', error.message);
  }
}

checkRowMapping();
