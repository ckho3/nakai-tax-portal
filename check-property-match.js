const ExcelJS = require('exceljs');

async function checkPropertyMatch() {
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
    const interestSheet = workbook.getWorksheet('【不】⑤利息');
    const incomeSheet = workbook.getWorksheet('【不】①不動産収入');

    // 【不】④耐用年数シートの最初の物件名を取得
    const firstPropertyRow = usefulLifeSheet.getRow(51);
    const eCell = firstPropertyRow.getCell(5);
    console.log('【不】④耐用年数 Row 51 E列:', eCell.value);
    
    // 対応する物件情報テーブルの物件名
    const propertyInfoRow = incomeSheet.getRow(4);
    const propertyName = propertyInfoRow.getCell(7).value;
    console.log('【不】①不動産収入 Row 4 G列（物件名）:', propertyName);
    
    console.log('\n--- 【不】⑤利息シート B列で物件名を検索 ---\n');
    
    const normalize = (text) => {
      if (!text) return '';
      return text.toString().replace(/[\s　・]/g, '').toLowerCase();
    };
    
    const normalizedTarget = normalize(propertyName);
    
    for (let rowNum = 40; rowNum <= 100; rowNum++) {
      const row = interestSheet.getRow(rowNum);
      const bCell = row.getCell(2); // B列（物件名）
      const bValue = bCell.value;
      
      if (bValue) {
        const normalizedValue = normalize(bValue);
        const matches = normalizedValue.includes(normalizedTarget) || normalizedTarget.includes(normalizedValue);
        
        if (matches) {
          console.log('🎯 Row', rowNum, ':', bValue, '✅ マッチ');
          
          // D列の確認
          const dCell = row.getCell(4);
          if (dCell.value && typeof dCell.value === 'object' && 'formula' in dCell.value) {
            console.log('   D列: 数式あり -', dCell.value.formula);
          } else {
            console.log('   D列:', dCell.value);
          }
        }
      }
    }

  } catch (error) {
    console.error('エラー:', error.message);
  }
}

checkPropertyMatch();
