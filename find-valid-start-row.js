const ExcelJS = require('exceljs');

async function findValidStartRow() {
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

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    const interestSheet = workbook.getWorksheet('【不】⑤利息');

    console.log('===== 【不】⑤利息シート 有効なデータ行の検索 =====\n');
    
    // ヘッダー行を探す
    let headerRowNum = null;
    for (let rowNum = 10; rowNum <= 100; rowNum++) {
      const row = interestSheet.getRow(rowNum);
      const dCell = row.getCell(4);
      const value = dCell.value;
      
      if (value && typeof value === 'string' && value.includes('『減価償却』の「取得費」に転記')) {
        headerRowNum = rowNum;
        console.log('ヘッダー発見: Row', rowNum);
        break;
      }
    }
    
    if (!headerRowNum) {
      console.log('ヘッダーが見つかりませんでした');
      return;
    }
    
    console.log('\n--- データ行の検索（Row ' + (headerRowNum + 1) + '以降） ---\n');
    
    for (let rowNum = headerRowNum + 1; rowNum <= headerRowNum + 50; rowNum++) {
      const row = interestSheet.getRow(rowNum);
      const dCell = row.getCell(4); // D列（数式がある列）
      const bCell = row.getCell(2); // B列（物件名）
      const cCell = row.getCell(3); // C列（建物比率）
      const gCell = row.getCell(7); // G列（取得価額合計）
      
      const hasFormula = dCell.value && typeof dCell.value === 'object' && 'formula' in dCell.value;
      const hasPropertyName = bCell.value !== null && bCell.value !== undefined;
      const hasBuildingRatio = cCell.value !== null && cCell.value !== undefined;
      const hasTotalPrice = gCell.value !== null && gCell.value !== undefined;
      
      if (hasFormula) {
        const marker = (hasPropertyName && hasBuildingRatio && hasTotalPrice) ? ' ✅ 有効' : ' ❌ データ不足';
        console.log('Row ' + rowNum + ':');
        console.log('  D列: ' + (dCell.value.formula || dCell.value.sharedFormula));
        console.log('  B列(物件名): ' + (hasPropertyName ? '○' : '×'));
        console.log('  C列(建物比率): ' + (hasBuildingRatio ? '○ (' + cCell.value + ')' : '×'));
        console.log('  G列(取得価額): ' + (hasTotalPrice ? '○' : '×'));
        console.log('  判定:' + marker + '\n');
        
        if (hasPropertyName && hasBuildingRatio && hasTotalPrice) {
          console.log('🎯 最初の有効なデータ行: Row ' + rowNum);
          break;
        }
      }
    }

  } catch (error) {
    console.error('エラー:', error.message);
  }
}

findValidStartRow();
