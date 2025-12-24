const ExcelJS = require('exceljs');

async function testNewLogic() {
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
  console.log('使用ファイル:', files[0].name, '\n');

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    const interestSheet = workbook.getWorksheet('【不】⑤利息');

    console.log('===== 新しいロジックのテスト =====\n');
    console.log('条件: C列とD列に数値があり、G列が使われている行を探す\n');

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

    console.log('\n--- ヘッダー行以降のデータをチェック ---\n');

    // Row 10-43の物件情報テーブルをチェック
    for (let rowNum = 10; rowNum <= 43; rowNum++) {
      const row = interestSheet.getRow(rowNum);
      const cCell = row.getCell(3); // C列（土地）
      const dCell = row.getCell(4); // D列（建物）
      const gCell = row.getCell(7); // G列（取得価額合計）

      const cValue = cCell.value;
      const dValue = dCell.value;
      const gValue = gCell.value;

      // C列またはD列に数値がある
      const hasCValue = typeof cValue === 'number' && cValue > 0;
      const hasDValue = typeof dValue === 'number' && dValue > 0;
      
      // G列が使われている（数式または数値）
      const hasGValue = gValue !== null && gValue !== undefined && 
        (typeof gValue === 'number' || 
         (typeof gValue === 'object' && ('formula' in gValue || 'result' in gValue)));

      if ((hasCValue || hasDValue) && hasGValue) {
        console.log('Row ' + rowNum + ':');
        console.log('  C列:', cValue);
        console.log('  D列:', dValue);
        console.log('  G列:', typeof gValue === 'object' && 'formula' in gValue ? gValue.formula : gValue);
        console.log('  ✅ 条件を満たす\n');
      }
    }

    console.log('\n--- ヘッダー行(' + headerRowNum + ')以降でD列に数式がある行を検索 ---\n');

    let foundStartRow = null;

    for (let rowNum = headerRowNum + 1; rowNum <= headerRowNum + 50; rowNum++) {
      const row = interestSheet.getRow(rowNum);
      const dCell = row.getCell(4); // D列
      const dValue = dCell.value;

      // D列に数式がある
      const hasDFormula = dValue && typeof dValue === 'object' && 
        ('formula' in dValue || 'sharedFormula' in dValue);

      if (hasDFormula) {
        const formula = dValue.formula || dValue.sharedFormula;
        
        // 数式がG列を参照しているか確認（G10, G11, G12...）
        const gMatch = formula.match(/G(\d+)/);
        
        if (gMatch) {
          const gRowNum = parseInt(gMatch[1]);
          
          // そのG行のC列またはD列に数値があるかチェック
          const gRow = interestSheet.getRow(gRowNum);
          const gCCell = gRow.getCell(3);
          const gDCell = gRow.getCell(4);
          
          const hasCValue = typeof gCCell.value === 'number' && gCCell.value > 0;
          const hasDValue = typeof gDCell.value === 'number' && gDCell.value > 0;
          
          if (hasCValue || hasDValue) {
            console.log('Row ' + rowNum + ':');
            console.log('  D列数式:', formula);
            console.log('  参照: G' + gRowNum);
            console.log('  G' + gRowNum + 'のC列:', gCCell.value);
            console.log('  G' + gRowNum + 'のD列:', gDCell.value);
            console.log('  ✅ 開始行として適切\n');
            
            foundStartRow = rowNum;
            break;
          } else {
            console.log('Row ' + rowNum + ':');
            console.log('  D列数式:', formula);
            console.log('  参照: G' + gRowNum);
            console.log('  G' + gRowNum + 'のC列:', gCCell.value, '(数値なし)');
            console.log('  G' + gRowNum + 'のD列:', gDCell.value, '(数値なし)');
            console.log('  ❌ スキップ\n');
          }
        }
      }
    }

    if (foundStartRow) {
      console.log('🎯 最適な開始行: Row ' + foundStartRow);
    } else {
      console.log('⚠ 条件を満たす開始行が見つかりませんでした');
    }

  } catch (error) {
    console.error('エラー:', error.message);
  }
}

testNewLogic();
