const ExcelJS = require('exceljs');

async function verifyC95() {
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

    console.log('===== 【不】⑤利息シート Row 95の確認 =====\n');

    const row95 = interestSheet.getRow(95);
    
    console.log('Row 95:');
    console.log('  C95:', row95.getCell(3).value);
    console.log('  D95:', row95.getCell(4).value);
    
    const c95Value = row95.getCell(3).value;
    
    if (c95Value === 0.8) {
      console.log('\n✅ C95に0.8が正しく書き込まれています！');
    } else {
      console.log('\n❌ C95の値が0.8ではありません:', c95Value);
    }

  } catch (error) {
    console.error('エラー:', error.message);
  }
}

verifyC95();
