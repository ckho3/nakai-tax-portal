const ExcelJS = require('exceljs');
const path = require('path');

async function cleanExcelFormulas(inputPath, outputPath) {
  console.log('Excelファイルから数式を削除しています...');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputPath);

  // すべてのワークシートを処理
  workbook.eachSheet((worksheet) => {
    console.log(`シート: ${worksheet.name} を処理中...`);

    // すべての行を処理
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        // 数式がある場合は値のみ保持
        if (cell.formula) {
          const value = cell.value;
          cell.value = value && typeof value === 'object' && value.result !== undefined
            ? value.result
            : value;
          cell.formula = undefined;
        }
        // 共有数式も削除
        if (cell.sharedFormula) {
          cell.sharedFormula = undefined;
        }
      });
    });
  });

  await workbook.xlsx.writeFile(outputPath);
  console.log(`完了: ${outputPath}`);
}

// コマンドライン引数から実行
if (require.main === module) {
  const args = process.argv.slice(2);
  if (args.length !== 2) {
    console.log('使い方: node cleanExcel.js <input.xlsx> <output.xlsx>');
    process.exit(1);
  }

  cleanExcelFormulas(args[0], args[1])
    .then(() => console.log('成功'))
    .catch(err => {
      console.error('エラー:', err);
      process.exit(1);
    });
}

module.exports = { cleanExcelFormulas };
