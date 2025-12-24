const ExcelJS = require('exceljs');
const path = require('path');

async function inspectExcel() {
  try {
    const workbook = new ExcelJS.Workbook();
    const excelPath = path.join(__dirname, '【原本】R7確定申告フォーマット.xlsx');
    await workbook.xlsx.readFile(excelPath);

    const sheet = workbook.getWorksheet('【不】①不動産収入');

    if (!sheet) {
      console.log('シート【不】①不動産収入が見つかりません');
      return;
    }

    console.log('========================================');
    console.log('Row 50-150 の全データ確認 (E-I列)');
    console.log('========================================\n');

    for (let rowNum = 50; rowNum <= 150; rowNum++) {
      const row = sheet.getRow(rowNum);
      const cellA = row.getCell(1).value;
      const cellB = row.getCell(2).value;
      const cellC = row.getCell(3).value;
      const cellD = row.getCell(4).value;
      const cellE = row.getCell(5).value;
      const cellF = row.getCell(6).value;
      const cellG = row.getCell(7).value;
      const cellH = row.getCell(8).value;
      const cellI = row.getCell(9).value;

      // 少なくとも1つのセルに値がある行のみ表示
      if (cellA || cellB || cellC || cellD || cellE || cellF || cellG || cellH || cellI) {
        const formatCell = (val) => {
          if (!val) return '(空)';
          // 数式オブジェクトの場合
          if (typeof val === 'object' && val.formula) {
            return `[Formula: ${val.formula.substring(0, 30)}]`;
          }
          if (typeof val === 'object' && val.result !== undefined) {
            return `[Result: ${val.result}]`;
          }
          const str = val.toString();
          return str.length > 40 ? str.substring(0, 40) + '...' : str;
        };

        // A-D列に重要な情報がある行のみ表示
        const hasImportantData = cellA || cellB || cellC || cellD ||
                                 (cellG && typeof cellG === 'string') ||
                                 (cellH && typeof cellH === 'string');

        if (hasImportantData) {
          console.log(`Row ${rowNum}:`);
          if (cellA) console.log(`  A="${formatCell(cellA)}"`);
          if (cellB) console.log(`  B="${formatCell(cellB)}"`);
          if (cellC) console.log(`  C="${formatCell(cellC)}"`);
          if (cellD) console.log(`  D="${formatCell(cellD)}"`);
          if (cellE) console.log(`  E="${formatCell(cellE)}"`);
          if (cellF) console.log(`  F="${formatCell(cellF)}"`);
          if (cellG) console.log(`  G="${formatCell(cellG)}"`);
          if (cellH) console.log(`  H="${formatCell(cellH)}"`);
          if (cellI) console.log(`  I="${formatCell(cellI)}"`);
          console.log('');
        }
      }
    }

  } catch (error) {
    console.error('エラー:', error);
  }
}

inspectExcel();
