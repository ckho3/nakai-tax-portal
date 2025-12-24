const ExcelJS = require('exceljs');

async function checkFormulas() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  if (!sheet) {
    console.log('シートが見つかりません');
    return;
  }

  console.log('\n=== 数式が含まれているセルを検索中... ===\n');

  let formulaCount = 0;
  let sharedFormulaCount = 0;
  const formulaRanges = [];

  for (let rowNum = 1; rowNum <= sheet.rowCount; rowNum++) {
    const row = sheet.getRow(rowNum);
    for (let colNum = 1; colNum <= 30; colNum++) {
      const cell = row.getCell(colNum);

      if (cell.formula || cell.sharedFormula) {
        formulaCount++;

        if (cell.sharedFormula) {
          sharedFormulaCount++;
        }

        const colLetter = String.fromCharCode(64 + colNum);
        const cellRef = `${colLetter}${rowNum}`;

        formulaRanges.push({
          cell: cellRef,
          row: rowNum,
          type: cell.sharedFormula ? '共有数式' : '通常数式',
          formula: cell.formula || cell.sharedFormula
        });
      }
    }
  }

  console.log(`総数式セル数: ${formulaCount}`);
  console.log(`共有数式セル数: ${sharedFormulaCount}\n`);

  // 行ごとにグループ化
  const rowGroups = {};
  formulaRanges.forEach(item => {
    if (!rowGroups[item.row]) {
      rowGroups[item.row] = [];
    }
    rowGroups[item.row].push(item);
  });

  // 最初の50行分を表示
  console.log('=== 数式が含まれている行（最初の100行） ===\n');
  Object.keys(rowGroups).sort((a, b) => parseInt(a) - parseInt(b)).slice(0, 100).forEach(rowNum => {
    console.log(`Row ${rowNum}:`);
    rowGroups[rowNum].forEach(item => {
      console.log(`  ${item.cell}: ${item.type}`);
    });
  });

  // 55行目以降（データ入力エリア）の数式を確認
  console.log('\n=== Row 55以降（データ入力エリア）の数式 ===\n');
  Object.keys(rowGroups).filter(r => parseInt(r) >= 55).forEach(rowNum => {
    console.log(`Row ${rowNum}:`);
    rowGroups[rowNum].forEach(item => {
      console.log(`  ${item.cell}: ${item.type} = ${item.formula}`);
    });
  });
}

checkFormulas().catch(console.error);
