const ExcelJS = require('exceljs');

async function inspectSectionE() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Downloads/goal.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  if (!sheet) {
    console.log('シート【不】①不動産収入が見つかりません');
    return;
  }

  console.log('========================================');
  console.log('goal.xlsxの【E】サブリースセクションを確認');
  console.log('========================================\n');

  // 【E】セクションの開始位置を探す
  console.log('【E】セクションの位置を探索:');
  let eSectionStart = null;
  for (let i = 150; i <= 180; i++) {
    const cValue = sheet.getRow(i).getCell(3).value;
    if (cValue && cValue.toString().includes('【E】')) {
      eSectionStart = i;
      console.log(`  【E】セクション開始: Row ${i}`);
      console.log(`  C${i}: "${cValue}"`);
      break;
    }
  }
  console.log('');

  if (!eSectionStart) {
    console.log('【E】セクションが見つかりませんでした');
    return;
  }

  // 【E】セクションのデータ行を確認
  console.log('【E】セクションのデータ行:');
  let eDataCount = 0;
  for (let i = eSectionStart; i <= eSectionStart + 30; i++) {
    const hValue = sheet.getRow(i).getCell(8).value;
    if (hValue && typeof hValue !== 'object') {
      eDataCount++;
      if (eDataCount <= 5 || eDataCount >= 20) {
        console.log(`  Row ${i}: H="${hValue}"`);
      } else if (eDataCount === 6) {
        console.log('  ...');
      }
    }
  }
  console.log(`\n【E】セクションのデータ行数: ${eDataCount}行\n`);

  // H列の数式を確認
  console.log('========================================');
  console.log('H列の数式を確認（物件名の参照）');
  console.log('========================================\n');

  const firstDataRow = eSectionStart;
  for (let i = firstDataRow; i <= firstDataRow + 5; i++) {
    const cell = sheet.getRow(i).getCell(8);
    if (cell.formula) {
      console.log(`  Row ${i} H列の数式: ${cell.formula}`);
    } else if (cell.value) {
      console.log(`  Row ${i} H列の値: ${cell.value}`);
    }
  }
  console.log('');

  // I列の数式を確認
  console.log('========================================');
  console.log('I列の数式を確認（1月のサブリース計算）');
  console.log('========================================\n');

  for (let i = firstDataRow; i <= firstDataRow + 2; i++) {
    const cell = sheet.getRow(i).getCell(9); // I列
    if (cell.formula) {
      console.log(`  Row ${i} I列の数式: ${cell.formula}`);
    }
  }
  console.log('');

  // T列の数式を確認
  console.log('========================================');
  console.log('T列の数式を確認（12月のサブリース計算）');
  console.log('========================================\n');

  for (let i = firstDataRow; i <= firstDataRow + 2; i++) {
    const cell = sheet.getRow(i).getCell(20); // T列
    if (cell.formula) {
      console.log(`  Row ${i} T列の数式: ${cell.formula}`);
    }
  }
  console.log('');

  // 原本との比較
  console.log('========================================');
  console.log('原本との比較');
  console.log('========================================\n');

  const originalWorkbook = new ExcelJS.Workbook();
  await originalWorkbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');
  const originalSheet = originalWorkbook.getWorksheet('【不】①不動産収入');

  console.log('原本の【E】セクション:');
  for (let i = 147; i <= 150; i++) {
    const cValue = originalSheet.getRow(i).getCell(3).value;
    const hCell = originalSheet.getRow(i).getCell(8);
    const hFormula = hCell.formula || hCell.value;
    console.log(`  Row ${i}: C="${cValue}" H="${hFormula}"`);
  }
  console.log('');

  console.log(`goal.xlsxの【E】セクション開始: Row ${eSectionStart}`);
  console.log(`原本の【E】セクション開始: Row 147`);
  console.log(`差分: +${eSectionStart - 147}行\n`);

  console.log('========================================');
  console.log('結論');
  console.log('========================================\n');

  console.log('1. 【E】セクションも行が追加されている');
  console.log(`2. 原本はRow 147-166（20行）`);
  console.log(`3. goal.xlsxはRow ${eSectionStart}-${eSectionStart + eDataCount - 1}（${eDataCount}行）`);
  console.log(`4. つまり、【E】セクションにも${eDataCount - 20}行追加されている`);
}

inspectSectionE().catch(console.error);
