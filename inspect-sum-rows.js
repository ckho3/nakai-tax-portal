const ExcelJS = require('exceljs');

async function inspectSumRows() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('========================================');
  console.log('各セクションの合計行（SUM）を確認');
  console.log('========================================\n');

  console.log('【B】セクションの合計行:');
  console.log('  元々のデータ範囲: Row 80-99（20行）');
  console.log('  合計行: Row 100');
  const v100 = sheet.getRow(100).getCell(22);
  const w100 = sheet.getRow(100).getCell(23);
  const x100 = sheet.getRow(100).getCell(24);
  console.log(`  V100: ${v100.formula || v100.value || '(空)'}`);
  console.log(`  W100: ${w100.formula || w100.value || '(空)'}`);
  console.log(`  X100: ${x100.formula || x100.value || '(空)'}`);
  console.log('');

  console.log('【C】セクションの合計行:');
  console.log('  元々のデータ範囲: Row 103-122（20行）');
  console.log('  合計行: Row 123');
  const v123 = sheet.getRow(123).getCell(22);
  const w123 = sheet.getRow(123).getCell(23);
  const x123 = sheet.getRow(123).getCell(24);
  console.log(`  V123: ${v123.formula || v123.value || '(空)'}`);
  console.log(`  W123: ${w123.formula || w123.value || '(空)'}`);
  console.log(`  X123: ${x123.formula || x123.value || '(空)'}`);
  console.log('');

  console.log('【D】セクションの合計行:');
  console.log('  元々のデータ範囲: Row 126-145（20行）');
  console.log('  合計行: Row 146');
  const v146 = sheet.getRow(146).getCell(22);
  const w146 = sheet.getRow(146).getCell(23);
  const x146 = sheet.getRow(146).getCell(24);
  console.log(`  V146: ${v146.formula || v146.value || '(空)'}`);
  console.log(`  W146: ${w146.formula || w146.value || '(空)'}`);
  console.log(`  X146: ${x146.formula || x146.value || '(空)'}`);
  console.log('');

  console.log('========================================');
  console.log('23件PDFの場合の修正後の位置と数式');
  console.log('========================================\n');

  const extraRows = 3;

  console.log('【B】セクション:');
  console.log('  データ範囲: Row 81-103（ヘッダー・注釈含む23行）');
  console.log('  データのみ: Row 83-103（21行）');
  console.log('  合計行: Row 104（元々Row 100）');
  console.log('  修正後の数式:');
  const bDataStart = 83 + extraRows;
  const bDataEnd = 99 + extraRows * 2;
  console.log(`    V104: SUM(V${bDataStart}:V${bDataEnd}) = SUM(V86:V105)`);
  console.log('');
  console.log('  ❌ 間違い！実際は:');
  console.log('    データ範囲: Row 83-105（23行）');
  console.log('    合計行: Row 106');
  console.log(`    V106: SUM(V83:V105)`);
  console.log('');

  console.log('待って、ユーザーの指摘を確認:');
  console.log('  「23件の場合、bセッションは103行まで」');
  console.log('  → データ最終行: Row 103');
  console.log('  「130行目は元々=SUM(V101:V120)」');
  console.log('  → これは【C】セクションの合計行！');
  console.log('');

  console.log('========================================');
  console.log('再確認: 各セクションの構造');
  console.log('========================================\n');

  console.log('原本（20件）:');
  console.log('  【B】データ: Row 80-99（20行）');
  console.log('  【B】合計: Row 100');
  console.log('  【C】ヘッダー: Row 101');
  console.log('  【C】注釈: Row 102');
  console.log('  【C】データ: Row 103-122（20行）');
  console.log('  【C】合計: Row 123');
  console.log('  【D】ヘッダー: Row 124');
  console.log('  【D】注釈: Row 125');
  console.log('  【D】データ: Row 126-145（20行）');
  console.log('  【D】合計: Row 146');
  console.log('');

  console.log('23件の場合:');
  console.log('  【B】ヘッダー: Row 81（元Row 78 + 3）');
  console.log('  【B】注釈: Row 82（元Row 79 + 3）');
  console.log('  【B】データ: Row 83-105（23行）');
  console.log('  【B】合計: Row 106（元Row 100 + 6）');
  console.log('  【C】ヘッダー: Row 107（元Row 101 + 6）');
  console.log('  【C】注釈: Row 108（元Row 102 + 6）');
  console.log('  【C】データ: Row 109-131（23行）');
  console.log('  【C】合計: Row 132（元Row 123 + 9）');
  console.log('  【D】ヘッダー: Row 133（元Row 124 + 9）');
  console.log('  【D】注釈: Row 134（元Row 125 + 9）');
  console.log('  【D】データ: Row 135-157（23行）');
  console.log('  【D】合計: Row 158（元Row 146 + 12）');
  console.log('');

  console.log('あれ？ユーザーの指摘と違う...');
  console.log('');
  console.log('ユーザーの指摘:');
  console.log('  - 【B】は103行まで → Row 103');
  console.log('  - 【C】の合計行は130行目 → Row 130');
  console.log('  - 【D】は155行目まで → Row 155');
  console.log('  - 【D】の合計行は156行目 → Row 156');
  console.log('');

  console.log('計算が合わない！再計算:');
  console.log('  【B】データ最終: Row 103');
  console.log('  【B】合計: Row 104（ではなく106？）');
  console.log('');
  console.log('もしかして、私の理解が間違っている？');
  console.log('');

  console.log('========================================');
  console.log('正しい理解（ユーザーの指摘に基づく）');
  console.log('========================================\n');

  console.log('23件PDFの場合:');
  console.log('');
  console.log('【B】セクション:');
  console.log('  データ最終行: Row 103');
  console.log('  合計行: Row 104？');
  console.log('');
  console.log('【C】セクション:');
  console.log('  合計行: Row 130');
  console.log('  → 元々Row 123だったのが+7行ずれた');
  console.log('  → 計算: 123 + 7 = 130 ✅');
  console.log('  → これは【A】+3、【B】+3、【C】+1の影響？');
  console.log('');
  console.log('【D】セクション:');
  console.log('  データ最終行: Row 155');
  console.log('  合計行: Row 156');
  console.log('  → 元々Row 146だったのが+10行ずれた');
  console.log('  → 計算: 146 + 10 = 156 ✅');
  console.log('');

  console.log('========================================');
  console.log('goal.xlsxで実際の構造を確認');
  console.log('========================================\n');

  const goalWorkbook = new ExcelJS.Workbook();
  await goalWorkbook.xlsx.readFile('/Users/user/Downloads/goal.xlsx');
  const goalSheet = goalWorkbook.getWorksheet('【不】①不動産収入');

  console.log('【B】セクションの構造:');
  for (let i = 103; i <= 107; i++) {
    const gValue = goalSheet.getRow(i).getCell(7).value;
    const vValue = goalSheet.getRow(i).getCell(22).value;
    const vFormula = goalSheet.getRow(i).getCell(22).formula;
    console.log(`  Row ${i}: G="${gValue}" V="${vFormula || vValue || '(空)'}"`);
  }
  console.log('');

  console.log('【C】セクションの構造:');
  for (let i = 129; i <= 133; i++) {
    const gValue = goalSheet.getRow(i).getCell(7).value;
    const vValue = goalSheet.getRow(i).getCell(22).value;
    const vFormula = goalSheet.getRow(i).getCell(22).formula;
    console.log(`  Row ${i}: G="${gValue}" V="${vFormula || vValue || '(空)'}"`);
  }
  console.log('');

  console.log('【D】セクションの構造:');
  for (let i = 155; i <= 159; i++) {
    const gValue = goalSheet.getRow(i).getCell(7).value;
    const vValue = goalSheet.getRow(i).getCell(22).value;
    const vFormula = goalSheet.getRow(i).getCell(22).formula;
    console.log(`  Row ${i}: G="${gValue}" V="${vFormula || vValue || '(空)'}"`);
  }
}

inspectSumRows().catch(console.error);
