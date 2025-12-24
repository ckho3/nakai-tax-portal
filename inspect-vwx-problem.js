const ExcelJS = require('exceljs');

async function inspectVWXProblem() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('========================================');
  console.log('V-X列の問題を詳細分析');
  console.log('========================================\n');

  console.log('【A】セクションの最終データ行（Row 74）の数式:');
  const v74 = sheet.getRow(74).getCell(22);
  const w74 = sheet.getRow(74).getCell(23);
  const x74 = sheet.getRow(74).getCell(24);
  console.log(`  V74: ${v74.formula}`);
  console.log(`  W74: ${w74.formula}`);
  console.log(`  X74: ${x74.formula}`);
  console.log('');

  console.log('【A】セクションの75行目（元々空行）:');
  const v75 = sheet.getRow(75).getCell(22);
  const w75 = sheet.getRow(75).getCell(23);
  const x75 = sheet.getRow(75).getCell(24);
  console.log(`  V75: ${v75.formula || v75.value || '(空)'}`);
  console.log(`  W75: ${w75.formula || w75.value || '(空)'}`);
  console.log(`  X75: ${x75.formula || x75.value || '(空)'}`);
  console.log('');

  console.log('========================================');
  console.log('行追加の問題シミュレーション');
  console.log('========================================\n');

  console.log('duplicateRow(73, 1, true)を実行した場合:');
  console.log('  Row 73の内容をコピーしてRow 74に挿入');
  console.log('  元のRow 74以降が1行下にずれる');
  console.log('');

  console.log('Row 73の数式:');
  const v73 = sheet.getRow(73).getCell(22);
  const w73 = sheet.getRow(73).getCell(23);
  const x73 = sheet.getRow(73).getCell(24);
  console.log(`  V73: ${v73.formula}`);
  console.log(`  W73: ${w73.formula}`);
  console.log(`  X73: ${x73.formula}`);
  console.log('');

  console.log('duplicateRow後、新しいRow 74の数式（予想）:');
  console.log('  V74: SUMIF($I$53:$T$53,V$53,$I74:$T74) ← V73から自動調整');
  console.log('  W74: SUMIF($I$53:$T$53,W$53,$I74:$T74) ← W73から自動調整');
  console.log('  X74: SUM(V74:W74) ← X73から自動調整');
  console.log('');

  console.log('元のRow 74は Row 75に移動:');
  console.log(`  V75: SUMIF($I$53:$T$53,V$53,$I75:$T75) ← 元のV74から自動調整`);
  console.log('');

  console.log('========================================');
  console.log('問題点の発見');
  console.log('========================================\n');

  console.log('23件PDFの場合（3回duplicateRow実行）:');
  console.log('');

  console.log('1回目: duplicateRow(73, 1, true)');
  console.log('  Row 55-74: 【A】のデータ（20行）');
  console.log('  Row 74: V74 = SUMIF($I$53:$T$53,V$53,$I74:$T74) ← 新しく追加された行 ✅');
  console.log('  Row 75: V75 = SUMIF($I$53:$T$53,V$53,$I75:$T75) ← 元のRow 74');
  console.log('  Row 76: V76 = (空) ← 元のRow 75');
  console.log('');

  console.log('2回目: duplicateRow(74, 1, true)');
  console.log('  Row 55-75: 【A】のデータ（21行）');
  console.log('  Row 75: V75 = SUMIF($I$53:$T$53,V$53,$I75:$T75) ← 新しく追加された行 ✅');
  console.log('  Row 76: V76 = SUMIF($I$53:$T$53,V$53,$I76:$T76) ← 元のRow 75（元々のRow 74）');
  console.log('  Row 77: V77 = (空) ← 元のRow 76（元々のRow 75）');
  console.log('');

  console.log('3回目: duplicateRow(75, 1, true)');
  console.log('  Row 55-76: 【A】のデータ（22行）');
  console.log('  Row 76: V76 = SUMIF($I$53:$T$53,V$53,$I76:$T76) ← 新しく追加された行 ✅');
  console.log('  Row 77: V77 = SUMIF($I$53:$T$53,V$53,$I77:$T77) ← 元のRow 76（元々のRow 74）');
  console.log('  Row 78: V78 = (空) ← 元のRow 77（元々のRow 75）');
  console.log('');

  console.log('❌ 問題: Row 77には数式があるが、これは【A】の23行目のデータ行!');
  console.log('   本来、【A】の最終データ行（Row 77）には数式があるべき ✅');
  console.log('   しかし、その数式は元々Row 74のもの（20行目）');
  console.log('');

  console.log('待って...これは実は正しいのでは？');
  console.log('');
  console.log('確認:');
  console.log('  - duplicateRow(73)で、Row 73の数式がRow 74にコピーされる');
  console.log('  - Row 74: SUMIF($I$53:$T$53,V$53,$I74:$T74) ← 正しい ✅');
  console.log('  - 次にduplicateRow(74)で、Row 74の数式がRow 75にコピーされる');
  console.log('  - Row 75: SUMIF($I$53:$T$53,V$53,$I75:$T75) ← 正しい ✅');
  console.log('');
  console.log('結論: データ行（Row 55-77）の数式は自動調整される ✅');
  console.log('');

  console.log('========================================');
  console.log('本当の問題箇所');
  console.log('========================================\n');

  console.log('問題は「合計行」と「ヘッダー行」:');
  console.log('');

  console.log('【A】の合計行（元々Row 76）:');
  const v76 = sheet.getRow(76).getCell(22);
  const w76 = sheet.getRow(76).getCell(23);
  const x76 = sheet.getRow(76).getCell(24);
  console.log(`  V76: ${v76.formula || v76.value || '(空)'}`);
  console.log(`  W76: ${w76.formula || w76.value || '(空)'}`);
  console.log(`  X76: ${x76.formula || x76.value || '(空)'}`);
  console.log('');

  console.log('あれ？Row 76には数式がない！');
  console.log('');

  console.log('【B】セクションのヘッダー行（Row 78）を確認:');
  const v78 = sheet.getRow(78).getCell(22);
  console.log(`  V78: ${v78.formula || v78.value || '(空)'}`);
  console.log('');

  console.log('【B】セクションの最初のデータ行（Row 80）を確認:');
  const v80 = sheet.getRow(80).getCell(22);
  console.log(`  V80: ${v80.formula}`);
  console.log('');

  console.log('========================================');
  console.log('各セクションのデータ開始行を確認');
  console.log('========================================\n');

  console.log('【A】セクション:');
  console.log('  Row 55（ヘッダー兼データ1行目）: ' + sheet.getRow(55).getCell(22).formula);
  console.log('  Row 56（データ2行目）: ' + sheet.getRow(56).getCell(22).formula);
  console.log('');

  console.log('【B】セクション:');
  console.log('  Row 78（ヘッダー）: ' + (sheet.getRow(78).getCell(22).formula || sheet.getRow(78).getCell(3).value));
  console.log('  Row 79（注釈）: ' + (sheet.getRow(79).getCell(22).formula || sheet.getRow(79).getCell(3).value));
  console.log('  Row 80（データ1行目）: ' + sheet.getRow(80).getCell(22).formula);
  console.log('');

  console.log('あ！【B】セクションは Row 80 から数式が始まる！');
  console.log('Row 78-79はヘッダーと注釈行で数式がない');
  console.log('');

  console.log('========================================');
  console.log('真の問題');
  console.log('========================================\n');

  console.log('【A】セクション（Row 55-74、20行）:');
  console.log('  - 全行に数式がある（Row 55-74）');
  console.log('  - duplicateRow(73)で追加された行も数式がコピーされる ✅');
  console.log('');

  console.log('【B】セクション（Row 80-99、20行）:');
  console.log('  - Row 78-79: ヘッダー・注釈（数式なし）');
  console.log('  - Row 80-99: データ行（数式あり）');
  console.log('');

  console.log('問題: duplicateRow(97, 1, true)を実行すると:');
  console.log('  - Row 97の数式: SUMIF($I$53:$T$53,V$53,$I97:$T97)');
  console.log('  - 新しいRow 98: SUMIF($I$53:$T$53,V$53,$I98:$T98) ✅');
  console.log('  - 元のRow 98は Row 99に移動');
  console.log('  - 元のRow 99（合計行、数式なし）は Row 100に移動');
  console.log('');

  console.log('結論: データ行の数式は正しく自動調整される！');
  console.log('');
  console.log('あなたの指摘は正しいですか？具体的にどの行が問題ですか？');
  console.log('');

  console.log('========================================');
  console.log('goal.xlsxで確認してみましょう');
  console.log('========================================\n');

  const goalWorkbook = new ExcelJS.Workbook();
  await goalWorkbook.xlsx.readFile('/Users/user/Downloads/goal.xlsx');
  const goalSheet = goalWorkbook.getWorksheet('【不】①不動産収入');

  console.log('goal.xlsx（23件）の【A】セクション最終行（Row 77）:');
  const goalV77 = goalSheet.getRow(77).getCell(22);
  const goalW77 = goalSheet.getRow(77).getCell(23);
  const goalX77 = goalSheet.getRow(77).getCell(24);
  console.log(`  V77: ${goalV77.formula || goalV77.value || '(空)'}`);
  console.log(`  W77: ${goalW77.formula || goalW77.value || '(空)'}`);
  console.log(`  X77: ${goalX77.formula || goalX77.value || '(空)'}`);
  console.log('');

  console.log('goal.xlsxの【B】セクション最初のデータ行（Row 83）:');
  const goalV83 = goalSheet.getRow(83).getCell(22);
  console.log(`  V83: ${goalV83.formula || goalV83.value || '(空)'}`);
  console.log('');

  console.log('goal.xlsxの【B】セクション最終データ行（Row 103）:');
  const goalV103 = goalSheet.getRow(103).getCell(22);
  console.log(`  V103: ${goalV103.formula || goalV103.value || '(空)'}`);
}

inspectVWXProblem().catch(console.error);
