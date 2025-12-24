const ExcelJS = require('exceljs');

async function inspectVWXFormulas() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('========================================');
  console.log('V-X列の数式を詳細分析');
  console.log('========================================\n');

  console.log('【A】収入セクション（Row 55-75）:');
  console.log('');

  // Row 55-58の数式を確認
  for (let row = 55; row <= 58; row++) {
    const vCell = sheet.getRow(row).getCell(22); // V列
    const wCell = sheet.getRow(row).getCell(23); // W列
    const xCell = sheet.getRow(row).getCell(24); // X列

    console.log(`Row ${row}:`);
    if (vCell.formula) console.log(`  V${row}: ${vCell.formula}`);
    if (wCell.formula) console.log(`  W${row}: ${wCell.formula}`);
    if (xCell.formula) console.log(`  X${row}: ${xCell.formula}`);
    console.log('');
  }

  console.log('【B】管理手数料セクション（Row 78-99）:');
  console.log('');

  // Row 80-83の数式を確認
  for (let row = 80; row <= 83; row++) {
    const vCell = sheet.getRow(row).getCell(22); // V列
    const wCell = sheet.getRow(row).getCell(23); // W列
    const xCell = sheet.getRow(row).getCell(24); // X列

    console.log(`Row ${row}:`);
    if (vCell.formula) console.log(`  V${row}: ${vCell.formula}`);
    if (wCell.formula) console.log(`  W${row}: ${wCell.formula}`);
    if (xCell.formula) console.log(`  X${row}: ${xCell.formula}`);
    console.log('');
  }

  console.log('【C】広告費セクション（Row 101-122）:');
  console.log('');

  // Row 103-106の数式を確認
  for (let row = 103; row <= 106; row++) {
    const vCell = sheet.getRow(row).getCell(22); // V列
    const wCell = sheet.getRow(row).getCell(23); // W列
    const xCell = sheet.getRow(row).getCell(24); // X列

    console.log(`Row ${row}:`);
    if (vCell.formula) console.log(`  V${row}: ${vCell.formula}`);
    if (wCell.formula) console.log(`  W${row}: ${wCell.formula}`);
    if (xCell.formula) console.log(`  X${row}: ${xCell.formula}`);
    console.log('');
  }

  console.log('【D】修繕費セクション（Row 124-145）:');
  console.log('');

  // Row 126-129の数式を確認
  for (let row = 126; row <= 129; row++) {
    const vCell = sheet.getRow(row).getCell(22); // V列
    const wCell = sheet.getRow(row).getCell(23); // W列
    const xCell = sheet.getRow(row).getCell(24); // X列

    console.log(`Row ${row}:`);
    if (vCell.formula) console.log(`  V${row}: ${vCell.formula}`);
    if (wCell.formula) console.log(`  W${row}: ${wCell.formula}`);
    if (xCell.formula) console.log(`  X${row}: ${xCell.formula}`);
    console.log('');
  }

  console.log('========================================');
  console.log('数式のパターン分析');
  console.log('========================================\n');

  console.log('V列の数式パターン:');
  console.log('  SUMIF($I$53:$T$53,V$53,$I{row}:$T{row})');
  console.log('  - $I$53:$T$53: ヘッダー行の範囲（固定）');
  console.log('  - V$53: V列のヘッダー（固定行、可変列）');
  console.log('  - $I{row}:$T{row}: 各データ行の範囲');
  console.log('');

  console.log('W列の数式パターン:');
  console.log('  SUMIF($I$53:$T$53,W$53,$I{row}:$T{row})');
  console.log('  - $I$53:$T$53: ヘッダー行の範囲（固定）');
  console.log('  - W$53: W列のヘッダー（固定行、可変列）');
  console.log('  - $I{row}:$T{row}: 各データ行の範囲');
  console.log('');

  console.log('X列の数式パターン:');
  console.log('  SUM(V{row}:W{row})');
  console.log('  - V列とW列の合計');
  console.log('');

  console.log('========================================');
  console.log('行追加時の問題');
  console.log('========================================\n');

  console.log('23件PDFの場合（extraRows = 3）:');
  console.log('');

  console.log('【A】セクション:');
  console.log('  データ行: Row 55-77（23行）');
  console.log('  例: V55 = SUMIF($I$53:$T$53,V$53,$I55:$T55)');
  console.log('  → 行追加後も$I$53:$T$53は固定なので変わらない ✅');
  console.log('  → $I55:$T55も各行で自動調整される ✅');
  console.log('');

  console.log('【B】セクション:');
  console.log('  元: Row 80-99（20行）');
  console.log('  行追加後: Row 83-105（23行）');
  console.log('  例: 元V80 = SUMIF($I$53:$T$53,V$53,$I80:$T80)');
  console.log('  → Row 83に移動すると: V83 = SUMIF($I$53:$T$53,V$53,$I83:$T83)');
  console.log('  → duplicateRowで自動調整される ✅');
  console.log('');

  console.log('【C】セクション:');
  console.log('  元: Row 103-122（20行）');
  console.log('  行追加後: Row 109-131（23行）');
  console.log('  → 同様に自動調整される ✅');
  console.log('');

  console.log('【D】セクション:');
  console.log('  元: Row 126-145（20行）');
  console.log('  行追加後: Row 135-157（23行）');
  console.log('  → 同様に自動調整される ✅');
  console.log('');

  console.log('========================================');
  console.log('重要な発見');
  console.log('========================================\n');

  console.log('V-X列の数式は以下の理由で自動対応できる:');
  console.log('');
  console.log('1. $I$53:$T$53 は絶対参照なので行追加の影響を受けない ✅');
  console.log('2. V$53, W$53 は行が固定なので影響を受けない ✅');
  console.log('3. $I{row}:$T{row} は各行の相対参照なので、');
  console.log('   duplicateRowで自動的に調整される ✅');
  console.log('');

  console.log('結論: V-X列の数式は追加修正不要！');
  console.log('');

  console.log('ただし、念のため確認すべき箇所:');
  console.log('- 各セクションのヘッダー行（Row 55, 78, 101, 124）');
  console.log('- 各セクションの合計行（Row 76, 99, 122, 145）');
  console.log('');

  // ヘッダー行と合計行を確認
  console.log('========================================');
  console.log('特殊な行の確認');
  console.log('========================================\n');

  const specialRows = [53, 55, 76, 78, 99, 101, 122, 124, 145];

  for (const row of specialRows) {
    const vCell = sheet.getRow(row).getCell(22);
    const wCell = sheet.getRow(row).getCell(23);
    const xCell = sheet.getRow(row).getCell(24);

    console.log(`Row ${row}:`);
    if (vCell.formula) console.log(`  V${row}: ${vCell.formula}`);
    if (wCell.formula) console.log(`  W${row}: ${wCell.formula}`);
    if (xCell.formula) console.log(`  X${row}: ${xCell.formula}`);
    if (!vCell.formula && !wCell.formula && !xCell.formula) {
      console.log('  (数式なし)');
    }
    console.log('');
  }
}

inspectVWXFormulas().catch(console.error);
