const ExcelJS = require('exceljs');

async function inspectSummaryRows() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('========================================');
  console.log('各セクションの合計行の数式を確認');
  console.log('========================================\n');

  // Row 76 (【A】の合計行)
  console.log('Row 76（【A】収入セクションの合計行）:');
  console.log('  G76: ' + sheet.getRow(76).getCell(7).value);
  for (let col = 9; col <= 20; col++) {
    const cell = sheet.getRow(76).getCell(col);
    const colLetter = String.fromCharCode(64 + col);
    if (cell.formula) {
      console.log(`  ${colLetter}76: ${cell.formula}`);
    }
  }
  console.log('');

  // Row 99 (【B】の合計行)
  console.log('Row 99（【B】管理手数料セクションの合計行）:');
  console.log('  G99: ' + sheet.getRow(99).getCell(7).value);
  for (let col = 9; col <= 20; col++) {
    const cell = sheet.getRow(99).getCell(col);
    const colLetter = String.fromCharCode(64 + col);
    if (cell.formula) {
      console.log(`  ${colLetter}99: ${cell.formula}`);
    }
  }
  console.log('');

  // Row 122 (【C】の合計行)
  console.log('Row 122（【C】広告費セクションの合計行）:');
  console.log('  G122: ' + sheet.getRow(122).getCell(7).value);
  for (let col = 9; col <= 20; col++) {
    const cell = sheet.getRow(122).getCell(col);
    const colLetter = String.fromCharCode(64 + col);
    if (cell.formula) {
      console.log(`  ${colLetter}122: ${cell.formula}`);
    }
  }
  console.log('');

  // Row 145 (【D】の合計行)
  console.log('Row 145（【D】修繕費セクションの合計行）:');
  console.log('  G145: ' + sheet.getRow(145).getCell(7).value);
  for (let col = 9; col <= 20; col++) {
    const cell = sheet.getRow(145).getCell(col);
    const colLetter = String.fromCharCode(64 + col);
    if (cell.formula) {
      console.log(`  ${colLetter}145: ${cell.formula}`);
    }
  }
  console.log('');

  console.log('========================================');
  console.log('パターン分析');
  console.log('========================================\n');

  console.log('Row 76（【A】の合計）:');
  console.log('  I76 = I53  （【A】のヘッダー行を参照）');
  console.log('  ...');
  console.log('  T76 = T53');
  console.log('');

  console.log('Row 99（【B】の合計）:');
  console.log('  I99 = I76  （【A】の合計行を参照）');
  console.log('  ...');
  console.log('  T99 = T76');
  console.log('');

  console.log('Row 122（【C】の合計）:');
  console.log('  I122 = I99  （【B】の合計行を参照）');
  console.log('  ...');
  console.log('  T122 = T99');
  console.log('');

  console.log('Row 145（【D】の合計）:');
  console.log('  I145 = I122  （【C】の合計行を参照）');
  console.log('  ...');
  console.log('  T145 = T122');
  console.log('');

  console.log('========================================');
  console.log('行追加時の問題');
  console.log('========================================\n');

  console.log('23件PDFの場合（extraRows = 3）:');
  console.log('');

  console.log('【A】セクション:');
  console.log('  Row 55-77（23行）');
  console.log('  合計行: Row 78（元々Row 76）← +3行ずれる');
  console.log('');

  console.log('【B】セクション:');
  console.log('  Row 81-103（23行）');
  console.log('  合計行: Row 104（元々Row 99）← +5行ずれる（自セクション+3、前セクション+3の影響で+6）');
  console.log('');

  console.log('【C】セクション:');
  console.log('  Row 107-129（23行）');
  console.log('  合計行: Row 130（元々Row 122）← +8行ずれる');
  console.log('');

  console.log('【D】セクション:');
  console.log('  Row 133-155（23行）');
  console.log('  合計行: Row 156（元々Row 145）← +11行ずれる');
  console.log('');

  console.log('========================================');
  console.log('数式のずれ問題');
  console.log('========================================\n');

  console.log('Row 78（【A】の合計行）:');
  console.log('  元: I76 = I53');
  console.log('  行追加後: I78 = I53（duplicateRowで自動調整される）✅');
  console.log('');

  console.log('Row 104（【B】の合計行）:');
  console.log('  元: I99 = I76');
  console.log('  行追加後: I104 = I82（duplicateRowで+6される）');
  console.log('  しかし、【A】の合計行はRow 78にある ❌');
  console.log('  正しくは: I104 = I78 とすべき');
  console.log('');

  console.log('Row 130（【C】の合計行）:');
  console.log('  元: I122 = I99');
  console.log('  行追加後: I130 = I107（duplicateRowで+8される）');
  console.log('  しかし、【B】の合計行はRow 104にある ❌');
  console.log('  正しくは: I130 = I104 とすべき');
  console.log('');

  console.log('Row 156（【D】の合計行）:');
  console.log('  元: I145 = I122');
  console.log('  行追加後: I156 = I133（duplicateRowで+11される）');
  console.log('  しかし、【C】の合計行はRow 130にある ❌');
  console.log('  正しくは: I156 = I130 とすべき');
  console.log('');

  console.log('========================================');
  console.log('必要な修正');
  console.log('========================================\n');

  console.log('行追加後、各セクションの合計行の数式を修正:');
  console.log('');
  console.log('1. 【A】の合計行（Row 76 + extraRows）:');
  console.log('   I列-T列: =I53, =J53, ..., =T53（変更なし）');
  console.log('');
  console.log('2. 【B】の合計行（Row 99 + extraRows * 2）:');
  console.log('   I列-T列: =I76 → =I(【A】の合計行)');
  console.log('');
  console.log('3. 【C】の合計行（Row 122 + extraRows * 3）:');
  console.log('   I列-T列: =I99 → =I(【B】の合計行)');
  console.log('');
  console.log('4. 【D】の合計行（Row 145 + extraRows * 4）:');
  console.log('   I列-T列: =I122 → =I(【C】の合計行)');
}

inspectSummaryRows().catch(console.error);
