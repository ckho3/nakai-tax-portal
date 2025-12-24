const ExcelJS = require('exceljs');

async function inspectEFormulas() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('========================================');
  console.log('【E】セクションの数式を確認（原本）');
  console.log('========================================\n');

  console.log('Row 147-166 の全ての数式を確認:\n');

  // H列からT列までの数式を確認
  const columns = {
    7: 'G',
    8: 'H',
    9: 'I',
    10: 'J',
    11: 'K',
    12: 'L',
    13: 'M',
    14: 'N',
    15: 'O',
    16: 'P',
    17: 'Q',
    18: 'R',
    19: 'S',
    20: 'T',
    21: 'U'
  };

  // 最初の3行を詳しく確認
  for (let row = 147; row <= 149; row++) {
    console.log(`Row ${row}:`);
    for (let col = 7; col <= 21; col++) {
      const cell = sheet.getRow(row).getCell(col);
      if (cell.formula) {
        console.log(`  ${columns[col]}${row}: ${cell.formula}`);
      }
    }
    console.log('');
  }

  console.log('========================================');
  console.log('数式のパターン分析');
  console.log('========================================\n');

  const row147 = sheet.getRow(147);

  console.log('Row 147の数式:');
  console.log('  H147: ' + row147.getCell(8).formula);
  console.log('  I147: ' + row147.getCell(9).formula);
  console.log('  J147: ' + row147.getCell(10).formula);
  console.log('  T147: ' + row147.getCell(20).formula);
  console.log('  U147: ' + row147.getCell(21).formula);
  console.log('');

  console.log('Row 148の数式:');
  console.log('  H148: ' + row147.getCell(8).formula);
  console.log('  I148: ' + sheet.getRow(148).getCell(9).formula);
  console.log('  T148: ' + sheet.getRow(148).getCell(20).formula);
  console.log('');

  console.log('========================================');
  console.log('問題点の分析');
  console.log('========================================\n');

  console.log('H列の数式: =H124, =H125, =H126...');
  console.log('  → 【D】セクションの物件名を参照');
  console.log('  → 行追加すると、【D】がRow 133-155に移動');
  console.log('  → H147は自動的にH136になるが、本来はH133を参照すべき');
  console.log('');

  console.log('T列の数式: ROUNDDOWN((T55-T78)*$B$146,-2)');
  console.log('  → T55: 【A】の1行目（行追加後も変わらない）✅');
  console.log('  → T78: 【B】の元の位置');
  console.log('  → 行追加後は【B】がRow 81から始まる');
  console.log('  → T78はT90になるが、本来はT81を参照すべき ❌');
  console.log('');

  console.log('========================================');
  console.log('必要な対応');
  console.log('========================================\n');

  console.log('行追加後、【E】セクションの数式を修正する必要があります:');
  console.log('');
  console.log('1. H列の数式を修正:');
  console.log('   Row 159: =H124 → =H133（【D】の1行目）');
  console.log('   Row 160: =H125 → =H134（【D】の2行目）');
  console.log('   ...');
  console.log('');
  console.log('2. I-T列の数式内の参照を修正:');
  console.log('   (T55-T78) → (T55-T81)');
  console.log('   T55は【A】の1行目（固定）');
  console.log('   T78→T81は【B】の1行目（行追加後）');
  console.log('');

  console.log('========================================');
  console.log('解決策');
  console.log('========================================\n');

  console.log('Option 1: 行追加後に数式を書き換える ✅');
  console.log('  - 行追加実行後、【E】セクションの数式を正しい参照に修正');
  console.log('  - H列: =(【D】の開始行 + 行オフセット)');
  console.log('  - T列の数式内: (T55-T(【B】の開始行))に修正');
  console.log('');

  console.log('Option 2: 元のテンプレートで絶対参照にする');
  console.log('  - H147: =$H$124 とする');
  console.log('  - 問題: 行追加後も124を参照し続けてしまう');
  console.log('  - 動的に参照先を変更できない');
  console.log('');

  console.log('Option 3: OFFSET関数を使う');
  console.log('  - H147: =OFFSET($H$124, ROW()-147, 0)');
  console.log('  - 複雑で、テンプレート変更が必要');
  console.log('');

  console.log('推奨: Option 1（行追加後にプログラムで数式を修正）');
}

inspectEFormulas().catch(console.error);
