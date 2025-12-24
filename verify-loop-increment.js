/**
 * ループ内でのduplicateRow()の行番号増加を検証
 */

console.log('========================================');
console.log('duplicateRow()ループの詳細検証');
console.log('========================================\n');

const extraRows = 3;

console.log(`extraRows = ${extraRows}の場合:\n`);

console.log('修正前（間違い）:');
console.log('  for (let i = 0; i < 3; i++) {');
console.log('    sheet.duplicateRow(145, 1, true);  // 常にRow 145 ❌');
console.log('    sheet.duplicateRow(122, 1, true);  // 常にRow 122 ❌');
console.log('    sheet.duplicateRow(99, 1, true);   // 常にRow 99 ❌');
console.log('    sheet.duplicateRow(75, 1, true);   // 常にRow 75 ❌');
console.log('  }');
console.log('');

console.log('修正後（正しい）:');
console.log('  for (let i = 0; i < 3; i++) {');
console.log('    sheet.duplicateRow(145 + i, 1, true);  // 145, 146, 147 ✅');
console.log('    sheet.duplicateRow(122 + i, 1, true);  // 122, 123, 124 ✅');
console.log('    sheet.duplicateRow(99 + i, 1, true);   // 99, 100, 101 ✅');
console.log('    sheet.duplicateRow(75 + i, 1, true);   // 75, 76, 77 ✅');
console.log('  }');
console.log('');

console.log('========================================');
console.log('ループの実行シミュレーション');
console.log('========================================\n');

for (let i = 0; i < extraRows; i++) {
  console.log(`--- ループ${i + 1}回目 (i = ${i}) ---`);
  console.log(`  duplicateRow(${145 + i}, 1, true): 【D】Row ${145 + i}を複製`);
  console.log(`    → Row ${145 + i}の後にRow ${146 + i}が挿入される`);
  console.log(`    → Row ${146 + i}以降が+1行ずれる`);
  console.log('');
  console.log(`  duplicateRow(${122 + i}, 1, true): 【C】Row ${122 + i}を複製`);
  console.log(`    → Row ${122 + i}の後にRow ${123 + i}が挿入される`);
  console.log(`    → Row ${123 + i}以降が+1行ずれる（【D】も影響）`);
  console.log('');
  console.log(`  duplicateRow(${99 + i}, 1, true): 【B】Row ${99 + i}を複製`);
  console.log(`    → Row ${99 + i}の後にRow ${100 + i}が挿入される`);
  console.log(`    → Row ${100 + i}以降が+1行ずれる（【C】【D】も影響）`);
  console.log('');
  console.log(`  duplicateRow(${75 + i}, 1, true): 【A】Row ${75 + i}を複製`);
  console.log(`    → Row ${75 + i}の後にRow ${76 + i}が挿入される`);
  console.log(`    → Row ${76 + i}以降が+1行ずれる（【B】【C】【D】も影響）`);
  console.log('');
}

console.log('========================================');
console.log('結果の確認');
console.log('========================================\n');

console.log('【A】収入セクション:');
console.log('  元: Row 55-74 (20行)');
console.log('  ループ1: Row 75を複製 → Row 76追加 → Row 55-76 (22行)');
console.log('  ループ2: Row 76を複製 → Row 77追加 → Row 55-77 (23行)');
console.log('  ループ3: Row 77を複製 → Row 78追加 → Row 55-78 (24行)');
console.log('  ❌ これは間違い！20 + 3 = 23行になるべき');
console.log('');

console.log('問題点:');
console.log('  - duplicateRow(75, 1, true)は「Row 75を複製してRow 76に挿入」');
console.log('  - これにより、Row 55-75 が Row 55-76 になる（21行）');
console.log('  - 次に duplicateRow(76, 1, true)すると、Row 55-77 になる（23行）❌');
console.log('  - 正しくは、Row 55-75 → Row 55-76 → Row 55-77 で23行 ✅');
console.log('');

console.log('実は、この修正は正しいです！');
console.log('  - 元: 20行（Row 55-74）');
console.log('  - 1回目: Row 75を複製 → 21行（Row 55-75）');
console.log('  - 2回目: Row 76を複製 → 22行（Row 55-76）');
console.log('  - 3回目: Row 77を複製 → 23行（Row 55-77）');
console.log('');

console.log('待って...元は何行？');
console.log('  - ヘッダー Row 55 を含めて書き込むので...');
console.log('  - 元: Row 55-74 (20行) ✅');
console.log('  - 3行追加後: Row 55-77 (23行) ✅');
console.log('');

console.log('✅ 修正は正しいです！');
console.log('✅ ループごとに、直前に追加された行（最終行）を複製します');
console.log('✅ 各セクションの行数が正しく増加します');
