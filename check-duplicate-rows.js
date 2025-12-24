/**
 * 現在のduplicateRow()の行番号が正しいか確認
 */

console.log('========================================');
console.log('duplicateRow()で使用している行番号の確認');
console.log('========================================\n');

// Excelの実際の構造
const SECTION_A_HEADER = 55;  // 【A】収入ヘッダー
const SECTION_B_HEADER = 78;  // 【B】管理手数料ヘッダー
const SECTION_C_HEADER = 101; // 【C】広告費ヘッダー
const SECTION_D_HEADER = 124; // 【D】修繕費ヘッダー

console.log('Excelの各セクションのヘッダー行:');
console.log(`  【A】収入: Row ${SECTION_A_HEADER}`);
console.log(`  【B】管理手数料: Row ${SECTION_B_HEADER}`);
console.log(`  【C】広告費: Row ${SECTION_C_HEADER}`);
console.log(`  【D】修繕費: Row ${SECTION_D_HEADER}`);
console.log('');

// 各セクションのデータ行範囲（初期状態：20行分）
console.log('各セクションの初期データ行範囲（20行分）:');
console.log(`  【A】収入: Row ${SECTION_A_HEADER + 1} 〜 Row ${SECTION_A_HEADER + 20} (Row 56-75)`);
console.log(`  【B】管理手数料: Row ${SECTION_B_HEADER + 1} 〜 Row ${SECTION_B_HEADER + 20} (Row 79-98)`);
console.log(`  【C】広告費: Row ${SECTION_C_HEADER + 1} 〜 Row ${SECTION_C_HEADER + 20} (Row 102-121)`);
console.log(`  【D】修繕費: Row ${SECTION_D_HEADER + 1} 〜 Row ${SECTION_D_HEADER + 20} (Row 125-144)`);
console.log('');

console.log('========================================');
console.log('現在のコードで使用している行番号');
console.log('========================================\n');

const currentDuplicateRows = {
  A: 74,
  B: 97,
  C: 120,
  D: 143
};

console.log('現在のduplicateRow()の行番号:');
console.log(`  【A】収入: Row ${currentDuplicateRows.A}`);
console.log(`  【B】管理手数料: Row ${currentDuplicateRows.B}`);
console.log(`  【C】広告費: Row ${currentDuplicateRows.C}`);
console.log(`  【D】修繕費: Row ${currentDuplicateRows.D}`);
console.log('');

console.log('========================================');
console.log('正しい行番号（最終データ行）を計算');
console.log('========================================\n');

const correctDuplicateRows = {
  A: SECTION_A_HEADER + 20,  // 55 + 20 = 75
  B: SECTION_B_HEADER + 20,  // 78 + 20 = 98
  C: SECTION_C_HEADER + 20,  // 101 + 20 = 121
  D: SECTION_D_HEADER + 20   // 124 + 20 = 144
};

console.log('正しいduplicateRow()の行番号（20行目）:');
console.log(`  【A】収入: Row ${correctDuplicateRows.A}`);
console.log(`  【B】管理手数料: Row ${correctDuplicateRows.B}`);
console.log(`  【C】広告費: Row ${correctDuplicateRows.C}`);
console.log(`  【D】修繕費: Row ${correctDuplicateRows.D}`);
console.log('');

console.log('========================================');
console.log('比較結果');
console.log('========================================\n');

for (const section of ['A', 'B', 'C', 'D']) {
  const current = currentDuplicateRows[section];
  const correct = correctDuplicateRows[section];
  const diff = current - correct;

  console.log(`【${section}】セクション:`);
  console.log(`  現在の値: Row ${current}`);
  console.log(`  正しい値: Row ${correct}`);
  console.log(`  差分: ${diff}行 ${diff === 0 ? '✅ 正しい' : `❌ ${Math.abs(diff)}行${diff > 0 ? '下' : '上'}にずれている`}`);
  console.log('');
}

console.log('========================================');
console.log('修正が必要な理由');
console.log('========================================\n');

console.log('duplicateRow(rowNumber, amount, insert)の動作:');
console.log('  - rowNumber: 複製する元の行番号');
console.log('  - amount: 複製する行数');
console.log('  - insert: true = 元の行の後に挿入');
console.log('');
console.log('各セクションの最終データ行（20行目）を複製することで、');
console.log('21行目以降を追加する必要があります。');
console.log('');
console.log('現在の行番号は1行ずれているため、修正が必要です。');
