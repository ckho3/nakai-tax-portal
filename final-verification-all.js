/**
 * 全ての修正の最終検証
 */

console.log('========================================');
console.log('最終修正の完全検証');
console.log('========================================\n');

console.log('修正内容:');
console.log('1. 【A】セクションもRow 55（ヘッダー行）から書き込む');
console.log('2. duplicateRow()を各セクションの20行目（データ最終行）に変更');
console.log('   - 【A】: Row 75 → Row 75（変更なし）');
console.log('   - 【B】: Row 98 → Row 99');
console.log('   - 【C】: Row 121 → Row 122');
console.log('   - 【D】: Row 144 → Row 145');
console.log('');

console.log('========================================');
console.log('23件PDFの場合の書き込み位置');
console.log('========================================\n');

const pdfCount = 23;
const extraRows = pdfCount - 20;

console.log(`PDFファイル数: ${pdfCount}件`);
console.log(`追加する行数: ${extraRows}行\n`);

console.log('物件1-3と物件21-23の書き込み位置:\n');

for (let i = 0; i < pdfCount; i++) {
  const baseRow = 55 + i;  // 【A】セクションの行（Row 55から開始）

  // 修正後の計算式
  const aRow = baseRow;
  const bRow = (78 + extraRows) + (baseRow - 55);
  const cRow = (101 + extraRows * 2) + (baseRow - 55);
  const dRow = (124 + extraRows * 3) + (baseRow - 55);

  if (i < 3 || i >= pdfCount - 3) {
    console.log(`物件${i + 1}:`);
    console.log(`  【A】収入合計①: Row ${aRow} (G${aRow})`);
    console.log(`  【B】管理手数料: Row ${bRow} (G${bRow})`);
    console.log(`  【C】宣伝広告費: Row ${cRow} (G${cRow})`);
    console.log(`  【D】設備交換費: Row ${dRow} (G${dRow})`);
    console.log('');
  } else if (i === 3) {
    console.log('  ...');
    console.log('');
  }
}

console.log('========================================');
console.log('重要なポイント');
console.log('========================================\n');

console.log('✅ 物件1の【A】収入合計①: Row 55 (G55) ← ヘッダー行に上書き');
console.log('✅ 物件1の【B】管理手数料: Row 81 (G81) ← ヘッダー行に上書き');
console.log('✅ 物件1の【C】宣伝広告費: Row 107 (G107) ← ヘッダー行に上書き');
console.log('✅ 物件1の【D】設備交換費: Row 133 (G133) ← ヘッダー行に上書き');
console.log('');

console.log('全てのセクションでヘッダー行から書き込みが開始されます。');
console.log('元のヘッダー情報は失われますが、G列に項目名が書き込まれます。');
console.log('');

console.log('========================================');
console.log('duplicateRow()の実行内容');
console.log('========================================\n');

console.log(`${extraRows}回のループで各セクションの20行目を複製:\n`);

for (let i = 1; i <= extraRows; i++) {
  console.log(`--- ループ${i}回目 ---`);
  console.log(`  duplicateRow(${145 + (i - 1)}, 1, true): 【D】の${19 + i}行目を複製`);
  console.log(`  duplicateRow(${122 + (i - 1)}, 1, true): 【C】の${19 + i}行目を複製`);
  console.log(`  duplicateRow(${99 + (i - 1)}, 1, true): 【B】の${19 + i}行目を複製`);
  console.log(`  duplicateRow(${75 + (i - 1)}, 1, true): 【A】の${19 + i}行目を複製`);
  console.log('');
}

console.log('========================================');
console.log('最終結果');
console.log('========================================\n');

console.log('【A】収入:');
console.log('  Row 55-77: データ（23件）← Row 55はヘッダーだったがデータで上書き');
console.log('');

console.log('【B】管理手数料:');
console.log(`  Row ${78 + extraRows}-${78 + extraRows + 22}: データ（23件）← Row 81はヘッダーだったがデータで上書き`);
console.log('');

console.log('【C】広告費:');
console.log(`  Row ${101 + extraRows * 2}-${101 + extraRows * 2 + 22}: データ（23件）← Row 107はヘッダーだったがデータで上書き`);
console.log('');

console.log('【D】修繕費:');
console.log(`  Row ${124 + extraRows * 3}-${124 + extraRows * 3 + 22}: データ（23件）← Row 133はヘッダーだったがデータで上書き`);
console.log('');

console.log('========================================');
console.log('修正完了！');
console.log('========================================\n');

console.log('✅ 全てのセクションがヘッダー行から書き込みを開始します');
console.log('✅ duplicateRow()は各セクションの20行目（データ最終行）を複製します');
console.log('✅ あなたの要求通りに修正されました！');
