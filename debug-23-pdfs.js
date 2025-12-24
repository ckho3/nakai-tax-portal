/**
 * 23件PDFの場合の詳細確認
 */

console.log('========================================');
console.log('23件PDFの場合の動作確認');
console.log('========================================\n');

const pdfCount = 23;
const extraRows = pdfCount - 20;  // 3行

console.log(`PDFファイル数: ${pdfCount}件`);
console.log(`追加する行数: ${extraRows}行\n`);

console.log('【A】収入セクション:');
console.log('  元の範囲: Row 56-75 (20行)');
console.log(`  行追加後: Row 56-${75 + extraRows} (${20 + extraRows}行)`);
console.log('  → 収入合計①は Row 56-78 に書き込まれる');
console.log('');

console.log('【B】管理手数料セクションの構造:');
console.log('  Row 78: ヘッダー【B】支払手数料（管理手数料等）');
console.log('  Row 79: 注釈（→サブリースの場合、法人負担）');
console.log('  Row 80: データ開始行（元々）');
console.log('');

console.log('問題:');
console.log('  Row 56-78に収入合計①を書き込む');
console.log('  → Row 78はヘッダー行なので、収入合計①と重なってしまう！❌');
console.log('');

console.log('========================================');
console.log('正しい動作');
console.log('========================================\n');

console.log('【A】収入セクション:');
console.log('  Row 55: ヘッダー');
console.log('  Row 56-78: データ行（23件）');
console.log('  → 最後の物件（23件目）はRow 78に書き込まれる');
console.log('');

console.log('【B】管理手数料セクション:');
console.log('  元の構造: Row 78ヘッダー、Row 79注釈、Row 80-99データ');
console.log(`  【A】に${extraRows}行追加したので、【B】全体が${extraRows}行下にずれる`);
console.log(`  ヘッダー: Row ${78 + extraRows}`);
console.log(`  注釈: Row ${79 + extraRows}`);
console.log(`  データ開始: Row ${80 + extraRows}`);
console.log('');

console.log('つまり:');
console.log(`  【B】のデータ開始行は Row ${80 + extraRows} です`);
console.log('');

console.log('========================================');
console.log('現在のコードの計算');
console.log('========================================\n');

const baseRow = 56;  // 物件1
const currentFormula = (80 + extraRows) + (baseRow - 56);
console.log(`managementBaseRow = (80 + ${extraRows}) + (${baseRow} - 56) = ${currentFormula}`);
console.log('');

console.log('物件1-3の書き込み位置:');
for (let i = 0; i < 3; i++) {
  const base = 56 + i;
  const aRow = base;
  const bRow = (80 + extraRows) + (base - 56);
  console.log(`  物件${i + 1}: 【A】Row ${aRow}, 【B】Row ${bRow}`);
}
console.log('');

console.log('✅ 計算式は正しいです！');
console.log(`✅ 【A】の最後の物件（23件目）はRow ${56 + 22} = Row 78`);
console.log(`✅ 【B】のデータ開始はRow ${80 + extraRows} = Row 83`);
console.log(`✅ Row 78（【A】の最後）とRow 83（【B】の開始）は重ならない`);
console.log('');

console.log('========================================');
console.log('各セクションのヘッダーとデータの位置関係');
console.log('========================================\n');

console.log(`【A】収入:`);
console.log(`  Row 55: ヘッダー`);
console.log(`  Row 56-78: データ（23件）`);
console.log('');

console.log(`【B】管理手数料:`);
console.log(`  Row ${78 + extraRows}: ヘッダー`);
console.log(`  Row ${79 + extraRows}: 注釈`);
console.log(`  Row ${80 + extraRows}-${80 + extraRows + 22}: データ（23件）`);
console.log('');

console.log(`【C】広告費:`);
console.log(`  Row ${101 + extraRows * 2}: ヘッダー`);
console.log(`  Row ${102 + extraRows * 2}: 注釈`);
console.log(`  Row ${103 + extraRows * 2}-${103 + extraRows * 2 + 22}: データ（23件）`);
console.log('');

console.log(`【D】修繕費:`);
console.log(`  Row ${124 + extraRows * 3}: ヘッダー`);
console.log(`  Row ${125 + extraRows * 3}: 注釈`);
console.log(`  Row ${126 + extraRows * 3}-${126 + extraRows * 3 + 22}: データ（23件）`);
