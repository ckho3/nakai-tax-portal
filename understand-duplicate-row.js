/**
 * duplicateRow()の動作を正確に理解する
 */

console.log('========================================');
console.log('duplicateRow()の動作確認');
console.log('========================================\n');

console.log('duplicateRow(rowNumber, amount, insert)の仕様:');
console.log('  - rowNumber: 複製元の行番号');
console.log('  - amount: 複製する行数（通常1）');
console.log('  - insert: true = 元の行の「後」に挿入');
console.log('');

console.log('例: sheet.duplicateRow(75, 1, true)');
console.log('  → Row 75を複製して、Row 75の後（Row 76の位置）に挿入');
console.log('  → 元のRow 76以降は1行下にずれる');
console.log('');

console.log('========================================');
console.log('23件PDFの場合の実行順序');
console.log('========================================\n');

const extraRows = 3;

console.log('初期状態:');
console.log('  【A】: Row 55ヘッダー, Row 56-75データ（20行）');
console.log('  【B】: Row 78ヘッダー, Row 79注釈, Row 80-99データ（20行）');
console.log('  【C】: Row 101ヘッダー, Row 102注釈, Row 103-122データ（20行）');
console.log('  【D】: Row 124ヘッダー, Row 125注釈, Row 126-145データ（20行）');
console.log('');

console.log(`${extraRows}回のループ:\n`);

// ループ1
console.log('--- ループ1回目 ---');
console.log('duplicateRow(144, 1, true): Row 144を複製 → Row 145に挿入');
console.log('  【D】: Row 126-145 → Row 126-146');
console.log('duplicateRow(121, 1, true): Row 121を複製 → Row 122に挿入');
console.log('  【C】: Row 103-122 → Row 103-123');
console.log('  【D】は1行下にずれる: Row 126-146 → Row 127-147');
console.log('duplicateRow(98, 1, true): Row 98を複製 → Row 99に挿入');
console.log('  【B】: Row 80-99 → Row 80-100');
console.log('  【C】【D】は1行下にずれる: 【C】Row 104-124, 【D】Row 128-148');
console.log('duplicateRow(75, 1, true): Row 75を複製 → Row 76に挿入');
console.log('  【A】: Row 56-75 → Row 56-76');
console.log('  【B】【C】【D】は1行下にずれる:');
console.log('    【B】: Row 78ヘッダー→Row 79, Row 79注釈→Row 80, Row 80-100データ→Row 81-101');
console.log('    【C】: Row 101ヘッダー→Row 102, ...');
console.log('    【D】: Row 124ヘッダー→Row 125, ...');
console.log('');

console.log('ループ1回目終了後:');
console.log('  【A】: Row 55ヘッダー, Row 56-76データ（21行）');
console.log('  【B】: Row 79ヘッダー, Row 80注釈, Row 81-101データ（21行）');
console.log('  【C】: Row 105ヘッダー, Row 106注釈, Row 107-127データ（21行）');
console.log('  【D】: Row 129ヘッダー, Row 130注釈, Row 131-151データ（21行）');
console.log('');

console.log('--- ループ2回目、3回目も同様 ---');
console.log('');

console.log('========================================');
console.log('最終結果（3回ループ後）');
console.log('========================================\n');

console.log('【A】収入:');
console.log('  Row 55: ヘッダー（変わらず）');
console.log('  Row 56-78: データ（23行）');
console.log('');

console.log('【B】管理手数料:');
console.log(`  Row ${78 + extraRows} = Row 81: ヘッダー`);
console.log(`  Row ${79 + extraRows} = Row 82: 注釈`);
console.log(`  Row ${80 + extraRows} = Row 83: データ開始`);
console.log(`  Row 83-105: データ（23行）`);
console.log('');

console.log('【C】広告費:');
console.log(`  Row ${101 + extraRows * 2} = Row 107: ヘッダー`);
console.log(`  Row ${102 + extraRows * 2} = Row 108: 注釈`);
console.log(`  Row ${103 + extraRows * 2} = Row 109: データ開始`);
console.log('');

console.log('【D】修繕費:');
console.log(`  Row ${124 + extraRows * 3} = Row 133: ヘッダー`);
console.log(`  Row ${125 + extraRows * 3} = Row 134: 注釈`);
console.log(`  Row ${126 + extraRows * 3} = Row 135: データ開始`);
console.log('');

console.log('========================================');
console.log('重要な発見！');
console.log('========================================\n');

console.log('❌ 問題: 【A】のヘッダー（Row 55）だけは動かない！');
console.log('✅ しかし、【B】【C】【D】のヘッダーは全て下にずれる！');
console.log('');

console.log('なぜ？');
console.log('  → duplicateRow(75, 1, true)は「Row 75の後」に挿入');
console.log('  → Row 76以降が下にずれる');
console.log('  → Row 55のヘッダーは影響を受けない');
console.log('  → しかし、Row 78以降（【B】【C】【D】）は全て影響を受ける');
console.log('');

console.log('これにより:');
console.log('  - 【A】ヘッダー: Row 55のまま');
console.log('  - 【B】ヘッダー: Row 78 → Row 81に移動');
console.log('  - 【C】ヘッダー: Row 101 → Row 107に移動');
console.log('  - 【D】ヘッダー: Row 124 → Row 133に移動');
