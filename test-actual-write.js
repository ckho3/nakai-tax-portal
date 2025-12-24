/**
 * 実際の書き込み位置をテスト
 */

console.log('========================================');
console.log('23件PDFの実際の書き込み位置');
console.log('========================================\n');

const pdfCount = 23;
const extraRows = pdfCount - 20;

console.log(`PDFファイル数: ${pdfCount}件`);
console.log(`追加する行数: ${extraRows}行\n`);

console.log('全23件の物件の書き込み位置:\n');

for (let i = 0; i < pdfCount; i++) {
  const baseRow = 56 + i;  // 【A】セクションの行

  // 現在のコードの計算式
  const aRow = baseRow;
  const bRow = (80 + extraRows) + (baseRow - 56);
  const cRow = (103 + extraRows * 2) + (baseRow - 56);
  const dRow = (126 + extraRows * 3) + (baseRow - 56);

  if (i < 3 || i >= pdfCount - 3) {
    console.log(`物件${i + 1}:`);
    console.log(`  【A】収入合計①: Row ${aRow} (G${aRow})`);
    console.log(`  【B】管理手数料: Row ${bRow} (G${bRow})`);
    console.log(`  【C】広告費: Row ${cRow} (G${cRow})`);
    console.log(`  【D】修繕費: Row ${dRow} (G${dRow})`);
    console.log('');
  } else if (i === 3) {
    console.log('  ...');
    console.log('');
  }
}

console.log('========================================');
console.log('セクションのヘッダー位置（行追加後）');
console.log('========================================\n');

console.log('【A】収入:');
console.log('  Row 55: ヘッダー');
console.log('  Row 56-78: データ（23件）');
console.log('');

console.log('【B】管理手数料:');
console.log(`  Row ${78 + extraRows} = Row 81: ヘッダー【B】支払手数料...`);
console.log(`  Row ${79 + extraRows} = Row 82: 注釈`);
console.log(`  Row ${80 + extraRows}-${80 + extraRows + 22}: データ（Row 83-105）`);
console.log('');

console.log('【C】広告費:');
console.log(`  Row ${101 + extraRows * 2} = Row 107: ヘッダー`);
console.log(`  Row ${102 + extraRows * 2} = Row 108: 注釈`);
console.log(`  Row ${103 + extraRows * 2}-${103 + extraRows * 2 + 22}: データ（Row 109-131）`);
console.log('');

console.log('【D】修繕費:');
console.log(`  Row ${124 + extraRows * 3} = Row 133: ヘッダー`);
console.log(`  Row ${125 + extraRows * 3} = Row 134: 注釈`);
console.log(`  Row ${126 + extraRows * 3}-${126 + extraRows * 3 + 22}: データ（Row 135-157）`);
console.log('');

console.log('========================================');
console.log('検証');
console.log('========================================\n');

// 【A】の最後とヘッダーの確認
const aLast = 56 + (pdfCount - 1);  // Row 78
const bHeader = 78 + extraRows;      // Row 81

console.log(`【A】の最後の物件（23件目）: Row ${aLast} (G${aLast})`);
console.log(`【B】のヘッダー: Row ${bHeader} (C${bHeader})`);
console.log('');

if (aLast < bHeader) {
  console.log(`✅ 【A】の最後（Row ${aLast}）と【B】のヘッダー（Row ${bHeader}）は重ならない`);
} else {
  console.log(`❌ 【A】の最後（Row ${aLast}）が【B】のヘッダー（Row ${bHeader}）と重なる！`);
}
console.log('');

// 【B】のデータ開始
const bDataStart = 80 + extraRows;  // Row 83
console.log(`【B】のデータ開始: Row ${bDataStart} (G${bDataStart})`);
console.log('');

// 物件1の【B】書き込み位置
const物件1bRow = (80 + extraRows) + (56 - 56);
console.log(`物件1の【B】書き込み位置: Row ${物件1bRow} (G${物件1bRow})`);
console.log('');

if (物件1bRow === bDataStart) {
  console.log(`✅ 物件1の【B】はデータ開始行（Row ${bDataStart}）に正しく書き込まれる`);
} else {
  console.log(`❌ 物件1の【B】の書き込み位置が間違っている`);
}
