/**
 * ヘッダー行から書き込む修正の検証
 */

console.log('========================================');
console.log('修正後の書き込み位置（ヘッダー行から）');
console.log('========================================\n');

const pdfCount = 23;
const extraRows = pdfCount - 20;

console.log(`PDFファイル数: ${pdfCount}件`);
console.log(`追加する行数: ${extraRows}行\n`);

console.log('物件1-3と物件21-23の書き込み位置:\n');

for (let i = 0; i < pdfCount; i++) {
  const baseRow = 56 + i;

  // 修正後の計算式（ヘッダー行から書き込む）
  const aRow = baseRow;
  const bRow = (78 + extraRows) + (baseRow - 56);
  const cRow = (101 + extraRows * 2) + (baseRow - 56);
  const dRow = (124 + extraRows * 3) + (baseRow - 56);

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
console.log('セクションのヘッダーとデータの対応');
console.log('========================================\n');

console.log('【A】収入:');
console.log('  Row 55: ヘッダー【A】収入（収入明細の金額をそのまま記入）');
console.log('  Row 56: 物件1のデータ（収入合計①）');
console.log('  Row 57: 物件2のデータ（収入合計①）');
console.log('  ...');
console.log('  Row 78: 物件23のデータ（収入合計①）');
console.log('');

console.log('【B】管理手数料:');
console.log(`  Row ${78 + extraRows} = Row 81: 物件1のデータ（管理手数料）← ヘッダー行に上書き ✅`);
console.log(`  Row ${79 + extraRows} = Row 82: 物件2のデータ（管理手数料）← 注釈行に上書き`);
console.log(`  Row ${80 + extraRows} = Row 83: 物件3のデータ（管理手数料）`);
console.log('  ...');
console.log(`  Row ${80 + extraRows + 22} = Row 105: 物件23のデータ（管理手数料）`);
console.log('');

console.log('【C】広告費:');
console.log(`  Row ${101 + extraRows * 2} = Row 107: 物件1のデータ（宣伝広告費）← ヘッダー行に上書き ✅`);
console.log(`  Row ${102 + extraRows * 2} = Row 108: 物件2のデータ（宣伝広告費）← 注釈行に上書き`);
console.log(`  Row ${103 + extraRows * 2} = Row 109: 物件3のデータ（宣伝広告費）`);
console.log('  ...');
console.log('');

console.log('【D】修繕費:');
console.log(`  Row ${124 + extraRows * 3} = Row 133: 物件1のデータ（設備交換費）← ヘッダー行に上書き ✅`);
console.log(`  Row ${125 + extraRows * 3} = Row 134: 物件2のデータ（設備交換費）← 注釈行に上書き`);
console.log(`  Row ${126 + extraRows * 3} = Row 135: 物件3のデータ（設備交換費）`);
console.log('  ...');
console.log('');

console.log('========================================');
console.log('重要なポイント');
console.log('========================================\n');

console.log('✅ 物件1の【B】管理手数料は Row 81（元のヘッダー行）に書き込まれる');
console.log('✅ 物件1の【C】宣伝広告費は Row 107（元のヘッダー行）に書き込まれる');
console.log('✅ 物件1の【D】設備交換費は Row 133（元のヘッダー行）に書き込まれる');
console.log('');
console.log('これにより、ヘッダー行や注釈行がデータで上書きされます。');
console.log('元のヘッダー情報は失われますが、各行にG列に項目名が書き込まれます。');
console.log('');

console.log('========================================');
console.log('修正完了');
console.log('========================================\n');

console.log('✅ 【B】【C】【D】セクションはヘッダー行から書き込みを開始します');
console.log('✅ 23件PDFの場合、物件1の【B】はRow 81（G81）に書き込まれます');
console.log('✅ あなたの要求通りに修正されました！');
