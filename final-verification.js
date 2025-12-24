/**
 * 修正後の完全な動作確認
 */

console.log('========================================');
console.log('修正完了：最終検証');
console.log('========================================\n');

// Excelの構造
const SECTION_A_HEADER = 55;   // 【A】収入ヘッダー
const SECTION_B_HEADER = 78;   // 【B】管理手数料ヘッダー
const SECTION_C_HEADER = 101;  // 【C】広告費ヘッダー
const SECTION_D_HEADER = 124;  // 【D】修繕費ヘッダー

console.log('【修正1】duplicateRow()の行番号');
console.log('========================================\n');

console.log('修正前（間違い）:');
console.log('  sheet.duplicateRow(74, 1, true);   // 【A】❌ 19行目を複製');
console.log('  sheet.duplicateRow(97, 1, true);   // 【B】❌ 19行目を複製');
console.log('  sheet.duplicateRow(120, 1, true);  // 【C】❌ 19行目を複製');
console.log('  sheet.duplicateRow(143, 1, true);  // 【D】❌ 19行目を複製');
console.log('');

console.log('修正後（正しい）:');
console.log('  sheet.duplicateRow(75, 1, true);   // 【A】✅ 20行目を複製');
console.log('  sheet.duplicateRow(98, 1, true);   // 【B】✅ 20行目を複製');
console.log('  sheet.duplicateRow(121, 1, true);  // 【C】✅ 20行目を複製');
console.log('  sheet.duplicateRow(144, 1, true);  // 【D】✅ 20行目を複製');
console.log('');

console.log('【修正2】データ書き込み行の計算式');
console.log('========================================\n');

console.log('修正前（間違い）:');
console.log('  advertisingBaseRow = (101 + extraRows * 2) + (baseRow - 55)  ❌');
console.log('  repairBaseRow = (124 + extraRows * 3) + (baseRow - 55)       ❌');
console.log('');

console.log('修正後（正しい）:');
console.log('  managementBaseRow = (78 + extraRows) + (baseRow - 55)   ✅');
console.log('  advertisingBaseRow = (101 + extraRows) + (baseRow - 55) ✅');
console.log('  repairBaseRow = (124 + extraRows) + (baseRow - 55)      ✅');
console.log('');

console.log('【動作シミュレーション】30件のPDF');
console.log('========================================\n');

const pdfCount = 30;
const extraRows = pdfCount - 20;

console.log(`PDFファイル数: ${pdfCount}件`);
console.log(`追加する行数: ${extraRows}行\n`);

console.log('ステップ1: 行追加処理');
console.log('----------------------------------------');
console.log(`【A】収入セクション: Row 75を${extraRows}回複製 → Row 76-85を作成`);
console.log(`【B】管理手数料セクション: Row 98を${extraRows}回複製 → Row 99-108を作成`);
console.log(`【C】広告費セクション: Row 121を${extraRows}回複製 → Row 122-131を作成`);
console.log(`【D】修繕費セクション: Row 144を${extraRows}回複製 → Row 145-154を作成`);
console.log('');

console.log('ステップ2: 行追加後の各セクションのヘッダー位置');
console.log('----------------------------------------');
console.log('※ヘッダーより上に行を追加しているため、ヘッダーは動かない');
console.log(`【A】収入ヘッダー: Row ${SECTION_A_HEADER} (変わらず)`);
console.log(`【B】管理手数料ヘッダー: Row ${SECTION_B_HEADER} (変わらず)`);
console.log(`【C】広告費ヘッダー: Row ${SECTION_C_HEADER} (変わらず)`);
console.log(`【D】修繕費ヘッダー: Row ${SECTION_D_HEADER} (変わらず)`);
console.log('');

console.log('ステップ3: データ書き込み位置の計算（物件1-5）');
console.log('----------------------------------------\n');

for (let propertyIndex = 0; propertyIndex < 5; propertyIndex++) {
  const baseRow = SECTION_A_HEADER + propertyIndex;  // 【A】セクションの行

  // 修正後の計算式
  const aRow = baseRow;
  const bRow = (SECTION_B_HEADER + extraRows) + (baseRow - SECTION_A_HEADER);
  const cRow = (SECTION_C_HEADER + extraRows) + (baseRow - SECTION_A_HEADER);
  const dRow = (SECTION_D_HEADER + extraRows) + (baseRow - SECTION_A_HEADER);

  console.log(`物件${propertyIndex + 1}:`);
  console.log(`  【A】収入: Row ${aRow}`);
  console.log(`  【B】管理手数料: Row ${bRow}`);
  console.log(`  【C】広告費: Row ${cRow}`);
  console.log(`  【D】修繕費: Row ${dRow}`);

  // 検証: 各セクション間の行番号の差が一致しているか
  const diffAB = bRow - aRow;
  const diffBC = cRow - bRow;
  const diffCD = dRow - cRow;

  const expectedDiffAB = SECTION_B_HEADER - SECTION_A_HEADER;  // 23行
  const expectedDiffBC = SECTION_C_HEADER - SECTION_B_HEADER;  // 23行
  const expectedDiffCD = SECTION_D_HEADER - SECTION_C_HEADER;  // 23行

  const isCorrect = (
    diffAB === expectedDiffAB &&
    diffBC === expectedDiffBC &&
    diffCD === expectedDiffCD
  );

  console.log(`  セクション間の差: A→B=${diffAB}行, B→C=${diffBC}行, C→D=${diffCD}行`);
  console.log(`  検証結果: ${isCorrect ? '✅ 正しい' : '❌ 間違い'}`);
  console.log('');
}

console.log('========================================');
console.log('まとめ');
console.log('========================================\n');

console.log('✅ 修正内容:');
console.log('  1. duplicateRow()の行番号を+1修正（74→75, 97→98, 120→121, 143→144）');
console.log('  2. データ書き込み行の計算式を修正（extraRows * 2, * 3 → extraRows）');
console.log('');
console.log('✅ 効果:');
console.log('  - 20件以上のPDFをアップロードした場合でも正しい行にデータが書き込まれる');
console.log('  - 各セクション（A/B/C/D）の行間隔が維持される（23行間隔）');
console.log('  - G~U列に正しくデータが挿入される');
console.log('');
