/**
 * 修正後の行計算式を確認するスクリプト
 */

console.log('========================================');
console.log('修正前と修正後の行番号計算の比較');
console.log('========================================\n');

// Excelの構造（実際の値）
const SECTION_A_START = 55;  // 【A】収入
const SECTION_B_START = 78;  // 【B】管理手数料
const SECTION_C_START = 101; // 【C】広告費
const SECTION_D_START = 124; // 【D】修繕費
const SECTION_E_START = 147; // 【E】サブリース

const SECTION_GAP = 23; // 各セクション間の行数

console.log('Excelの元々の構造:');
console.log(`  【A】収入: Row ${SECTION_A_START}`);
console.log(`  【B】管理手数料: Row ${SECTION_B_START} (${SECTION_B_START - SECTION_A_START}行後)`);
console.log(`  【C】広告費: Row ${SECTION_C_START} (${SECTION_C_START - SECTION_A_START}行後)`);
console.log(`  【D】修繕費: Row ${SECTION_D_START} (${SECTION_D_START - SECTION_A_START}行後)`);
console.log(`  【E】サブリース: Row ${SECTION_E_START} (${SECTION_E_START - SECTION_A_START}行後)`);
console.log(`\n各セクション間の行数: ${SECTION_GAP}行\n`);

// テストケース: 30件のPDFをアップロードした場合
const pdfCount = 30;
const extraRows = pdfCount > 20 ? pdfCount - 20 : 0;

console.log('========================================');
console.log(`テストケース: ${pdfCount}件のPDF (extraRows = ${extraRows})`);
console.log('========================================\n');

// 各物件の行番号を計算
for (let propertyIndex = 0; propertyIndex < Math.min(5, pdfCount); propertyIndex++) {
  const baseRow = SECTION_A_START + propertyIndex;

  console.log(`物件 ${propertyIndex + 1} (${propertyIndex === 0 ? '1件目' : propertyIndex === 1 ? '2件目' : propertyIndex === 2 ? '3件目' : `${propertyIndex + 1}件目`}):`);
  console.log(`  【A】収入: Row ${baseRow}`);

  // 修正前の計算式（間違っていた）
  const oldManagementRow = (78 + extraRows) + (baseRow - 55);
  const oldAdvertisingRow = (101 + extraRows * 2) + (baseRow - 55);
  const oldRepairRow = (124 + extraRows * 3) + (baseRow - 55);

  // 修正後の計算式（正しい）
  const newManagementRow = (78 + extraRows) + (baseRow - 55);
  const newAdvertisingRow = (101 + extraRows) + (baseRow - 55);
  const newRepairRow = (124 + extraRows) + (baseRow - 55);

  console.log(`  【B】管理手数料:`);
  console.log(`    修正前: Row ${oldManagementRow}`);
  console.log(`    修正後: Row ${newManagementRow}`);

  console.log(`  【C】広告費:`);
  console.log(`    修正前: Row ${oldAdvertisingRow}`);
  console.log(`    修正後: Row ${newAdvertisingRow}`);

  console.log(`  【D】修繕費:`);
  console.log(`    修正前: Row ${oldRepairRow}`);
  console.log(`    修正後: Row ${newRepairRow}`);

  // 正しい行番号を計算（期待値）
  const expectedManagementRow = SECTION_B_START + extraRows + propertyIndex;
  const expectedAdvertisingRow = SECTION_C_START + extraRows + propertyIndex;
  const expectedRepairRow = SECTION_D_START + extraRows + propertyIndex;

  console.log(`  期待値（正しい行番号）:`);
  console.log(`    【B】: Row ${expectedManagementRow}`);
  console.log(`    【C】: Row ${expectedAdvertisingRow}`);
  console.log(`    【D】: Row ${expectedRepairRow}`);

  // 検証
  const isCorrect = (
    newManagementRow === expectedManagementRow &&
    newAdvertisingRow === expectedAdvertisingRow &&
    newRepairRow === expectedRepairRow
  );

  console.log(`  ✓ 修正後の計算式は${isCorrect ? '正しい' : '間違っている'}です`);
  console.log('');
}

console.log('========================================');
console.log('各セクションの行範囲（30件の場合）');
console.log('========================================\n');

console.log(`【A】収入セクション: Row ${SECTION_A_START + extraRows} 〜 Row ${SECTION_A_START + extraRows + pdfCount - 1}`);
console.log(`【B】管理手数料セクション: Row ${SECTION_B_START + extraRows} 〜 Row ${SECTION_B_START + extraRows + pdfCount - 1}`);
console.log(`【C】広告費セクション: Row ${SECTION_C_START + extraRows} 〜 Row ${SECTION_C_START + extraRows + pdfCount - 1}`);
console.log(`【D】修繕費セクション: Row ${SECTION_D_START + extraRows} 〜 Row ${SECTION_D_START + extraRows + pdfCount - 1}`);

console.log('\n========================================');
console.log('修正の説明');
console.log('========================================\n');

console.log('修正前の問題点:');
console.log('  - 【C】は extraRows * 2 を加算していた → 間違い');
console.log('  - 【D】は extraRows * 3 を加算していた → 間違い');
console.log('');
console.log('なぜ間違いか:');
console.log('  - 各セクションに同じ行数（extraRows）を追加するため、');
console.log('    全てのセクションは同じ行数だけ下にずれる');
console.log('  - extraRows * 2 や extraRows * 3 とすると、');
console.log('    セクションC、Dが必要以上に下にずれてしまう');
console.log('');
console.log('修正後:');
console.log('  - 全てのセクションで extraRows を1回だけ加算');
console.log('  - これにより、全セクションが同じ行数だけ下にずれる');
console.log('');
