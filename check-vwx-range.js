/**
 * V-X列の数式書き込み範囲を確認
 */

console.log('========================================');
console.log('V-X列の数式書き込み範囲（コード内）');
console.log('========================================\n');

const extraRows = 3; // 23件PDFの場合

console.log('【A】セクション:');
const aDataStart = 55;
const aDataEnd = 74 + extraRows;
console.log(`  開始行: Row ${aDataStart}`);
console.log(`  終了行: Row ${aDataEnd}`);
console.log(`  範囲: Row ${aDataStart}-${aDataEnd} (${aDataEnd - aDataStart + 1}行)`);
console.log('');

console.log('【B】セクション:');
const bDataStart = 80 + extraRows;
const bDataEnd = 99 + extraRows * 2;
console.log(`  開始行: Row ${bDataStart}`);
console.log(`  終了行: Row ${bDataEnd}`);
console.log(`  範囲: Row ${bDataStart}-${bDataEnd} (${bDataEnd - bDataStart + 1}行)`);
console.log('');

console.log('【C】セクション:');
const cDataStart = 103 + extraRows * 2;
const cDataEnd = 122 + extraRows * 3;
console.log(`  開始行: Row ${cDataStart}`);
console.log(`  終了行: Row ${cDataEnd}`);
console.log(`  範囲: Row ${cDataStart}-${cDataEnd} (${cDataEnd - cDataStart + 1}行)`);
console.log('');

console.log('【D】セクション:');
const dDataStart = 126 + extraRows * 3;
const dDataEnd = 145 + extraRows * 4;
console.log(`  開始行: Row ${dDataStart}`);
console.log(`  終了行: Row ${dDataEnd}`);
console.log(`  範囲: Row ${dDataStart}-${dDataEnd} (${dDataEnd - dDataStart + 1}行)`);
console.log('');

console.log('========================================');
console.log('各セクションの構造（23件PDFの場合）');
console.log('========================================\n');

console.log('【A】セクション（収入）:');
console.log('  Row 55: ヘッダー兼データ1行目（数式あり）✅');
console.log('  Row 56-77: データ2-23行目（数式あり）✅');
console.log('  Row 78: 空行');
console.log('  Row 79: 合計行（数式なし）');
console.log('');

console.log('【B】セクション（管理手数料）:');
console.log('  Row 81: ヘッダー（数式あり？）← 要確認');
console.log('  Row 82: 注釈（数式あり？）← 要確認');
console.log('  Row 83: データ1行目（数式あり）✅');
console.log('  Row 84-105: データ2-23行目（数式あり）✅');
console.log('  Row 106: 空行');
console.log('  Row 107: 合計行（数式なし）');
console.log('');

console.log('【C】セクション（広告費）:');
console.log('  Row 107: ヘッダー（数式あり？）← 要確認');
console.log('  Row 108: 注釈（数式あり？）← 要確認');
console.log('  Row 109: データ1行目（数式あり）✅');
console.log('  Row 110-131: データ2-23行目（数式あり）✅');
console.log('  Row 132: 空行');
console.log('  Row 133: 合計行（数式なし）');
console.log('');

console.log('【D】セクション（修繕費）:');
console.log('  Row 133: ヘッダー（数式あり？）← 要確認');
console.log('  Row 134: 注釈（数式あり？）← 要確認');
console.log('  Row 135: データ1行目（数式あり）✅');
console.log('  Row 136-157: データ2-23行目（数式あり）✅');
console.log('  Row 158: 空行');
console.log('  Row 159: 合計行（数式なし）');
console.log('');

console.log('========================================');
console.log('問題点の確認');
console.log('========================================\n');

console.log('❌ 【B】セクションの開始行が間違っている:');
console.log(`  コード: Row ${bDataStart} (Row 83)`);
console.log('  実際のヘッダー行: Row 81');
console.log('  実際の注釈行: Row 82');
console.log('  実際のデータ開始: Row 83');
console.log('');
console.log('  → Row 81-82にも数式があるべきか確認が必要');
console.log('');

console.log('❌ 【C】セクションの開始行が間違っている:');
console.log(`  コード: Row ${cDataStart} (Row 109)`);
console.log('  実際のヘッダー行: Row 107');
console.log('  実際の注釈行: Row 108');
console.log('  実際のデータ開始: Row 109');
console.log('');
console.log('  → Row 107-108にも数式があるべきか確認が必要');
console.log('');

console.log('❌ 【D】セクションの開始行が間違っている:');
console.log(`  コード: Row ${dDataStart} (Row 135)`);
console.log('  実際のヘッダー行: Row 133');
console.log('  実際の注釈行: Row 134');
console.log('  実際のデータ開始: Row 135');
console.log('');
console.log('  → Row 133-134にも数式があるべきか確認が必要');
console.log('');

console.log('========================================');
console.log('原本で確認すべきこと');
console.log('========================================\n');

console.log('1. Row 78（【B】のヘッダー）にV-X列の数式はあるか？');
console.log('2. Row 79（【B】の注釈）にV-X列の数式はあるか？');
console.log('3. Row 101（【C】のヘッダー）にV-X列の数式はあるか？');
console.log('4. Row 102（【C】の注釈）にV-X列の数式はあるか？');
console.log('5. Row 124（【D】のヘッダー）にV-X列の数式はあるか？');
console.log('6. Row 125（【D】の注釈）にV-X列の数式はあるか？');
