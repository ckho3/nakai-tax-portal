const ExcelJS = require('exceljs');

async function test164InsertDetail() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/user/Documents/中井システム/システム依頼 _最新/pdf-to-excel-app/【原本】R7確定申告フォーマット.xlsx');

  const sheet = workbook.getWorksheet('【不】①不動産収入');

  console.log('========================================');
  console.log('duplicateRow(164)の詳細な動作確認');
  console.log('========================================\n');

  console.log('初期状態:');
  console.log(`  Row 147: ${sheet.getRow(147).getCell(3).value || ''}`);
  console.log(`  Row 163: ${sheet.getRow(163).getCell(3).value || ''}`);
  console.log(`  Row 164: ${sheet.getRow(164).getCell(3).value || ''}`);
  console.log(`  Row 165: ${sheet.getRow(165).getCell(3).value || ''}`);
  console.log(`  Row 166: ${sheet.getRow(166).getCell(3).value || ''}`);
  console.log('');

  console.log('duplicateRow(164, 1, true) を1回実行:');
  sheet.duplicateRow(164, 1, true);

  console.log(`  Row 147: ${sheet.getRow(147).getCell(3).value || ''}`);
  console.log(`  Row 163: ${sheet.getRow(163).getCell(3).value || ''}`);
  console.log(`  Row 164: ${sheet.getRow(164).getCell(3).value || ''} ← 新しく挿入された行`);
  console.log(`  Row 165: ${sheet.getRow(165).getCell(3).value || ''} ← 元のRow 164`);
  console.log(`  Row 166: ${sheet.getRow(166).getCell(3).value || ''} ← 元のRow 165`);
  console.log(`  Row 167: ${sheet.getRow(167).getCell(3).value || ''} ← 元のRow 166`);
  console.log('');

  console.log('結論:');
  console.log('  duplicateRow(164, 1, true) は Row 164の位置に挿入');
  console.log('  → Row 164以降が全て+1される');
  console.log('  → Row 147は影響を受けない（Row 164より前）');
}

test164InsertDetail().catch(console.error);
