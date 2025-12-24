const ExcelJS = require('exceljs');
const path = require('path');

// Excelファイルのシート名を確認するスクリプト
async function checkSheetNames() {
  const excelPath = path.join(__dirname, '【原本】R7確定申告フォーマット.xlsx');

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    console.log('\n===== Excelファイル内のシート名一覧 =====');
    workbook.eachSheet((worksheet, sheetId) => {
      console.log(`シートID ${sheetId}: "${worksheet.name}"`);
    });
    console.log('==========================================\n');

    // 「減価償却」を含むシート名を検索
    console.log('「減価償却」を含むシート:');
    workbook.eachSheet((worksheet, sheetId) => {
      if (worksheet.name.includes('減価償却')) {
        console.log(`  "${worksheet.name}"`);
      }
    });

  } catch (error) {
    console.error('エラー:', error.message);
  }
}

checkSheetNames();
