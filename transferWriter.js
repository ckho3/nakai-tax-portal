const ExcelJS = require('exceljs');
const path = require('path');

/**
 * 【不】④耐用年数シートに譲渡対価証明書の情報を書き込む
 * @param {string} excelPath - Excelファイルのパス
 * @param {Array} transferData - 譲渡対価証明書のデータ配列
 * @param {Array} pdfDataArray - 年間収支一覧表PDFのデータ配列
 * @returns {Object} 処理結果
 */
async function writeTransferDates(excelPath, transferData, pdfDataArray = []) {
  try {
    console.log(`\n===== 【不】④耐用年数シートへの書き込み =====`);

    // Excelファイルを読み込み
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    // 【不】④耐用年数シートを取得
    const sheet = workbook.getWorksheet('【不】④耐用年数');
    if (!sheet) {
      throw new Error('【不】④耐用年数シートが見つかりません');
    }

    // 【不】①不動産収入シートを取得
    const incomeSheet = workbook.getWorksheet('【不】①不動産収入');
    if (!incomeSheet) {
      throw new Error('【不】①不動産収入シートが見つかりません');
    }

    // 【不】①不動産収入シートのG列から物件名一覧を取得
    const propertyNamesMap = new Map(); // 物件名 → 【不】④耐用年数シートの行番号

    // 【不】①不動産収入シートの全行をループ
    incomeSheet.eachRow((incomeRow, incomeRowNumber) => {
      if (incomeRowNumber < 4) return; // ヘッダー行をスキップ（Row 1-3）

      const cellG = incomeRow.getCell(7); // G列（物件名）
      let propertyName = cellG.value;

      // 数式セルの場合は result を取得
      if (propertyName && typeof propertyName === 'object' && propertyName.result !== undefined) {
        propertyName = propertyName.result;
      }

      if (propertyName) {
        const trimmedName = propertyName.toString().trim();

        // 物件名が既にマップに存在しない場合のみ追加
        if (!propertyNamesMap.has(trimmedName)) {
          // 【不】④耐用年数シートの行番号を計算
          // 【不】①不動産収入のRow 4（最初の物件）→ 【不】④耐用年数のRow 5（最初の物件の最初の行）
          // 各物件は【不】④耐用年数シートで2行使用するため：
          // propertyIndex = incomeRowNumber - 4 (0-indexed)
          // durabilityRow = 5 + (propertyIndex * 2)
          const propertyIndex = incomeRowNumber - 4;
          const durabilityRow = 5 + (propertyIndex * 2);

          console.log(`  Adding to map: "${trimmedName}" → 【不】①不動産収入 Row ${incomeRowNumber} → 【不】④耐用年数 Row ${durabilityRow}`);
          propertyNamesMap.set(trimmedName, durabilityRow);
        }
      }
    });

    console.log(`\n【不】①不動産収入シート: ${propertyNamesMap.size}件の物件名をマッピングしました`);
    console.log('\n===== 【不】①不動産収入シートから取得した物件名一覧 =====');
    for (const [name, rowNum] of propertyNamesMap.entries()) {
      console.log(`  Row ${rowNum}: "${name}"`);
    }
    console.log('==================================================\n');

    let updateCount = 0;
    const results = [];

    // フォルダパスから物件名を特定するマップを作成
    // フォルダパス → 物件名のマップ
    const folderToPropertyMap = new Map();
    if (pdfDataArray && pdfDataArray.length > 0) {
      console.log('\n===== フォルダパスから物件名のマッピングを作成 =====');
      for (const pdfData of pdfDataArray) {
        if (pdfData.folderPath && pdfData.propertyName) {
          folderToPropertyMap.set(pdfData.folderPath, pdfData.propertyName);
          console.log(`フォルダ: "${pdfData.folderPath}" → 物件名: "${pdfData.propertyName}"`);
        }
      }
      console.log('===============================================\n');
    }

    // 各譲渡対価証明書データを処理
    console.log('\n===== 譲渡対価証明書の物件名一覧 =====');
    for (const transfer of transferData) {
      console.log(`  物件名: "${transfer.propertyName}"`);
    }
    console.log('==================================================\n');

    for (const transfer of transferData) {
      const { propertyName, completionDate, transferDate, folderPath } = transfer;

      // 物件名がない場合はスキップ
      if (!propertyName) {
        console.log(`⚠ 物件名が取得できませんでした。スキップします。`);
        results.push({
          propertyName: `(不明)`,
          status: 'skipped',
          message: '物件名が取得できませんでした'
        });
        continue;
      }

      console.log(`\n処理中: 物件名「${propertyName.trim()}\"`);

      // propertyNamesMapから行番号を取得
      const foundRow = propertyNamesMap.get(propertyName.trim());

      if (!foundRow) {
        console.log(`  ⚠ 物件「${propertyName}」が【不】④耐用年数シートに見つかりませんでした`);
        results.push({
          propertyName,
          status: 'not_found',
          message: '物件が【不】④耐用年数シートに見つかりませんでした'
        });
        continue;
      }

      console.log(`  ✓ 【不】④耐用年数シート Row ${foundRow} に一致 (物件名: "${propertyName.trim()}"`);

      // F列（竣工日）とG列（購入日）に日付を書き込み（2行連続）
      // 1行目
      const row1 = sheet.getRow(foundRow);
      if (completionDate) {
        row1.getCell(6).value = completionDate; // F列: 竣工日
      }
      if (transferDate) {
        row1.getCell(7).value = transferDate; // G列: 購入日
      }
      // E列を「躯体」に設定
      row1.getCell(5).value = '躯体';

      // 2行目（次の行）
      const row2 = sheet.getRow(foundRow + 1);
      if (completionDate) {
        row2.getCell(6).value = completionDate; // F列: 竣工日
      }
      if (transferDate) {
        row2.getCell(7).value = transferDate; // G列: 購入日
      }
      // E列を「設備」に設定
      row2.getCell(5).value = '設備';

      updateCount++;
      console.log(`✓ 物件「${propertyName}」(Row ${foundRow}, ${foundRow + 1}): 竣工日=${completionDate || '(なし)'}, 購入日=${transferDate || '(なし)'}`);

      results.push({
        propertyName,
        status: 'success',
        message: `竣工日: ${completionDate || '(なし)'}, 購入日: ${transferDate || '(なし)'}`
      });
    }

    // Excelファイルを上書き保存
    await workbook.xlsx.writeFile(excelPath);

    console.log(`【不】④耐用年数シートへの書き込み完了: ${updateCount}/${transferData.length}件\n`);

    return {
      success: true,
      updateCount,
      results,
      outputPath: excelPath
    };

  } catch (error) {
    console.error('【不】④耐用年数シートへの書き込みエラー:', error);
    throw new Error(`【不】④耐用年数シートへの書き込みに失敗しました: ${error.message}`);
  }
}

module.exports = { writeTransferDates };
