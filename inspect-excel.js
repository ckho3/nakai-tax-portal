const ExcelJS = require('exceljs');
const path = require('path');

async function inspectExcel() {
  try {
    const workbook = new ExcelJS.Workbook();
    const excelPath = path.join(__dirname, '【原本】R7確定申告フォーマット.xlsx');
    await workbook.xlsx.readFile(excelPath);

    const sheet = workbook.getWorksheet('【不】①不動産収入');

    if (!sheet) {
      console.log('シート【不】①不動産収入が見つかりません');
      return;
    }

    console.log('========================================');
    console.log('Excelファイル構造の確認');
    console.log('========================================\n');

    // セクションの見出しを探す
    const sections = {
      'A': { name: '収入', startRow: null, endRow: null },
      'B': { name: '管理手数料', startRow: null, endRow: null },
      'C': { name: '広告費', startRow: null, endRow: null },
      'D': { name: '修繕費', startRow: null, endRow: null }
    };

    // Row 1-250を調査
    console.log('全ての行をスキャンしています...\n');

    for (let rowNum = 1; rowNum <= 250; rowNum++) {
      const row = sheet.getRow(rowNum);
      const cellE = row.getCell(5).value; // E列
      const cellF = row.getCell(6).value; // F列
      const cellG = row.getCell(7).value; // G列
      const cellH = row.getCell(8).value; // H列

      // 全ての非空行を表示（E-H列）
      if (cellE || cellF || cellG || cellH) {
        const cellEText = cellE ? cellE.toString().substring(0, 30) : '';
        const cellFText = cellF ? cellF.toString().substring(0, 30) : '';
        const cellGText = cellG ? cellG.toString().substring(0, 30) : '';
        const cellHText = cellH ? cellH.toString().substring(0, 30) : '';

        // 重要そうな行のみ表示
        if (cellEText.includes('【') || cellFText.includes('【') ||
            cellGText.includes('【') || cellHText.includes('【') ||
            cellEText.includes('収入') || cellFText.includes('収入') ||
            cellGText.includes('収入') || cellHText.includes('収入') ||
            cellEText.includes('管理') || cellFText.includes('管理') ||
            cellGText.includes('管理') || cellHText.includes('管理') ||
            cellEText.includes('広告') || cellFText.includes('広告') ||
            cellGText.includes('広告') || cellHText.includes('広告') ||
            cellEText.includes('修繕') || cellFText.includes('修繕') ||
            cellGText.includes('修繕') || cellHText.includes('修繕') ||
            cellEText.includes('支払') || cellFText.includes('支払') ||
            cellGText.includes('支払') || cellHText.includes('支払')) {
          console.log(`Row ${rowNum}: E="${cellEText}" F="${cellFText}" G="${cellGText}" H="${cellHText}"`);
        }
      }

      if (cellG) {
        const cellText = cellG.toString();

        // セクションヘッダーを検出
        if (cellText.includes('【A】') || (cellText.includes('収入') && rowNum >= 50 && rowNum <= 60)) {
          if (!sections.A.startRow) {
            sections.A.startRow = rowNum;
          }
        } else if (cellText.includes('【B】') || (cellText.includes('管理') && cellText.includes('手数料'))) {
          if (!sections.B.startRow) {
            sections.B.startRow = rowNum;
          }
        } else if (cellText.includes('【C】') || (cellText.includes('広告') && cellText.includes('費'))) {
          if (!sections.C.startRow) {
            sections.C.startRow = rowNum;
          }
        } else if (cellText.includes('【D】') || (cellText.includes('修繕') && cellText.includes('費'))) {
          if (!sections.D.startRow) {
            sections.D.startRow = rowNum;
          }
        }
      }
    }

    console.log('\n========================================');
    console.log('検出されたセクション:');
    console.log('========================================\n');

    for (const [key, section] of Object.entries(sections)) {
      if (section.startRow) {
        const row = sheet.getRow(section.startRow);
        const cellG = row.getCell(7).value;
        console.log(`【${key}】${section.name}: Row ${section.startRow} - G列: "${cellG}"`);
      }
    }

    console.log('\n========================================');
    console.log('各セクションのデータ行範囲を確認');
    console.log('========================================\n');

    // 各セクションのデータ行数をカウント
    for (const [key, section] of Object.entries(sections)) {
      if (section.startRow) {
        console.log(`\n【${key}】${section.name}セクション:`);
        console.log(`  ヘッダー行: Row ${section.startRow}`);

        // データ行を探す（ヘッダーの次の行から）
        let dataStartRow = section.startRow + 1;
        let dataEndRow = section.startRow + 1;
        let emptyCount = 0;

        for (let rowNum = dataStartRow; rowNum <= dataStartRow + 50; rowNum++) {
          const row = sheet.getRow(rowNum);
          const cellH = row.getCell(8).value; // H列（物件名）

          // 次のセクションヘッダーに到達したら終了
          const cellG = row.getCell(7).value;
          if (cellG && cellG.toString().includes('【')) {
            break;
          }

          if (cellH || cellG) {
            dataEndRow = rowNum;
            emptyCount = 0;
          } else {
            emptyCount++;
            // 連続5行が空なら終了
            if (emptyCount >= 5) {
              break;
            }
          }
        }

        const rowCount = dataEndRow - dataStartRow + 1;
        console.log(`  データ開始行: Row ${dataStartRow}`);
        console.log(`  データ終了行: Row ${dataEndRow}`);
        console.log(`  データ行数: ${rowCount}行`);

        // 最初の3行のデータを表示
        console.log(`  最初のデータ行サンプル:`);
        for (let i = 0; i < Math.min(3, rowCount); i++) {
          const rowNum = dataStartRow + i;
          const row = sheet.getRow(rowNum);
          const cellG = row.getCell(7).value;
          const cellH = row.getCell(8).value;
          console.log(`    Row ${rowNum}: G="${cellG || '(空)'}" H="${cellH || '(空)'}"`);
        }
      }
    }

    console.log('\n========================================');
    console.log('現在のコードの計算式チェック');
    console.log('========================================\n');

    const extraRows = 10; // 例: 10行追加する場合
    console.log(`extraRows = ${extraRows} の場合:\n`);

    for (let baseRow = 55; baseRow <= 57; baseRow++) {
      console.log(`物件がRow ${baseRow}に配置される場合:`);
      console.log(`  【A】収入: Row ${baseRow}`);
      console.log(`  【B】管理手数料: Row ${(78 + extraRows) + (baseRow - 55)} (現在の計算式)`);
      console.log(`  【C】広告費: Row ${(101 + extraRows * 2) + (baseRow - 55)} (現在の計算式)`);
      console.log(`  【D】修繕費: Row ${(124 + extraRows * 3) + (baseRow - 55)} (現在の計算式)`);
      console.log('');
    }

  } catch (error) {
    console.error('エラー:', error);
  }
}

inspectExcel();
