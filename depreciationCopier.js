const ExcelJS = require('exceljs');

/**
 * 【不】④耐用年数シートから【不】③減価償却(JDLよりエクスポート)シートへデータをコピー
 *
 * 戦略: 数式に頼らず、【不】⑤利息シートから直接データを取得
 * - C列（物件名）: 【不】⑤利息シートのB列から取得
 * - G列（取得価額）: 【不】⑤利息シートのC列（土地）+ D列（建物）
 * - L列（償却率）: 【不】④耐用年数シートのI列（耐用年数）から計算
 *
 * @param {string} excelPath - Excelファイルのパス
 * @returns {Object} 結果
 */
async function copyDepreciationData(excelPath) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    const usefulLifeSheet = workbook.getWorksheet('【不】④耐用年数');
    const depreciationSheet = workbook.getWorksheet('【不】③減価償却（JDLよりエクスポート）');
    const interestSheet = workbook.getWorksheet('【不】⑤利息');
    const incomeSheet = workbook.getWorksheet('【不】①不動産収入');

    if (!usefulLifeSheet) {
      throw new Error('【不】④耐用年数シートが見つかりません');
    }

    if (!depreciationSheet) {
      throw new Error('【不】③減価償却（JDLよりエクスポート）シートが見つかりません');
    }

    if (!interestSheet) {
      throw new Error('【不】⑤利息シートが見つかりません');
    }

    if (!incomeSheet) {
      throw new Error('【不】①不動産収入シートが見つかりません');
    }

    console.log(`\n===== 【不】④耐用年数 → 【不】③減価償却(JDLよりエクスポート) データ転記開始 =====`);

    const usefulLifeStartRow = 5; // 【不】④耐用年数シートのデータ開始行
    const interestStartRow = 10; // 【不】⑤利息シートのデータ開始行
    const incomeStartRow = 55; // 【不】①不動産収入シートの【A】セクション開始行
    const targetStartRow = 4; // 【不】③減価償却シートの開始行（E4, G4, I4から）

    // 【不】③減価償却シートの共有数式を通常の数式に変換
    console.log('【不】③減価償却シートの共有数式を通常の数式に変換しています...');
    for (let rowNum = 1; rowNum <= 300; rowNum++) {
      const row = depreciationSheet.getRow(rowNum);
      for (let colNum = 1; colNum <= 30; colNum++) {
        const cell = row.getCell(colNum);
        if (cell.formula && cell.type === ExcelJS.ValueType.Formula) {
          const formula = cell.formula;
          cell.value = { formula: formula };
        }
      }
    }
    console.log('【不】③減価償却シートの共有数式の変換が完了しました。');

    // 転記先のセル範囲（E4以降、G4以降、I4以降、J4以降）をクリア
    console.log('転記先セル範囲をクリアしています...');
    for (let rowNum = targetStartRow; rowNum <= targetStartRow + 200; rowNum++) {
      const row = depreciationSheet.getRow(rowNum);
      row.getCell(5).value = null; // E列
      row.getCell(7).value = null; // G列
      row.getCell(9).value = null; // I列
      row.getCell(10).value = null; // J列
    }
    console.log('転記先セル範囲のクリアが完了しました。\n');

    let copyCount = 0;
    let interestStartRowNum = null; // 【不】⑤利息シートの開始行番号

    // 【不】④耐用年数シートの行を順番に確認
    // 最大200行をチェック（動的に追加された行も含む）
    for (let rowNum = usefulLifeStartRow; rowNum <= usefulLifeStartRow + 200; rowNum++) {
      const usefulLifeRow = usefulLifeSheet.getRow(rowNum);
      const eCell = usefulLifeRow.getCell(5); // E列（構造・用途）

      // E列に値がある場合のみ処理
      if (eCell.value) {
        // 【不】④耐用年数は2行で1セット（躯体・設備）なので、2で割る
        const setIndex = Math.floor((rowNum - usefulLifeStartRow) / 2);

        // 【不】①不動産収入シートの物件情報テーブル（G列 Row 4-43）から物件名を取得
        // 【不】④耐用年数のRow 5, 6 → G4
        // 【不】④耐用年数のRow 7, 8 → G5
        const propertyInfoRowNum = 4 + setIndex;
        const propertyInfoRow = incomeSheet.getRow(propertyInfoRowNum);
        const propertyNameCell = propertyInfoRow.getCell(7); // G列（不動産の所在地）
        const propertyName = getDirectValue(propertyNameCell);

        // 【不】④耐用年数シートのG列から取得価額を取得
        const gCell = usefulLifeRow.getCell(7); // G列（購入日 = 取得価額）
        let acquisitionPrice = getDirectValue(gCell);

        // G列が日付オブジェクトとして取得された場合、Excelのシリアル値に戻す
        // Excel内部では数値として保存されているが、ExcelJSが日付に変換している
        // 日付シリアル値 = (日付 - 1899/12/30) の日数
        if (acquisitionPrice instanceof Date) {
          // Excelの日付シリアル値（1900年1月1日を1とする）
          const excelEpoch = new Date(1899, 11, 30); // 1899年12月30日
          const diffTime = acquisitionPrice.getTime() - excelEpoch.getTime();
          const diffDays = diffTime / (1000 * 60 * 60 * 24);
          acquisitionPrice = Math.round(diffDays);
        }

        // 最初の1件のみ: 【不】⑤利息シートのD列で「『減価償却』の「取得費」に転記」を探して、その下から検索開始
        if (copyCount === 0) {
          let headerRowNum = null;

          // まず「『減価償却』の「取得費」に転記」というテキストを探す
          for (let searchRow = interestStartRow; searchRow <= interestStartRow + 200; searchRow++) {
            const interestRow = interestSheet.getRow(searchRow);
            const dCell = interestRow.getCell(4); // D列
            const dValue = getCellValue(dCell, workbook); // 数式のresultも取得

            if (dValue && typeof dValue === 'string' && dValue.includes('『減価償却』の「取得費」に転記')) {
              headerRowNum = searchRow;
              console.log(`  【不】⑤利息シートのD列で「　『減価償却』の「取得費」に転記」を Row ${searchRow} で発見`);
              break;
            }
          }

          // ヘッダー行が見つかった場合、その下から実際にデータがある行を探す
          if (headerRowNum) {
            console.log(`  【不】⑤利息シートでC/D列に数値があり、G列が使われている行を検索中...`);

            // ヘッダー行以降でD列に数式があり、その数式が参照するG行のC/D列に数値がある行を探す
            for (let searchRow = headerRowNum + 1; searchRow <= headerRowNum + 200; searchRow++) {
              const searchRowData = interestSheet.getRow(searchRow);
              const dCell = searchRowData.getCell(4); // D列
              const dValue = dCell.value;

              // D列に数式がある
              const hasDFormula = dValue && typeof dValue === 'object' &&
                ('formula' in dValue || 'sharedFormula' in dValue);

              if (hasDFormula) {
                const formula = dValue.formula || dValue.sharedFormula;

                // 数式がG列を参照しているか確認（G10, G11, G12...）
                const gMatch = formula.match(/G(\d+)/);

                if (gMatch) {
                  const gRowNum = parseInt(gMatch[1]);

                  // そのG行のC列またはD列に数値があるかチェック
                  const gRow = interestSheet.getRow(gRowNum);
                  const gCCell = gRow.getCell(3);
                  const gDCell = gRow.getCell(4);

                  const hasCValue = typeof gCCell.value === 'number' && gCCell.value > 0;
                  const hasDValue = typeof gDCell.value === 'number' && gDCell.value > 0;

                  if (hasCValue || hasDValue) {
                    interestStartRowNum = searchRow;
                    console.log(`  【不】⑤利息シートのデータ開始行を Row ${interestStartRowNum} に設定`);
                    console.log(`    D列数式: ${formula}`);
                    console.log(`    参照行: G${gRowNum}`);
                    console.log(`    G${gRowNum}のC列: ${gCCell.value} (土地)`);
                    console.log(`    G${gRowNum}のD列: ${gDCell.value} (建物)`);

                    // 開始行以降、IFERROR(INT(G*C),0) パターンの全行のC列に0.8を書き込み
                    console.log(`\n  【不】⑤利息シートの該当行すべてのC列に0.8を書き込み中...`);
                    let writeCount = 0;
                    const pattern = /IFERROR\(INT\(G\d+\*C\d+\),0\)/;

                    for (let writeRow = searchRow; writeRow <= searchRow + 200; writeRow++) {
                      const targetRowData = interestSheet.getRow(writeRow);
                      const targetDCell = targetRowData.getCell(4); // D列
                      const targetDValue = targetDCell.value;

                      if (targetDValue && typeof targetDValue === 'object' && 'formula' in targetDValue) {
                        const targetFormula = targetDValue.formula;

                        // パターンに一致する場合、C列に0.8を書き込み
                        if (pattern.test(targetFormula)) {
                          const targetCCell = targetRowData.getCell(3); // C列
                          targetCCell.value = 0.8;
                          writeCount++;

                          if (writeCount <= 5) {
                            console.log(`    Row ${writeRow}: C列に0.8を書き込み (D列: ${targetFormula})`);
                          }
                        }
                      } else {
                        // D列に数式がない場合、ループを終了
                        break;
                      }
                    }

                    console.log(`  合計 ${writeCount}行のC列に0.8を書き込みました\n`);

                    break;
                  }
                }
              }
            }

            if (!interestStartRowNum) {
              console.log(`  警告: 【不】⑤利息シートでC/D列に数値がある行が見つかりませんでした`);
            }
          } else {
            console.log(`  警告: 【不】⑤利息シートでヘッダー行が見つかりませんでした`);
          }

          console.log(`  デバッグ:`);
          console.log(`    【不】④耐用年数 Row ${rowNum} → 【不】③減価償却 Row ${targetStartRow + copyCount}`);
          console.log(`    物件名: "${propertyName}"`);
          console.log(`    取得価額: ${acquisitionPrice}`);
          console.log(`    償却率: 【不】④耐用年数 L${rowNum}列を参照する数式を設定`);
          console.log(`    利息開始行: ${interestStartRowNum || '(未発見)'}`);
        }

        // 転記先の行番号を計算
        const targetRowNum = targetStartRow + copyCount;
        const targetRow = depreciationSheet.getRow(targetRowNum);

        // E列: 物件名（値）
        targetRow.getCell(5).value = propertyName;

        // G列: 取得価額（値）
        targetRow.getCell(7).value = acquisitionPrice;

        // I列: 償却率（【不】④耐用年数シートのL列を参照する数式）
        // ExcelJSは自動的に = を追加するので、数式には = を含めない
        const depreciationFormula = `'【不】④耐用年数'!L${rowNum}`;
        targetRow.getCell(9).value = { formula: depreciationFormula };

        // J列: 【不】⑤利息シートのD列を連番で参照
        // 最初の物件名から特定した開始行 + copyCount で順番に参照
        if (interestStartRowNum) {
          const currentInterestRow = interestStartRowNum + copyCount;
          const interestFormula = `'【不】⑤利息'!D${currentInterestRow}`;
          targetRow.getCell(10).value = { formula: interestFormula };
          console.log(`  Row ${rowNum} → Row ${targetRowNum}: 物件名="${propertyName}", 取得価額=${formatNumber(acquisitionPrice)}, 償却率=L${rowNum}参照, 利息=D${currentInterestRow}参照`);
        } else {
          // 開始行が見つからない場合は空白
          targetRow.getCell(10).value = null;
          console.log(`  Row ${rowNum} → Row ${targetRowNum}: 物件名="${propertyName}", 取得価額=${formatNumber(acquisitionPrice)}, 償却率=L${rowNum}参照, 利息=(開始行未発見)`);
        }

        copyCount++;
      }
    }

    console.log(`\n転記完了: ${copyCount}件のデータを【不】③減価償却(JDLよりエクスポート)シートに転記しました`);

    // ファイルを保存（元のファイルパスをそのまま使用）
    const path = require('path');
    const outputPath = excelPath;
    await workbook.xlsx.writeFile(outputPath);

    console.log(`出力ファイル: ${outputPath}\n`);

    return {
      success: true,
      outputPath,
      copyCount
    };

  } catch (error) {
    console.error('減価償却データのコピーエラー:', error);
    throw new Error(`減価償却データのコピーに失敗しました: ${error.message}`);
  }
}

/**
 * セルの直接の値を取得（数式は無視）
 * @param {Object} cell - ExcelJSのセルオブジェクト
 * @returns {any} セルの値（数式の場合はnull）
 */
function getDirectValue(cell) {
  if (!cell || cell.value === null || cell.value === undefined) {
    return null;
  }

  const value = cell.value;

  // 数式オブジェクトの場合は、resultがあればそれを返す
  if (typeof value === 'object' && value !== null) {
    if (value instanceof Date) {
      return value;
    }

    // resultプロパティがあればそれを返す
    if ('result' in value && value.result !== undefined && value.result !== null) {
      return value.result;
    }

    // 数式でresultがない場合はnull
    if ('formula' in value || 'sharedFormula' in value) {
      return null;
    }
  }

  // 通常の値
  return value;
}

/**
 * セルの値を取得（数式の場合は計算結果を取得）
 * @param {Object} cell - ExcelJSのセルオブジェクト
 * @param {Object} workbook - ワークブックオブジェクト（他のシート参照用）
 * @returns {any} セルの値または計算結果
 */
function getCellValue(cell, workbook = null) {
  if (!cell || cell.value === null || cell.value === undefined) {
    return null;
  }

  const value = cell.value;

  // 数式オブジェクトの場合（{ formula: "...", result: ... }）
  if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
    // DateオブジェクトやRichTextオブジェクトではない場合
    if (value instanceof Date) {
      return value;
    }

    // 数式オブジェクトの場合
    if ('formula' in value || 'sharedFormula' in value) {
      // resultがある場合はそれを返す
      if ('result' in value && value.result !== undefined && value.result !== null) {
        return value.result;
      }

      // resultがない場合、数式が他のシートを参照しているか確認
      const formula = value.formula || '';

      // 他のシート参照の場合（例: '【不】⑤利息'!B95）
      if (formula.includes('!') && workbook) {
        try {
          const match = formula.match(/'?([^'!]+)'?!([A-Z]+)(\d+)/);
          if (match) {
            const sheetName = match[1];
            const colLetter = match[2];
            const rowNum = parseInt(match[3]);

            const refSheet = workbook.getWorksheet(sheetName);
            if (refSheet) {
              const colNum = columnLetterToNumber(colLetter);
              const refCell = refSheet.getRow(rowNum).getCell(colNum);
              // 再帰的に値を取得（無限ループを防ぐため、workbookは渡さない）
              return getCellValue(refCell, null);
            }
          }
        } catch (error) {
          console.log(`  警告: 数式の解決に失敗しました: ${formula}`);
        }
      }

      // 数式を解決できない場合はnull
      return null;
    }
  }

  // 通常の値（数値、文字列、日付など）
  return value;
}

/**
 * 列文字を数値に変換（A=1, B=2, ..., Z=26, AA=27, ...）
 * @param {string} letters - 列文字（例: "A", "AB"）
 * @returns {number} 列番号
 */
function columnLetterToNumber(letters) {
  let num = 0;
  for (let i = 0; i < letters.length; i++) {
    num = num * 26 + (letters.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return num;
}

/**
 * 数値を見やすくフォーマット（ログ出力用）
 * @param {any} value - 値
 * @returns {string} フォーマットされた文字列
 */
function formatNumber(value) {
  if (value === null || value === undefined) {
    return '(空)';
  }
  if (typeof value === 'number') {
    return value.toLocaleString();
  }
  return String(value);
}

module.exports = { copyDepreciationData };
