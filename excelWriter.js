const ExcelJS = require('exceljs');
const path = require('path');

/**
 * 数式内のセル参照を調整する（行/列のオフセットを適用）
 * @param {string} formula - 元の数式
 * @param {number} rowOffset - 行のオフセット
 * @param {number} colOffset - 列のオフセット
 * @returns {string} 調整された数式
 */
function adjustFormulaReferences(formula, rowOffset, colOffset) {
  // セル参照を調整: B29 → B30, B31など
  return formula.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
    const newRow = parseInt(row) + rowOffset;
    // 列のオフセットは今回は使わないが、将来のために残す
    return `${col}${newRow}`;
  });
}

/**
 * D列の数式を調整する（G列は常にG10から、C/D列は行に応じて調整）
 * 例: =IFERROR(INT(G10*C41),0) → G10, G11, G12...（セット番号に応じて）
 *     =IFERROR(G10-D41,0) → G10は固定、D41は1つ前の行を参照
 * @param {string} formula - 元の数式
 * @param {number} setNumber - セット番号（0から開始、21件目は0、22件目は1...）
 * @param {number} currentRow - 現在の行番号（C/D列の参照に使用）
 * @param {boolean} isSecondRow - 2行目（偶数行）の場合true（D列で1つ前の行を参照）
 * @returns {string} 調整された数式
 */
function adjustDColumnFormula(formula, setNumber, currentRow, isSecondRow = false) {
  return formula.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
    let newRow;
    if (col === 'G') {
      // G列は常にG10から開始し、セット番号に応じて増加
      // 21件目(setNumber=0) → G10、22件目(setNumber=1) → G11...
      newRow = 10 + setNumber;
    } else if (col === 'D' && isSecondRow) {
      // 2行目のD列の場合、1つ前の行を参照
      newRow = currentRow - 1;
    } else {
      // C列、または1行目のD列は現在の行番号をそのまま使用
      newRow = currentRow;
    }
    return `${col}${newRow}`;
  });
}

/**
 * Row 4以降の物件情報テーブルで該当行を検索、なければ次の空き行を返す
 * @param {Object} worksheet - ExcelJSワークシート
 * @param {string} propertyName - 管理物件名
 * @returns {Object} {row: 行番号, isNew: 新規かどうか}
 */
function findOrCreatePropertyInfoRow(worksheet, propertyName) {
  // 管理物件名を正規化
  const normalize = (text) => {
    if (!text) return '';
    return text.replace(/[\s　]/g, '').toLowerCase();
  };

  const normalizedSearch = normalize(propertyName);

  // Row 4から検索開始（物件情報テーブル）
  for (let rowNum = 4; rowNum <= 50; rowNum++) {
    const row = worksheet.getRow(rowNum);
    const cellValue = row.getCell(7).value; // G列（不動産の所在地）

    if (cellValue) {
      const normalizedCell = normalize(cellValue.toString());

      // 部分一致で検索
      if (normalizedCell.includes(normalizedSearch) || normalizedSearch.includes(normalizedCell)) {
        return { row: rowNum, isNew: false };
      }
    } else {
      // 空の行が見つかった = 新規追加
      return { row: rowNum, isNew: true };
    }
  }

  // すべて埋まっている場合は最後の行の次
  return { row: 51, isNew: true };
}

/**
 * 物件名からエクセル内の該当行番号を検索、なければ次の空き行を返す（Row 55以降）
 * @param {Object} worksheet - ExcelJSワークシート
 * @param {string} propertyName - 物件名
 * @returns {Object} {row: 行番号, isNew: 新規かどうか}
 */
function findOrCreatePropertyRow(worksheet, propertyName) {
  // 物件名を正規化（全角スペース、半角スペース、記号を削除）
  const normalizePropertyName = (name) => {
    if (!name) return '';
    return name.replace(/[\s　・]/g, '').toLowerCase();
  };

  const normalizedSearchName = normalizePropertyName(propertyName);

  // Row 55から検索開始（【A】収入セクションのヘッダー行）
  // ヘッダー行から書き込む
  const startRow = 55;
  const initialMaxRow = 74; // 初期の最大行（20行目: Row 55-74）

  for (let rowNum = startRow; rowNum <= initialMaxRow + 500; rowNum++) {
    const row = worksheet.getRow(rowNum);
    const cellValue = row.getCell(8).value; // H列（物件名）

    if (cellValue) {
      const normalizedCellValue = normalizePropertyName(cellValue.toString());

      // 部分一致で検索
      if (normalizedCellValue.includes(normalizedSearchName) ||
          normalizedSearchName.includes(normalizedCellValue)) {
        return { row: rowNum, isNew: false, needsInsertion: false };
      }
    } else {
      // 空の行が見つかった = 新規追加
      // 20行目を超える場合は、行挿入が必要
      const needsInsertion = rowNum > initialMaxRow;
      return { row: rowNum, isNew: true, needsInsertion };
    }
  }

  // 500行すべて埋まっている場合はエラー
  throw new Error('収入セクションの行数が上限（500行）に達しました');
}

/**
 * Excelファイルにデータを書き込む
 * @param {string} excelPath - Excelファイルのパス
 * @param {Array} pdfDataArray - PDFから抽出したデータの配列
 * @param {Object} itemMapping - 項目名からセクション(B/C/D)へのマッピング
 * @param {Array} settlementFileNames - 決済明細書PDFのファイル名配列（オプション）
 * @returns {Promise<Object>} 処理結果
 */
async function updateExcel(excelPath, pdfDataArray, itemMapping = {}, settlementFileNames = [], propertiesData = [], folderName = '') {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    const sheet = workbook.getWorksheet('【不】①不動産収入');

    if (!sheet) {
      throw new Error('シート【不】①不動産収入が見つかりません');
    }

    // 数式のあるセル(H, V, W, X列)は一切触らない
    // データはG列、I-U列のみに書き込む

    // 共有数式の問題を回避するため、まず共有数式を通常の数式に変換
    // すべての列の共有数式を展開（広範囲に対応）
    console.log('共有数式を通常の数式に変換しています...');
    for (let rowNum = 1; rowNum <= 300; rowNum++) {
      const row = sheet.getRow(rowNum);
      // すべての列をチェック（1〜30列）
      for (let colNum = 1; colNum <= 30; colNum++) {
        const cell = row.getCell(colNum);
        if (cell.formula && cell.type === ExcelJS.ValueType.Formula) {
          // 共有数式を通常の数式に変換
          const formula = cell.formula;
          cell.value = { formula: formula };
        }
      }
    }
    console.log('共有数式の変換が完了しました。');

    // PDFの数をカウント（年間収支一覧表PDFと決済明細書PDFの合計）
    const annualIncomeCount = pdfDataArray.length;
    const settlementCount = settlementFileNames.length;
    const totalCount = annualIncomeCount + settlementCount;
    console.log(`処理するPDFファイル数: 年間収支=${annualIncomeCount}件, 決済明細書=${settlementCount}件, 合計=${totalCount}件`);

    // 20件を超える場合、各セクションに行を追加
    const extraRows = totalCount > 20 ? totalCount - 20 : 0;

    if (extraRows > 0) {
      console.log(`\n${extraRows}行を各セクションに追加します...`);

      // 各セクションの最終行の-2行目を複製
      // 重要: 各ループで行番号が動的に変わるため、累積オフセットを計算する
      for (let i = 0; i < extraRows; i++) {
        // 各セクションで追加された行数の累積
        // このループのi回目では、既に(i * 4)行が前のセクションで追加されている

        // 【E】Row 164を複製（原本の166 - 2行目）
        // 注意: 【D】【C】【B】【A】で各i行追加されているため、Row 164は164 + i*4の位置にずれている
        const eRowToDuplicate = 164 + i * 4;
        sheet.duplicateRow(eRowToDuplicate, 1, true);
        console.log(`  【E】Row ${eRowToDuplicate}を複製（${i + 1}回目）`);

        // 【D】Row 143を複製（原本の145 - 2行目）
        const dRowToDuplicate = 143 + i * 3;
        sheet.duplicateRow(dRowToDuplicate, 1, true);
        console.log(`  【D】Row ${dRowToDuplicate}を複製（${i + 1}回目）`);

        // 【C】Row 120を複製（原本の122 - 2行目）
        const cRowToDuplicate = 120 + i * 2;
        sheet.duplicateRow(cRowToDuplicate, 1, true);
        console.log(`  【C】Row ${cRowToDuplicate}を複製（${i + 1}回目）`);

        // 【B】Row 97を複製（原本の99 - 2行目）
        const bRowToDuplicate = 97 + i;
        sheet.duplicateRow(bRowToDuplicate, 1, true);
        console.log(`  【B】Row ${bRowToDuplicate}を複製（${i + 1}回目）`);

        // 【A】Row 73を複製（原本の75 - 2行目）
        sheet.duplicateRow(73, 1, true);
        console.log(`  【A】Row 73を複製（${i + 1}回目）`);
      }
      console.log('【不】①不動産収入シートの行追加完了\n');

      // 他のシートにも行を追加
      console.log(`\n他のシートにも行を追加します...`);

      // 【不】⑤利息シート - Row 29をextraRows回複製
      const interestSheet = workbook.getWorksheet('【不】⑤利息');
      if (interestSheet) {
        console.log(`\n【不】⑤利息シート: Row 29を${extraRows}回複製`);
        for (let i = 0; i < extraRows; i++) {
          interestSheet.duplicateRow(29, 1, true);
          console.log(`  Row 29を複製（${i + 1}/${extraRows}回目）`);
        }

        // 数式の調整: 追加された各行の数式を修正
        console.log('【不】⑤利息シートの数式を調整中...');
        for (let i = 0; i < extraRows; i++) {
          const currentRow = 30 + i; // Row 29の次の行から (Row 30, 31, 32...)
          const row = interestSheet.getRow(currentRow);
          const rowOffset = i + 1; // Row 30はRow 29から+1, Row 31は+2...

          // A列: 連番(21 + i)
          row.getCell(1).value = 21 + i;

          // B列: '【不】①不動産収入'!H73の参照を調整
          const refRow = 73 + (currentRow - 28);
          const bFormula = `'【不】①不動産収入'!H${refRow}`;
          row.getCell(2).value = { formula: bFormula };

          // F列: 元の数式を取得して行番号を調整（例: =B29 → =B30, =B31...）
          const fCell = row.getCell(6);
          if (fCell.value && fCell.value.formula) {
            const adjustedF = adjustFormulaReferences(fCell.value.formula, rowOffset, 0);
            row.getCell(6).value = { formula: adjustedF };
          }

          // G列: 元の数式を取得して行番号を調整
          const gCell = row.getCell(7);
          if (gCell.value && gCell.value.formula) {
            const adjustedG = adjustFormulaReferences(gCell.value.formula, rowOffset, 0);
            row.getCell(7).value = { formula: adjustedG };
          }

          // I列: 元の数式を取得して行番号を調整
          const iCell = row.getCell(9);
          if (iCell.value && iCell.value.formula) {
            const adjustedI = adjustFormulaReferences(iCell.value.formula, rowOffset, 0);
            row.getCell(9).value = { formula: adjustedI };
          }

          // J列: 元の数式を取得して行番号を調整
          const jCell = row.getCell(10);
          if (jCell.value && jCell.value.formula) {
            const adjustedJ = adjustFormulaReferences(jCell.value.formula, rowOffset, 0);
            row.getCell(10).value = { formula: adjustedJ };
          }

          // K列: 元の数式を取得して行番号を調整
          const kCell = row.getCell(11);
          if (kCell.value && kCell.value.formula) {
            const adjustedK = adjustFormulaReferences(kCell.value.formula, rowOffset, 0);
            row.getCell(11).value = { formula: adjustedK };
          }

          // L列: 元の数式を取得して行番号を調整
          const lCell = row.getCell(12);
          if (lCell.value && lCell.value.formula) {
            const adjustedL = adjustFormulaReferences(lCell.value.formula, rowOffset, 0);
            row.getCell(12).value = { formula: adjustedL };
          }

          console.log(`  Row ${currentRow}: 連番=${21 + i}, B列=${bFormula}, F-L列の数式を調整`);
        }

        // C41-D80の範囲の数式を調整（行が追加された分だけ行番号をずらす）
        console.log(`\n【不】⑤利息シート: C41-D80の数式を調整中...`);
        const formulaStartRow = 41 + extraRows; // 元のRow 41は extraRows分下にずれる
        const formulaEndRow = 80 + extraRows;   // 元のRow 80も同様にずれる

        for (let rowNum = formulaStartRow; rowNum <= formulaEndRow; rowNum++) {
          const row = interestSheet.getRow(rowNum);

          // C列の数式を調整
          const cCell = row.getCell(3);
          if (cCell.value && cCell.value.formula) {
            const adjustedC = adjustFormulaReferences(cCell.value.formula, extraRows, 0);
            row.getCell(3).value = { formula: adjustedC };
            console.log(`  Row ${rowNum} C列: ${cCell.value.formula} → ${adjustedC}`);
          }

          // D列の数式を調整
          const dCell = row.getCell(4);
          if (dCell.value && dCell.value.formula) {
            // 元の行番号を計算（extraRows分ずれているので戻す）
            const originalRow = rowNum - extraRows;
            // setNumberを計算: Row 41 → 0, Row 43 → 1, Row 45 → 2...
            const setNumber = Math.floor((originalRow - 41) / 2);
            // 偶数行（2行目）かどうかを判定: Row 42, 44, 46... → true
            const isSecondRow = (originalRow - 41) % 2 === 1;
            // D列の公式を調整（G列は常にG10から、C/D列は現在の行番号を使用、偶数行は1つ前の行を参照）
            const adjustedD = adjustDColumnFormula(dCell.value.formula, setNumber, rowNum, isSecondRow);
            row.getCell(4).value = { formula: adjustedD };
            console.log(`  Row ${rowNum} D列: ${dCell.value.formula} → ${adjustedD} (setNumber=${setNumber}, isSecondRow=${isSecondRow})`);
          }
        }
        console.log('【不】⑤利息シート: C41-D80の数式調整完了\n');

        // 20個以上PDFがある場合、A41-D42テンプレートセットを追加で複製
        const totalPDFs = pdfDataArray.length + (propertiesData ? propertiesData.length : 0);
        if (totalPDFs > 20) {
          const extraSets = totalPDFs - 20;
          console.log(`\n【不】⑤利息シート: PDF総数${totalPDFs}件 → ${extraSets}セット追加で複製が必要`);

          // 最後のテンプレートセット（Row 79-80）の位置（extraRows分オフセット適用後）
          const lastTemplateRow1 = 79 + extraRows; // Row 79の現在位置
          const lastTemplateRow2 = 80 + extraRows; // Row 80の現在位置

          console.log(`  Row ${lastTemplateRow1}-${lastTemplateRow2}（最後のテンプレートセット）を基準に${extraSets}セット手動コピー`);

          // 最後のテンプレート行（Row 79-80）のセル値と数式を読み取る
          const template1 = interestSheet.getRow(lastTemplateRow1);
          const template2 = interestSheet.getRow(lastTemplateRow2);

          // 各セルのデータを保存（A-D列）
          const templateData = [
            {
              a: { value: template1.getCell(1).value, style: template1.getCell(1).style },
              b: { value: template1.getCell(2).value, style: template1.getCell(2).style },
              c: { value: template1.getCell(3).value, style: template1.getCell(3).style },
              d: { value: template1.getCell(4).value, style: template1.getCell(4).style }
            },
            {
              a: { value: template2.getCell(1).value, style: template2.getCell(1).style },
              b: { value: template2.getCell(2).value, style: template2.getCell(2).style },
              c: { value: template2.getCell(3).value, style: template2.getCell(3).style },
              d: { value: template2.getCell(4).value, style: template2.getCell(4).style }
            }
          ];

          // デバッグ: テンプレートの内容を表示
          console.log(`\n  テンプレート Row ${lastTemplateRow1}:`);
          console.log(`    A列: ${JSON.stringify(templateData[0].a.value)}`);
          console.log(`    B列: ${JSON.stringify(templateData[0].b.value)}`);
          console.log(`    C列: ${JSON.stringify(templateData[0].c.value)}`);
          console.log(`    D列: ${JSON.stringify(templateData[0].d.value)}`);
          console.log(`  テンプレート Row ${lastTemplateRow2}:`);
          console.log(`    A列: ${JSON.stringify(templateData[1].a.value)}`);
          console.log(`    B列: ${JSON.stringify(templateData[1].b.value)}`);
          console.log(`    C列: ${JSON.stringify(templateData[1].c.value)}`);
          console.log(`    D列: ${JSON.stringify(templateData[1].d.value)}\n`);

          // Row 81以降に順番にコピー
          const startRow = lastTemplateRow2 + 1; // Row 80の次（Row 81）

          for (let setIndex = 0; setIndex < extraSets; setIndex++) {
            const targetRow1 = startRow + (setIndex * 2);
            const targetRow2 = targetRow1 + 1;

            const row1 = interestSheet.getRow(targetRow1);
            const row2 = interestSheet.getRow(targetRow2);

            // 1行目をコピー
            // A列（番号）- Row 79のA列の値に基づいて連番（20 → 21, 22, 23...）
            const aValue1 = templateData[0].a.value;
            if (typeof aValue1 === 'number') {
              row1.getCell(1).value = aValue1 + (setIndex + 1);
            } else {
              row1.getCell(1).value = aValue1;
            }
            row1.getCell(1).style = { ...templateData[0].a.style };

            // B列（物件名数式）- Row 79のB列の公式を基準に調整
            // 例: Row 79が=B29 → Row 81は=B30、Row 83は=B31...
            const bFormula1 = templateData[0].b.value?.formula || templateData[0].b.value;
            if (typeof bFormula1 === 'string' && templateData[0].b.value?.formula) {
              const adjustedB = adjustFormulaReferences(bFormula1, (setIndex + 1), 0);
              row1.getCell(2).value = { formula: adjustedB };
              console.log(`      Row ${targetRow1} B列: ${bFormula1} → ${adjustedB}`);
            } else {
              row1.getCell(2).value = bFormula1;
            }
            row1.getCell(2).style = { ...templateData[0].b.style };

            // C列（数式）- Row 79のC列: 例: Row 87が1（または定数） → Row 81も1
            // または Row 87が=1-C86の場合 → Row 81は=1-C80（1行前を参照）
            const cFormula1 = templateData[0].c.value?.formula || templateData[0].c.value;
            if (typeof cFormula1 === 'string' && templateData[0].c.value?.formula) {
              // C列は1行前を参照するパターン: =1-C86 → =1-C80（targetRow1 - 1）
              const adjustedC = adjustFormulaReferences(cFormula1, (setIndex + 1) * 2, 0);
              row1.getCell(3).value = { formula: adjustedC };
              console.log(`      Row ${targetRow1} C列: ${cFormula1} → ${adjustedC}`);
            } else {
              row1.getCell(3).value = cFormula1;
            }
            row1.getCell(3).style = { ...templateData[0].c.style };

            // D列（数式）- Row 79のD列: 例: =IFERROR(INT(G37*C87),0)
            // 注意: G列は20件目の続きから（21件目=G30, 22件目=G31...）
            // C/D列は現在の行番号を使用
            const dFormula1 = templateData[0].d.value?.formula || templateData[0].d.value;
            if (typeof dFormula1 === 'string' && templateData[0].d.value?.formula) {
              // D列の公式を個別に調整
              // setIndex: 0, 1, 2... → setIndex + 20 = 20, 21, 22... → G30, G31, G32...
              const adjustedD = adjustDColumnFormula(dFormula1, setIndex + 20, targetRow1);
              row1.getCell(4).value = { formula: adjustedD };
              console.log(`      Row ${targetRow1} D列: ${dFormula1} → ${adjustedD}`);
            } else {
              row1.getCell(4).value = dFormula1;
            }
            row1.getCell(4).style = { ...templateData[0].d.style };

            // 2行目をコピー
            // A列（必ず空白）
            row2.getCell(1).value = null;
            row2.getCell(1).style = { ...templateData[1].a.style };

            // B列（物件名数式）- Row 80のB列: Row 79と同じ参照 =B29 → Row 82も=B30
            const bFormula2 = templateData[1].b.value?.formula || templateData[1].b.value;
            if (typeof bFormula2 === 'string' && templateData[1].b.value?.formula) {
              const adjustedB = adjustFormulaReferences(bFormula2, (setIndex + 1), 0);
              row2.getCell(2).value = { formula: adjustedB };
              console.log(`      Row ${targetRow2} B列: ${bFormula2} → ${adjustedB}`);
            } else {
              row2.getCell(2).value = bFormula2;
            }
            row2.getCell(2).style = { ...templateData[1].b.style };

            // C列（数式）- Row 80のC列: 例: =1-C87 → Row 82は=1-C81（1行前）
            const cFormula2 = templateData[1].c.value?.formula || templateData[1].c.value;
            if (typeof cFormula2 === 'string' && templateData[1].c.value?.formula) {
              const adjustedC = adjustFormulaReferences(cFormula2, (setIndex + 1) * 2, 0);
              row2.getCell(3).value = { formula: adjustedC };
              console.log(`      Row ${targetRow2} C列: ${cFormula2} → ${adjustedC}`);
            } else {
              row2.getCell(3).value = cFormula2;
            }
            row2.getCell(3).style = { ...templateData[1].c.style };

            // D列（数式）- Row 80のD列: 例: =IFERROR(G37-D79,0) → Row 82は=IFERROR(G30-D81,0)
            // 注意: G列は20件目の続きから（21件目=G30, 22件目=G31...）
            // D列は1つ前の行を参照（2行目なのでisSecondRow=true）
            const dFormula2 = templateData[1].d.value?.formula || templateData[1].d.value;
            if (typeof dFormula2 === 'string' && templateData[1].d.value?.formula) {
              // D列の公式を個別に調整
              // setIndex: 0, 1, 2... → setIndex + 20 = 20, 21, 22... → G30, G31, G32...
              const adjustedD = adjustDColumnFormula(dFormula2, setIndex + 20, targetRow2, true);
              row2.getCell(4).value = { formula: adjustedD };
              console.log(`      Row ${targetRow2} D列: ${dFormula2} → ${adjustedD}`);
            } else {
              row2.getCell(4).value = dFormula2;
            }
            row2.getCell(4).style = { ...templateData[1].d.style };

            console.log(`    セット${setIndex + 1}/${extraSets}: Row ${targetRow1}-${targetRow2}に手動コピー完了`);
          }

          console.log(`【不】⑤利息シート: ${extraSets}セットの追加複製完了\n`);
        }

        console.log('【不】⑤利息シートの数式調整完了\n');
      } else {
        console.log('⚠ 【不】⑤利息シートが見つかりません');
      }

      // 【不】④耐用年数シート - Row 44をextraRows × 3回複製
      const usefulLifeSheet = workbook.getWorksheet('【不】④耐用年数');
      if (usefulLifeSheet) {
        const usefulLifeInsertCount = extraRows * 3;
        console.log(`\n【不】④耐用年数シート: Row 44を${usefulLifeInsertCount}回複製（${extraRows} × 3）`);
        for (let i = 0; i < usefulLifeInsertCount; i++) {
          usefulLifeSheet.duplicateRow(44, 1, true);
          console.log(`  Row 44を複製（${i + 1}/${usefulLifeInsertCount}回目）`);
        }

        // 数式の調整: C列(Row 5以降すべて)の数式を調整
        console.log('【不】④耐用年数シートの数式を調整中...');

        // C列(C5以降すべて)の数式を調整: ='【不】⑤利息'!B41 → B(41+extraRows)
        // C5: B41, C6: B42, ..., C44: B80, C45: B81, C46: B82...
        console.log('  C列(C5以降すべて)の数式を調整...');

        // Row 5から Row 44 + usefulLifeInsertCount まで
        const totalRows = 44 + usefulLifeInsertCount;
        for (let rowNum = 5; rowNum <= totalRows; rowNum++) {
          const row = usefulLifeSheet.getRow(rowNum);
          // C5 → B41, C6 → B42, ... と順番に続く
          const originalRefRow = 41 + (rowNum - 5);
          const newRefRow = originalRefRow + extraRows;
          const cFormula = `'【不】⑤利息'!B${newRefRow}`;
          row.getCell(3).value = { formula: cFormula };
        }
        console.log(`  C5-C${totalRows}: ='【不】⑤利息'!B${41 + extraRows} ~ B${41 + extraRows + (totalRows - 5)}`);

        // B列の番号を1から振り直し（B5から開始）
        console.log('  B列の番号を1から振り直し...');
        for (let rowNum = 5; rowNum <= totalRows; rowNum++) {
          const row = usefulLifeSheet.getRow(rowNum);
          row.getCell(2).value = rowNum - 4; // B5=1, B6=2, B7=3...
        }
        console.log(`  B5-B${totalRows}: 1 ~ ${totalRows - 4}`);

        // E列のドロップダウン（データ検証）をコピー
        console.log('  E列のドロップダウンをコピー...');
        const templateRow = usefulLifeSheet.getRow(5);
        const templateValidation = templateRow.getCell(5).dataValidation;

        if (templateValidation) {
          // Row 45以降（新しく追加された行）にドロップダウンをコピー
          for (let rowNum = 45; rowNum <= totalRows; rowNum++) {
            const row = usefulLifeSheet.getRow(rowNum);
            row.getCell(5).dataValidation = { ...templateValidation };
          }
          console.log(`  E45-E${totalRows}: ドロップダウンをコピー完了`);
        } else {
          console.log('  ⚠ E5にドロップダウンが見つかりませんでした');
        }

        // I列の数式を調整（Row 5以降すべて）
        // =IF(E5="","",IF(E5="躯体",VLOOKUP(D5,$N$5:$O$9,2,FALSE),15))
        console.log('  I列の数式を調整...');
        for (let rowNum = 5; rowNum <= totalRows; rowNum++) {
          const row = usefulLifeSheet.getRow(rowNum);
          const iFormula = `IF(E${rowNum}="","",IF(E${rowNum}="躯体",VLOOKUP(D${rowNum},$N$5:$O$9,2,FALSE),15))`;
          row.getCell(9).value = { formula: iFormula };
        }
        console.log(`  I5-I${totalRows}: 数式を設定完了`);

        console.log('【不】④耐用年数シートの数式調整完了\n');
      } else {
        console.log('⚠ 【不】④耐用年数シートが見つかりません');
      }

      // 【不】②減価償却（新規入力用）シート - Row 40をextraRows回複製
      const depreciationSheet = workbook.getWorksheet('【不】②減価償却（新規入力用）');
      if (depreciationSheet) {
        console.log(`\n【不】②減価償却（新規入力用）シート: Row 40を${extraRows}回複製`);
        for (let i = 0; i < extraRows; i++) {
          depreciationSheet.duplicateRow(40, 1, true);
          console.log(`  Row 40を複製（${i + 1}/${extraRows}回目）`);
        }
        // duplicateRow(..., true)により相対参照は自動調整されます
        console.log('【不】②減価償却（新規入力用）シートの行追加完了\n');
      } else {
        console.log('⚠ 【不】②減価償却（新規入力用）シートが見つかりません');
      }

      console.log('全シートの行追加完了\n');

      // 行追加後、【E】サブリースセクションの数式を修正
      console.log('【E】セクションの数式を修正中...');

      // 【E】の開始行計算:
      // - 【A】【B】【C】【D】の影響: extraRows * 4
      // - 動的な行番号で複製するようにしたため、+1のズレは不要
      const eStartRow = 147 + extraRows * 4;
      const dStartRow = 124 + extraRows * 3; // 【D】の開始行
      const bStartRow = 78 + extraRows;      // 【B】の開始行

      console.log(`  【E】セクション開始: Row ${eStartRow}`);
      console.log(`  【D】セクション開始: Row ${dStartRow}`);
      console.log(`  【B】セクション開始: Row ${bStartRow}`);

      // 【E】セクションの各行（20行 + extraRows）を修正
      for (let i = 0; i < 20 + extraRows; i++) {
        const currentRow = eStartRow + i;
        const row = sheet.getRow(currentRow);

        // H列の数式を修正: =H124 → =H(dStartRow + i)
        const hCell = row.getCell(8); // H列
        const newHFormula = `H${dStartRow + i}`;
        hCell.value = { formula: newHFormula };

        // I-T列の数式を修正（各月の計算）
        for (let col = 9; col <= 20; col++) { // I列からT列
          const cell = row.getCell(col);
          const colLetter = String.fromCharCode(64 + col); // 9→I, 10→J, ...

          if (col === 20) { // T列（12月）の特殊な数式
            // ROUNDDOWN((T55-T78)*$B$146,-2) の形式
            // 【A】と【B】の対応する行を参照
            const aRow = 55 + i; // 【A】セクションの対応行（Row 55から開始）
            const bRow = bStartRow + i; // 【B】セクションの対応行（bStartRowから開始）
            const newFormula = `IF(T$${eStartRow - 1}>=$U${currentRow},IF(T$53="サブリース",ROUNDDOWN((T${aRow}-T${bRow})*$B$${eStartRow - 1},-2),0),0)`;
            cell.value = { formula: newFormula };
          } else {
            // I-S列の数式: IF(I$146>=$U147,IF(I$53="サブリース",J147,0),0)
            const nextColLetter = String.fromCharCode(64 + col + 1);
            const newFormula = `IF(${colLetter}$${eStartRow - 1}>=$U${currentRow},IF(${colLetter}$53="サブリース",${nextColLetter}${currentRow},0),0)`;
            cell.value = { formula: newFormula };
          }
        }
      }

      console.log('【E】セクションの数式修正完了\n');

      // 各セクションのA列（連番）を振り直し
      console.log('各セクションのA列の連番を振り直し中...');

      // 【A】セクション: Row 55-74（データ20行）+ extraRows
      const aRowStart = 55;
      const aRowCount = 20 + extraRows;
      for (let i = 0; i < aRowCount; i++) {
        sheet.getRow(aRowStart + i).getCell(1).value = i + 1;
      }
      console.log(`  【A】セクション: Row ${aRowStart}~${aRowStart + aRowCount - 1} に 1~${aRowCount} を設定`);

      // 【B】セクション: Row 78-97（データ20行）+ extraRows
      const bRowStart = 78 + extraRows;
      const bRowCount = 20 + extraRows;
      for (let i = 0; i < bRowCount; i++) {
        sheet.getRow(bRowStart + i).getCell(1).value = i + 1;
      }
      console.log(`  【B】セクション: Row ${bRowStart}~${bRowStart + bRowCount - 1} に 1~${bRowCount} を設定`);

      // 【C】セクション: Row 101-120（データ20行）+ extraRows
      const cRowStart = 101 + extraRows * 2;
      const cRowCount = 20 + extraRows;
      for (let i = 0; i < cRowCount; i++) {
        sheet.getRow(cRowStart + i).getCell(1).value = i + 1;
      }
      console.log(`  【C】セクション: Row ${cRowStart}~${cRowStart + cRowCount - 1} に 1~${cRowCount} を設定`);

      // 【D】セクション: Row 124-143（データ20行）+ extraRows
      const dRowStart = 124 + extraRows * 3;
      const dRowCount = 20 + extraRows;
      for (let i = 0; i < dRowCount; i++) {
        sheet.getRow(dRowStart + i).getCell(1).value = i + 1;
      }
      console.log(`  【D】セクション: Row ${dRowStart}~${dRowStart + dRowCount - 1} に 1~${dRowCount} を設定`);

      // 【E】セクション: Row 147-166（データ20行）+ extraRows
      const eRowStartA = 147 + extraRows * 4;
      const eRowCountA = 20 + extraRows;
      for (let i = 0; i < eRowCountA; i++) {
        sheet.getRow(eRowStartA + i).getCell(1).value = i + 1;
      }
      console.log(`  【E】セクション: Row ${eRowStartA}~${eRowStartA + eRowCountA - 1} に 1~${eRowCountA} を設定`);

      console.log('A列の連番振り直し完了\n');

      // 各セクションの合計行の数式を修正
      console.log('各セクションの合計行の数式を修正中...');

      const aSummaryRow = 76 + extraRows;       // 【A】の合計行
      const bSummaryRow = 99 + extraRows * 2;   // 【B】の合計行
      const cSummaryRow = 122 + extraRows * 3;  // 【C】の合計行
      const dSummaryRow = 145 + extraRows * 4;  // 【D】の合計行

      console.log(`  【A】合計行: Row ${aSummaryRow}`);
      console.log(`  【B】合計行: Row ${bSummaryRow}`);
      console.log(`  【C】合計行: Row ${cSummaryRow}`);
      console.log(`  【D】合計行: Row ${dSummaryRow}`);

      // 【A】の合計行: I53, J53, ..., T53を参照（変更なし、すでにduplicateRowで調整済み）
      // ※念のため明示的に設定
      const aRow = sheet.getRow(aSummaryRow);
      for (let col = 9; col <= 20; col++) {
        const colLetter = String.fromCharCode(64 + col);
        aRow.getCell(col).value = { formula: `${colLetter}53` };
      }

      // 【B】の合計行: 【A】の合計行を参照
      const bRow = sheet.getRow(bSummaryRow);
      for (let col = 9; col <= 20; col++) {
        const colLetter = String.fromCharCode(64 + col);
        bRow.getCell(col).value = { formula: `${colLetter}${aSummaryRow}` };
      }

      // 【C】の合計行: 【B】の合計行を参照
      const cRow = sheet.getRow(cSummaryRow);
      for (let col = 9; col <= 20; col++) {
        const colLetter = String.fromCharCode(64 + col);
        cRow.getCell(col).value = { formula: `${colLetter}${bSummaryRow}` };
      }

      // 【D】の合計行: 【C】の合計行を参照
      const dRow = sheet.getRow(dSummaryRow);
      for (let col = 9; col <= 20; col++) {
        const colLetter = String.fromCharCode(64 + col);
        dRow.getCell(col).value = { formula: `${colLetter}${cSummaryRow}` };
      }

      console.log('各セクションの合計行の数式修正完了\n');

      // V-X列の数式を修正（各データ行）
      console.log('V-X列の数式を修正中...');

      // 【A】セクションのデータ行（Row 55-74+extraRows）
      const aDataStart = 55;
      const aDataEnd = 74 + extraRows;
      for (let row = aDataStart; row <= aDataEnd; row++) {
        const rowObj = sheet.getRow(row);
        rowObj.getCell(22).value = { formula: `SUMIF($I$53:$T$53,V$53,$I${row}:$T${row})` }; // V列
        rowObj.getCell(23).value = { formula: `SUMIF($I$53:$T$53,W$53,$I${row}:$T${row})` }; // W列
        rowObj.getCell(24).value = { formula: `SUM(V${row}:W${row})` }; // X列
      }

      // 【A】セクションの合計行（Row 75+extraRows）
      const aSumRowVWX = 75 + extraRows;
      const aSumRowObj = sheet.getRow(aSumRowVWX);
      aSumRowObj.getCell(22).value = { formula: `SUM(V${aDataStart}:V${aDataEnd})` }; // V列
      aSumRowObj.getCell(23).value = { formula: `SUM(W${aDataStart}:W${aDataEnd})` }; // W列
      aSumRowObj.getCell(24).value = { formula: `SUM(X${aDataStart}:X${aDataEnd})` }; // X列

      // 【B】セクションのデータ行（Row 78+extraRows - 97+extraRows*2）
      // ヘッダー行（Row 78）と注釈行（Row 79）も含む、Row 98-100は数式なし
      const bDataStart = 78 + extraRows;
      const bDataEnd = 97 + extraRows * 2;
      for (let row = bDataStart; row <= bDataEnd; row++) {
        const rowObj = sheet.getRow(row);
        rowObj.getCell(22).value = { formula: `SUMIF($I$53:$T$53,V$53,$I${row}:$T${row})` }; // V列
        rowObj.getCell(23).value = { formula: `SUMIF($I$53:$T$53,W$53,$I${row}:$T${row})` }; // W列
        rowObj.getCell(24).value = { formula: `SUM(V${row}:W${row})` }; // X列
      }

      // 【C】セクションのデータ行（Row 101+extraRows*2 - 121+extraRows*3）
      // ヘッダー行（Row 101）と注釈行（Row 102）も含む、Row 122-123は数式なし
      const cDataStart = 101 + extraRows * 2;
      const cDataEnd = 121 + extraRows * 3;
      for (let row = cDataStart; row <= cDataEnd; row++) {
        const rowObj = sheet.getRow(row);
        rowObj.getCell(22).value = { formula: `SUMIF($I$53:$T$53,V$53,$I${row}:$T${row})` }; // V列
        rowObj.getCell(23).value = { formula: `SUMIF($I$53:$T$53,W$53,$I${row}:$T${row})` }; // W列
        rowObj.getCell(24).value = { formula: `SUM(V${row}:W${row})` }; // X列
      }

      // 【D】セクションのデータ行（Row 124+extraRows*3 - 143+extraRows*4）
      // ヘッダー行（Row 124）と注釈行（Row 125）も含む
      const dDataStart = 124 + extraRows * 3;
      const dDataEnd = 143 + extraRows * 4;
      for (let row = dDataStart; row <= dDataEnd; row++) {
        const rowObj = sheet.getRow(row);
        rowObj.getCell(22).value = { formula: `SUMIF($I$53:$T$53,V$53,$I${row}:$T${row})` }; // V列
        rowObj.getCell(23).value = { formula: `SUMIF($I$53:$T$53,W$53,$I${row}:$T${row})` }; // W列
        rowObj.getCell(24).value = { formula: `SUM(V${row}:W${row})` }; // X列
      }

      // 【D】セクションの合計行（Row 144+extraRows*4）
      const dSumRowVWX = 144 + extraRows * 4;
      const dSumRowVWXObj = sheet.getRow(dSumRowVWX);
      dSumRowVWXObj.getCell(22).value = { formula: `SUM(V${dDataStart}:V${dDataEnd})` }; // V列
      dSumRowVWXObj.getCell(23).value = { formula: `SUM(W${dDataStart}:W${dDataEnd})` }; // W列
      dSumRowVWXObj.getCell(24).value = { formula: `SUM(X${dDataStart}:X${dDataEnd})` }; // X列

      console.log('V-X列の数式修正完了\n');

      // 各セクションの合計行（SUM）の数式を修正
      console.log('各セクションのSUM合計行を修正中...');

      // 【C】セクションの合計行
      const cSumRow = 123 + extraRows * 3; // 元々Row 123
      const cDataStartForSum = 101 + extraRows * 2; // 【C】のヘッダー行
      const cDataEndForSum = 122 + extraRows * 3; // 【C】の最終データ行
      const cSumRowObj = sheet.getRow(cSumRow);
      cSumRowObj.getCell(22).value = { formula: `SUM(V${cDataStartForSum}:V${cDataEndForSum})` }; // V列
      cSumRowObj.getCell(23).value = { formula: `SUM(W${cDataStartForSum}:W${cDataEndForSum})` }; // W列
      cSumRowObj.getCell(24).value = { formula: `SUM(X${cDataStartForSum}:X${cDataEndForSum})` }; // X列
      console.log(`  【C】合計行 Row ${cSumRow}: SUM(V${cDataStartForSum}:V${cDataEndForSum})`);

      // 【D】セクションの合計行
      const dSumRow = 146 + extraRows * 4; // 元々Row 146
      const dDataStartForSum = 124 + extraRows * 3; // 【D】のヘッダー行
      const dDataEndForSum = 145 + extraRows * 4; // 【D】の最終データ行
      const dSumRowObj = sheet.getRow(dSumRow);
      dSumRowObj.getCell(22).value = { formula: `SUM(V${dDataStartForSum}:V${dDataEndForSum})` }; // V列
      dSumRowObj.getCell(23).value = { formula: `SUM(W${dDataStartForSum}:W${dDataEndForSum})` }; // W列
      dSumRowObj.getCell(24).value = { formula: `SUM(X${dDataStartForSum}:X${dDataEndForSum})` }; // X列
      console.log(`  【D】合計行 Row ${dSumRow}: SUM(V${dDataStartForSum}:V${dDataEndForSum})`);

      console.log('SUM合計行の修正完了\n');

      // H列の物件名参照数式を修正
      console.log('H列の物件名参照数式を修正中...');

      // 【B】セクションのH列（各データ行）: =H55（【A】の1件目）を参照
      // Row 78 + extraRows が【B】の開始行
      const bSectionStart = 78 + extraRows;
      const bDataRows = 20 + extraRows; // 【B】セクションのデータ行数（動的）

      for (let i = 0; i < bDataRows; i++) {
        const bRow = bSectionStart + i;
        const aRow = 55 + i; // 対応する【A】セクションの行
        sheet.getRow(bRow).getCell(8).value = { formula: `H${aRow}` };
      }
      console.log(`  【B】セクション H${bSectionStart}~H${bSectionStart + bDataRows - 1}: =H55~H${55 + bDataRows - 1}`);

      // 【C】セクションのH列（各データ行）: =H78（【B】の1件目）を参照
      const cSectionStart = 101 + extraRows * 2;
      const cDataRows = 20 + extraRows; // 【C】セクションのデータ行数（動的）

      for (let i = 0; i < cDataRows; i++) {
        const cRow = cSectionStart + i;
        const bRow = bSectionStart + i; // 対応する【B】セクションの行
        sheet.getRow(cRow).getCell(8).value = { formula: `H${bRow}` };
      }
      console.log(`  【C】セクション H${cSectionStart}~H${cSectionStart + cDataRows - 1}: =H${bSectionStart}~H${bSectionStart + cDataRows - 1}`);

      // 【D】セクションのH列（各データ行）: =H101（【C】の1件目）を参照
      const dSectionStart = 124 + extraRows * 3;
      const dDataRows = 20 + extraRows; // 【D】セクションのデータ行数（動的）

      for (let i = 0; i < dDataRows; i++) {
        const dRow = dSectionStart + i;
        const cRow = cSectionStart + i; // 対応する【C】セクションの行
        sheet.getRow(dRow).getCell(8).value = { formula: `H${cRow}` };
      }
      console.log(`  【D】セクション H${dSectionStart}~H${dSectionStart + dDataRows - 1}: =H${cSectionStart}~H${cSectionStart + dDataRows - 1}`);

      // 【E】セクションのH列（各データ行）: =H124（【D】の1件目）を参照
      const eSectionStart = 147 + extraRows * 4;
      const eDataRows = 20 + extraRows; // 【E】セクションのデータ行数（動的）

      for (let i = 0; i < eDataRows; i++) {
        const eRow = eSectionStart + i;
        const dRow = dSectionStart + i; // 対応する【D】セクションの行
        sheet.getRow(eRow).getCell(8).value = { formula: `H${dRow}` };
      }
      console.log(`  【E】セクション H${eSectionStart}~H${eSectionStart + eDataRows - 1}: =H${dSectionStart}~H${dSectionStart + eDataRows - 1}`);

      console.log('H列の物件名参照数式の修正完了\n');
    }

    // 【不】⑤利息シートへの決済明細書データの書き込み（extraRowsに関係なく実行）
    const interestSheet = workbook.getWorksheet('【不】⑤利息');
    if (interestSheet && propertiesData && propertiesData.length > 0) {
      console.log(`\n【不】⑤利息シート: 決済明細書データの書き込み`);

      // 【不】⑤利息シートのデータはRow 10から始まる
      // 年間収支一覧表PDF: Row 10 ~ (10 + annualIncomeCount - 1)
      // 決済明細書PDF: Row (10 + annualIncomeCount) ~
      const annualIncomeCount = pdfDataArray.length;
      const settlementStartRow = 10 + annualIncomeCount;

      propertiesData.forEach((property, index) => {
        const rowNum = settlementStartRow + index;
        const row = interestSheet.getRow(rowNum);

        // C列: 土地取得価額
        row.getCell(3).value = property.landPrice || 0;

        // D列: 建物取得価額
        row.getCell(4).value = property.buildingPrice || 0;

        console.log(`  Row ${rowNum}: ${property.propertyName} - 土地: ¥${(property.landPrice || 0).toLocaleString()}, 建物: ¥${(property.buildingPrice || 0).toLocaleString()}`);
      });

      console.log('【不】⑤利息シート: 決済明細書データの書き込み完了\n');
    }

    const results = [];

    for (let pdfIndex = 0; pdfIndex < pdfDataArray.length; pdfIndex++) {
      const pdfData = pdfDataArray[pdfIndex];
      const currentPdfNumber = pdfIndex + 1;  // 1件目、2件目...

      console.log(`\n${currentPdfNumber}件目を処理中: ${pdfData.propertyName || '不明'}`);
      const propertyName = pdfData.propertyName;

      if (!propertyName) {
        results.push({
          propertyName: '不明',
          status: 'error',
          message: '物件名が抽出できませんでした'
        });
        continue;
      }

      // ===== Row 4以降の物件情報テーブルに書き込み =====
      const { row: infoRow, isNew: isNewInfo } = findOrCreatePropertyInfoRow(sheet, propertyName);
      const propertyInfoRow = sheet.getRow(infoRow);

      console.log(`${isNewInfo ? '新規追加' : '更新'} (物件情報): ${propertyName} at row ${infoRow}`);

      // G列 (7): 不動産の所在地（管理物件名を入れる）
      const propertyLocationCell = propertyInfoRow.getCell(7);
      propertyLocationCell.value = propertyName;

      // H列 (8): 賃借人の住所・氏名
      if (pdfData.tenantName) {
        propertyInfoRow.getCell(8).value = pdfData.tenantName;
      }

      // I列 (9): 契約期間①自
      if (pdfData.contractStartDate) {
        propertyInfoRow.getCell(9).value = pdfData.contractStartDate;
      }

      // J列 (10): 契約期間①至
      if (pdfData.contractEndDate) {
        propertyInfoRow.getCell(10).value = pdfData.contractEndDate;
      }

      // M列 (13): 貸付面積（専有面積）
      if (pdfData.rentalArea) {
        propertyInfoRow.getCell(13).value = pdfData.rentalArea;
      }

      // P列 (16): 賃貸料年額
      if (pdfData.totalRent) {
        propertyInfoRow.getCell(16).value = pdfData.totalRent;
      }

      // ===== Row 55以降の月別収支データに書き込み =====
      // 各セクションの書き込み位置を計算
      // 【A】収入: Row 55 + pdfIndex
      // 【B】管理手数料: Row (78 + extraRows) + pdfIndex
      // 【C】広告費: Row (101 + extraRows * 2) + pdfIndex
      // 【D】修繕費: Row (124 + extraRows * 3) + pdfIndex

      // 【A】収入セクションに物件名と収入合計を書き込み（Row 55以降）
      const incomeRowNumber = 55 + pdfIndex;
      const incomeRow = sheet.getRow(incomeRowNumber);

      console.log(`物件${currentPdfNumber}: 【A】Row ${incomeRowNumber}に書き込み`);

      // G列: 収入項目名
      incomeRow.getCell(7).value = '収入合計①';

      // H列: 物件名（G列の物件情報テーブルを参照する数式）
      const gRowRef = infoRow; // G列の物件情報テーブルの行番号
      incomeRow.getCell(8).value = { formula: `G${gRowRef}` };

      // 月別賃料を書き込み（I列〜T列 = 9〜20）
      // I列=1月、J列=2月、...、T列=12月
      if (pdfData.monthlyRents && pdfData.monthlyRents.length === 12) {
        for (let month = 0; month < 12; month++) {
          const columnIndex = 9 + month; // I列(9)から始まる
          const value = pdfData.monthlyRents[month];
          if (value !== null && value !== undefined) {
            incomeRow.getCell(columnIndex).value = value;
          }
        }
      }

      // サブリース列（U列 = 21）のみ書き込み（V,W,X列は数式があるので触らない）
      incomeRow.getCell(21).value = pdfData.totalRent || 0; // 収入合計

      // 【B】管理手数料セクション
      const managementRowNumber = (78 + extraRows) + pdfIndex;
      const managementRow = sheet.getRow(managementRowNumber);

      console.log(`物件${currentPdfNumber}: 【B】Row ${managementRowNumber}に書き込み`);

      // G列: 支払項目名（PDFの項目名をそのまま使用）
      managementRow.getCell(7).value = '管理手数料';

      // H列: 物件名（G列の物件情報テーブルを参照する数式）
      managementRow.getCell(8).value = { formula: `G${gRowRef}` };

      // 月別管理手数料（I列〜T列 = 9〜20）
      const managementMonthlyTotals = new Array(12).fill(0);
      if (pdfData.managementFees && pdfData.managementFees.length === 12) {
        for (let month = 0; month < 12; month++) {
          managementMonthlyTotals[month] += pdfData.managementFees[month] || 0;
        }
      }

      // その他の支払項目でセクションBに分類されたものを加算
      if (pdfData.otherExpenseItems) {
        for (const [itemName, monthlyData] of Object.entries(pdfData.otherExpenseItems)) {
          if (itemMapping[itemName] === 'B' && monthlyData.length === 12) {
            for (let month = 0; month < 12; month++) {
              managementMonthlyTotals[month] += monthlyData[month] || 0;
            }
          }
        }
      }

      // 月別データを書き込み
      for (let month = 0; month < 12; month++) {
        const columnIndex = 9 + month;
        if (managementMonthlyTotals[month] > 0) {
          managementRow.getCell(columnIndex).value = managementMonthlyTotals[month];
        }
      }

      // 合計（U列 = 21）※V,W,X列は数式があるので触らない
      const totalManagementFee = managementMonthlyTotals.reduce((sum, val) => sum + val, 0);
      managementRow.getCell(21).value = totalManagementFee;

      // 【C】広告費等セクション
      const advertisingRowNumber = (101 + extraRows * 2) + pdfIndex;
      const advertisingRow = sheet.getRow(advertisingRowNumber);

      console.log(`物件${currentPdfNumber}: 【C】Row ${advertisingRowNumber}に書き込み`);

      // G列: 支払項目名（PDFの項目名をそのまま使用）
      advertisingRow.getCell(7).value = '宣伝広告費';

      // H列: 物件名（G列の物件情報テーブルを参照する数式）
      advertisingRow.getCell(8).value = { formula: `G${gRowRef}` };

      // 月別広告費（I列〜T列 = 9〜20）
      const advertisingMonthlyTotals = new Array(12).fill(0);
      if (pdfData.advertisingCosts && pdfData.advertisingCosts.length === 12) {
        for (let month = 0; month < 12; month++) {
          advertisingMonthlyTotals[month] += pdfData.advertisingCosts[month] || 0;
        }
      }

      // その他の支払項目でセクションCに分類されたものを加算
      if (pdfData.otherExpenseItems) {
        for (const [itemName, monthlyData] of Object.entries(pdfData.otherExpenseItems)) {
          if (itemMapping[itemName] === 'C' && monthlyData.length === 12) {
            for (let month = 0; month < 12; month++) {
              advertisingMonthlyTotals[month] += monthlyData[month] || 0;
            }
          }
        }
      }

      // 月別データを書き込み
      for (let month = 0; month < 12; month++) {
        const columnIndex = 9 + month;
        if (advertisingMonthlyTotals[month] > 0) {
          advertisingRow.getCell(columnIndex).value = advertisingMonthlyTotals[month];
        }
      }

      // 合計（U列 = 21）※V,W,X列は数式があるので触らない
      const totalAdvertising = advertisingMonthlyTotals.reduce((sum, val) => sum + val, 0);
      if (totalAdvertising > 0) {
        advertisingRow.getCell(21).value = totalAdvertising;
      }

      // 【D】修繕費・設備費セクション
      const repairRowNumber = (124 + extraRows * 3) + pdfIndex;
      const repairRow = sheet.getRow(repairRowNumber);

      console.log(`物件${currentPdfNumber}: 【D】Row ${repairRowNumber}に書き込み`);

      // G列: 支払項目名（PDFの項目名をそのまま使用）
      repairRow.getCell(7).value = '設備交換費';

      // H列: 物件名（G列の物件情報テーブルを参照する数式）
      repairRow.getCell(8).value = { formula: `G${gRowRef}` };

      // 月別修繕費（I列〜T列 = 9〜20）
      const repairMonthlyTotals = new Array(12).fill(0);
      if (pdfData.equipmentCosts && pdfData.equipmentCosts.length === 12) {
        for (let month = 0; month < 12; month++) {
          repairMonthlyTotals[month] += pdfData.equipmentCosts[month] || 0;
        }
      }

      // その他の支払項目でセクションDに分類されたものを加算
      if (pdfData.otherExpenseItems) {
        for (const [itemName, monthlyData] of Object.entries(pdfData.otherExpenseItems)) {
          if (itemMapping[itemName] === 'D' && monthlyData.length === 12) {
            for (let month = 0; month < 12; month++) {
              repairMonthlyTotals[month] += monthlyData[month] || 0;
            }
          }
        }
      }

      // 月別データを書き込み
      for (let month = 0; month < 12; month++) {
        const columnIndex = 9 + month;
        if (repairMonthlyTotals[month] > 0) {
          repairRow.getCell(columnIndex).value = repairMonthlyTotals[month];
        }
      }

      // 合計（U列 = 21）※V,W,X列は数式があるので触らない
      const totalRepair = repairMonthlyTotals.reduce((sum, val) => sum + val, 0);
      if (totalRepair > 0) {
        repairRow.getCell(21).value = totalRepair;
      }

      results.push({
        propertyName,
        status: 'success',
        message: `物件${currentPdfNumber}のデータを書き込みました`,
        rows: {
          income: incomeRowNumber,
          management: managementRowNumber,
          advertising: advertisingRowNumber,
          repair: repairRowNumber
        }
      });
    }

    // ===================================
    // V-X列の公式を修正とクリア
    // ===================================
    console.log('\n===== V-X列の公式を修正開始 =====');

    // 【C】セクションの合計行（Row 121 → 121 + extraRows * 3）
    const cSumRow = 121 + extraRows * 3;
    const cDataStart = 101 + extraRows * 2;
    const cDataEnd = 120 + extraRows * 3;
    const cSumRowObj = sheet.getRow(cSumRow);
    cSumRowObj.getCell(22).value = { formula: `SUM(V${cDataStart}:V${cDataEnd})` };
    cSumRowObj.getCell(23).value = { formula: `SUM(W${cDataStart}:W${cDataEnd})` };
    cSumRowObj.getCell(24).value = { formula: `SUM(X${cDataStart}:X${cDataEnd})` };
    console.log(`【C】合計行 Row ${cSumRow}: V=SUM(V${cDataStart}:V${cDataEnd})`);

    // 【C】セクション後のテキスト行と空白行をクリア（Row 122-123 → Row 122+extraRows*3 ~ 123+extraRows*3）
    const cTextRow = 122 + extraRows * 3; // 元Row 122「経費表へ転記」
    const cBlankRow = 123 + extraRows * 3; // 元Row 123（空白）
    sheet.getRow(cTextRow).getCell(22).value = '経費表へ転記';
    sheet.getRow(cTextRow).getCell(23).value = '（×）';
    sheet.getRow(cTextRow).getCell(24).value = null;
    sheet.getRow(cBlankRow).getCell(22).value = null;
    sheet.getRow(cBlankRow).getCell(23).value = null;
    sheet.getRow(cBlankRow).getCell(24).value = null;
    console.log(`【C】後のテキスト行 Row ${cTextRow}, 空白行 Row ${cBlankRow} をクリア`);

    // 【D】セクションの合計行（Row 144 → 144 + extraRows * 4）
    const dSumRow = 144 + extraRows * 4;
    const dDataStart = 124 + extraRows * 3;
    const dDataEnd = 143 + extraRows * 4;
    const dSumRowObj = sheet.getRow(dSumRow);
    dSumRowObj.getCell(22).value = { formula: `SUM(V${dDataStart}:V${dDataEnd})` };
    dSumRowObj.getCell(23).value = { formula: `SUM(W${dDataStart}:W${dDataEnd})` };
    dSumRowObj.getCell(24).value = { formula: `SUM(X${dDataStart}:X${dDataEnd})` };
    console.log(`【D】合計行 Row ${dSumRow}: V=SUM(V${dDataStart}:V${dDataEnd})`);

    // 【D】セクション後のテキスト行と空白行をクリア（Row 145-146 → Row 145+extraRows*4 ~ 146+extraRows*4）
    const dTextRow = 145 + extraRows * 4; // 元Row 145「（×）」
    const dBlankRow = 146 + extraRows * 4; // 元Row 146「←サブリース割合」
    sheet.getRow(dTextRow).getCell(22).value = '（×）';
    sheet.getRow(dTextRow).getCell(23).value = '（×）';
    sheet.getRow(dTextRow).getCell(24).value = '経費表へ転記';
    sheet.getRow(dBlankRow).getCell(22).value = null;
    sheet.getRow(dBlankRow).getCell(23).value = null;
    sheet.getRow(dBlankRow).getCell(24).value = null;
    console.log(`【D】後のテキスト行 Row ${dTextRow}, 空白行 Row ${dBlankRow} をクリア`);

    // 【E】セクションのデータ行（Row 147-166 → 動的計算）
    // 注意: 前のセクションで既に eStartRow を計算済み
    // しかし、このスコープでは見えないので再計算
    const eStartRowVWX = 147 + extraRows * 4; // V-X列用の【E】開始行
    const eDataRows = 20; // 原本は20行
    for (let i = 0; i < eDataRows + extraRows; i++) {
      const currentRow = eStartRowVWX + i;
      const rowObj = sheet.getRow(currentRow);
      rowObj.getCell(22).value = { formula: `SUMIF($I$53:$T$53,V$53,$I${currentRow}:$T${currentRow})` };
      rowObj.getCell(23).value = { formula: `SUMIF($I$53:$T$53,W$53,$I${currentRow}:$T${currentRow})` };
      rowObj.getCell(24).value = { formula: `SUM(V${currentRow}:W${currentRow})` };
    }
    console.log(`【E】データ行 Row ${eStartRowVWX}~${eStartRowVWX + eDataRows + extraRows - 1}: SUMIF公式を設定`);

    // 【E】セクションの合計行（Row 167 → 実測値に基づく計算）
    const eSumRow = eStartRowVWX + eDataRows + extraRows; // データ行の次の行
    const eDataEndRow = eStartRowVWX + eDataRows + extraRows - 1; // データ最終行
    const eSumRowObj = sheet.getRow(eSumRow);
    eSumRowObj.getCell(22).value = { formula: `SUM(V${eStartRowVWX}:V${eDataEndRow})` };
    eSumRowObj.getCell(23).value = { formula: `SUM(W${eStartRowVWX}:W${eDataEndRow})` };
    eSumRowObj.getCell(24).value = { formula: `SUM(X${eStartRowVWX}:X${eDataEndRow})` };
    console.log(`【E】合計行 Row ${eSumRow}: V=SUM(V${eStartRowVWX}:V${eDataEndRow})`);

    // 【E】セクション後のテキスト行をクリア（Row 168 → 実測値に基づく計算）
    const eTextRow = eSumRow + 1; // 元Row 168「（×）」
    sheet.getRow(eTextRow).getCell(22).value = '（×）';
    sheet.getRow(eTextRow).getCell(23).value = '上記表へ転記';
    sheet.getRow(eTextRow).getCell(24).value = null;
    console.log(`【E】後のテキスト行 Row ${eTextRow} をクリア`);

    console.log('===== V-X列の公式修正完了 =====\n');

    // 決済明細書PDFがアップロードされている場合、抽出された物件名をG列に追加
    if (propertiesData && propertiesData.length > 0) {
      console.log(`\n===== 決済明細書から物件名を抽出してG列に追加 =====`);
      console.log(`決済明細書件数: ${propertiesData.length}件`);

      const propertyLocationCol = 7; // G列

      // Row 4〜43の範囲で最終データ行を見つける
      let lastRow = 3; // Row 4の前の行
      for (let rowNum = 4; rowNum <= 43; rowNum++) {
        const cell = sheet.getRow(rowNum).getCell(propertyLocationCol);
        if (cell.value) {
          lastRow = rowNum;
        }
      }

      console.log(`物件情報テーブルの最終行: Row ${lastRow}`);

      // 各決済明細書から抽出された物件名をG列に追加
      propertiesData.forEach((property, index) => {
        const propertyName = property.propertyName;

        if (propertyName) {
          const targetRow = lastRow + index + 1;

          // Row 43を超えないようにチェック
          if (targetRow <= 43) {
            const row = sheet.getRow(targetRow);
            const cell = row.getCell(propertyLocationCol);
            cell.value = propertyName;

            console.log(`  物件${index + 1}: ${propertyName} → Row ${targetRow}, G列`);
          } else {
            console.log(`  ⚠ 物件${index + 1}: ${propertyName} → Row ${targetRow}は範囲外（G列はRow 43まで）のためスキップ`);
          }
        } else {
          console.log(`  ⚠ 物件${index + 1}: 物件名が抽出できませんでした`);
        }
      });

      console.log(`決済明細書からの物件名追加完了\n`);
    }

    // 更新されたファイルを保存
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
    const path = require('path');
    const outputDir = path.join(__dirname, 'output');

    // フォルダ名がある場合は「フォルダ名_R7確定申告_updated_タイムスタンプ.xlsx」
    // ない場合は従来通り
    let outputFilename;
    if (folderName) {
      outputFilename = `${folderName}_R7確定申告_updated_${timestamp}.xlsx`;
    } else {
      const originalFilename = path.basename(excelPath);
      outputFilename = originalFilename.replace('.xlsx', `_updated_${timestamp}.xlsx`);
    }

    const outputPath = path.join(outputDir, outputFilename);
    await workbook.xlsx.writeFile(outputPath);

    return {
      success: true,
      outputPath,
      results
    };

  } catch (error) {
    console.error('Excel update error:', error);
    throw new Error('Excelの更新に失敗しました: ' + error.message);
  }
}

module.exports = { updateExcel, findOrCreatePropertyRow };
