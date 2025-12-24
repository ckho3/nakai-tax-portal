const ExcelJS = require('exceljs');

/**
 * 新規不動産データをExcelに書き込む
 * @param {string} excelPath - Excelファイルのパス
 * @param {Array} propertiesData - 物件データの配列
 * @param {number} annualIncomeCount - 年間収支一覧表PDFの件数（オプション）
 * @returns {Object} 結果
 */
async function writeNewProperties(excelPath, propertiesData, annualIncomeCount = 0) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    const sheet = workbook.getWorksheet('【不】新規不動産');
    if (!sheet) {
      throw new Error('【不】新規不動産シートが見つかりません');
    }

    const incomeSheet = workbook.getWorksheet('【不】①不動産収入');
    if (!incomeSheet) {
      throw new Error('【不】①不動産収入シートが見つかりません');
    }

    console.log(`\n===== 新規不動産データの書き込み開始 =====`);
    console.log(`物件数: ${propertiesData.length}件`);

    // テンプレート（A1-E15）をコピーして必要な数だけ配置
    const TEMPLATE_ROWS = 15;
    const TEMPLATE_COLS = 5; // A-E列
    const MAX_COLS_PER_ROW = 5; // 横に5セットまで
    const ROW_GROUP_HEIGHT = 16; // 15行 + 1行空白

    // 必要なテンプレート数を計算
    const templateCount = propertiesData.length;

    // テンプレートをコピー（2件目以降）
    for (let i = 1; i < templateCount; i++) {
      const colIndex = i % MAX_COLS_PER_ROW; // 横位置（0-4）
      const rowGroupIndex = Math.floor(i / MAX_COLS_PER_ROW); // 縦位置（0, 1, 2...）

      // コピー先の開始位置
      const targetStartCol = 1 + colIndex * TEMPLATE_COLS; // 1, 6, 11, 16, 21...
      const targetStartRow = 1 + rowGroupIndex * ROW_GROUP_HEIGHT;

      // テンプレート範囲（A1-E15）をコピー
      copyTemplate(sheet, 1, 1, targetStartRow, targetStartCol, TEMPLATE_ROWS, TEMPLATE_COLS);
    }

    // 各物件を順番に書き込み
    for (let i = 0; i < propertiesData.length; i++) {
      const data = propertiesData[i];

      const colIndex = i % MAX_COLS_PER_ROW; // 横位置（0-4）
      const rowGroupIndex = Math.floor(i / MAX_COLS_PER_ROW); // 縦位置

      // 書き込み開始位置
      const startCol = 1 + colIndex * TEMPLATE_COLS; // 1(A), 6(F), 11(K), 16(P), 21(U)
      const startRow = 1 + rowGroupIndex * ROW_GROUP_HEIGHT;

      // 列の計算: startCol=1 → B(2), C(3), D(4)
      //           startCol=6 → G(7), H(8), I(9)
      const nameCol = startCol + 1;  // B, G, L, Q, V列
      const monthCol = startCol + 2; // C, H, M, R, W列
      const amountCol = startCol + 3; // D, I, N, S, X列

      writePropertyData(sheet, data, startRow, nameCol, monthCol, amountCol);

      console.log(`  物件${i + 1}: ${data.propertyName} → Row ${startRow}, Col ${startCol}`);
    }

    // 【不】①不動産収入シートに物件名を追加
    console.log(`\n===== 【不】①不動産収入シートへの物件名追加 =====`);

    // 注意: G列（不動産の所在地）への書き込みは excelWriter.js で行われるため、ここでは行わない
    // このファイルでは、propertiesData（決済明細書PDF）の物件名を【A】セクションのH列に追加

    // 【A】セクションのH列（物件名）に公式を追加（G列を参照）
    console.log(`\n【A】セクションのH列（物件名）に決済明細書の物件名公式を追加`);

    const propertyNameCol = 8; // H列
    const propertyLocationCol = 7; // G列
    const sectionAStartRow = 55; // 【A】セクションの開始行

    // 年間収支一覧表PDFの件数から最終行を計算
    // 【A】セクションは Row 55 から開始、年間収支一覧表PDFの件数分の行が既に埋まっている
    const lastRowH = annualIncomeCount > 0 ? (sectionAStartRow + annualIncomeCount - 1) : 54;

    console.log(`年間収支一覧表PDF件数: ${annualIncomeCount}件`);
    console.log(`H列（物件名）の最終行（計算値）: Row ${lastRowH}`);

    // G列（Row 4-43）で最後のデータ行を見つける
    let lastRowG = 3; // Row 4の前の行
    for (let rowNum = 4; rowNum <= 43; rowNum++) {
      const cell = incomeSheet.getRow(rowNum).getCell(propertyLocationCol);
      if (cell.value) {
        lastRowG = rowNum;
      }
    }

    console.log(`G列（物件情報テーブル）の最終行: Row ${lastRowG}`);

    // 決済明細書の物件名がG列に書き込まれているのは lastRowG - propertiesData.length + 1 以降
    const gStartRow = lastRowG - propertiesData.length + 1;

    // propertiesDataの物件名公式をRow 55以降のH列に追加
    propertiesData.forEach((data, index) => {
      const targetRow = lastRowH + index + 1;
      const gRowRef = gStartRow + index; // G列の参照行
      const row = incomeSheet.getRow(targetRow);

      // G列を参照する公式を設定
      row.getCell(propertyNameCol).value = { formula: `G${gRowRef}` };
      console.log(`  物件${index + 1}: ${data.propertyName} → Row ${targetRow}, H列 (=G${gRowRef})（【A】セクション）`);
    });

    console.log(`【不】①不動産収入シートへの追加完了\n`);

    // ファイルを保存（元のファイルパスをそのまま使用）
    const path = require('path');
    const outputPath = excelPath;
    await workbook.xlsx.writeFile(outputPath);

    console.log(`\n新規不動産データの書き込み完了`);
    console.log(`出力ファイル: ${outputPath}`);

    return {
      success: true,
      outputPath,
      propertyCount: propertiesData.length
    };
  } catch (error) {
    console.error('新規不動産データの書き込みエラー:', error);
    throw new Error(`Excelの更新に失敗しました: ${error.message}`);
  }
}

/**
 * テンプレート範囲をコピーする
 * @param {Object} sheet - ワークシート
 * @param {number} srcStartRow - コピー元開始行
 * @param {number} srcStartCol - コピー元開始列
 * @param {number} destStartRow - コピー先開始行
 * @param {number} destStartCol - コピー先開始列
 * @param {number} rowCount - 行数
 * @param {number} colCount - 列数
 */
function copyTemplate(sheet, srcStartRow, srcStartCol, destStartRow, destStartCol, rowCount, colCount) {
  for (let r = 0; r < rowCount; r++) {
    const srcRow = sheet.getRow(srcStartRow + r);
    const destRow = sheet.getRow(destStartRow + r);

    for (let c = 0; c < colCount; c++) {
      const srcCell = srcRow.getCell(srcStartCol + c);
      const destCell = destRow.getCell(destStartCol + c);

      // セルの値をコピー（数式も含む）
      // ExcelJSでは formula, formulaType, result を個別にチェック
      if (srcCell.formula || srcCell.value?.formula) {
        // 数式の場合、相対参照を調整
        const colOffset = destStartCol - srcStartCol;
        const rowOffset = destStartRow - srcStartRow;

        let formulaText = srcCell.formula || srcCell.value.formula;
        const adjustedFormula = adjustFormula(formulaText, rowOffset, colOffset);

        // 数式を設定
        if (srcCell.value?.sharedFormula) {
          // 共有数式の場合
          destCell.value = {
            formula: adjustedFormula,
            sharedFormula: srcCell.value.sharedFormula
          };
        } else {
          // 通常の数式
          destCell.value = { formula: adjustedFormula };
        }

        console.log(`数式コピー: Row ${srcStartRow + r}, Col ${srcStartCol + c} → Row ${destStartRow + r}, Col ${destStartCol + c}`);
        console.log(`  元: ${formulaText} → 調整後: ${adjustedFormula}`);
      } else {
        destCell.value = srcCell.value;
      }

      // スタイルをコピー
      destCell.style = { ...srcCell.style };

      // セルの幅をコピー（列ごとに1回だけ）
      if (r === 0) {
        const srcColObj = sheet.getColumn(srcStartCol + c);
        const destColObj = sheet.getColumn(destStartCol + c);
        if (srcColObj.width) {
          destColObj.width = srcColObj.width;
        }
      }
    }

    // 行の高さをコピー
    if (srcRow.height) {
      destRow.height = srcRow.height;
    }
  }
}

/**
 * 数式の相対参照を調整する
 * @param {string} formula - 元の数式
 * @param {number} rowOffset - 行のオフセット
 * @param {number} colOffset - 列のオフセット
 * @returns {string} 調整された数式
 */
function adjustFormula(formula, rowOffset, colOffset) {
  // セル参照を調整: SUM(D2:D14) → SUM(I2:I14) など
  return formula.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
    // 列名を数値に変換
    let colNum = 0;
    for (let i = 0; i < col.length; i++) {
      colNum = colNum * 26 + (col.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }

    // オフセットを適用
    const newColNum = colNum + colOffset;
    const newRow = parseInt(row) + rowOffset;

    // 数値を列名に戻す
    let newCol = '';
    let num = newColNum;
    while (num > 0) {
      const remainder = (num - 1) % 26;
      newCol = String.fromCharCode('A'.charCodeAt(0) + remainder) + newCol;
      num = Math.floor((num - 1) / 26);
    }

    return `${newCol}${newRow}`;
  });
}

/**
 * 1物件分のデータを書き込む
 * @param {Object} sheet - ワークシート
 * @param {Object} data - 物件データ
 * @param {number} startRow - 開始行
 * @param {number} nameCol - 物件名列（B列またはG列）
 * @param {number} monthCol - 月数列（C列またはH列）
 * @param {number} amountCol - 金額列（D列またはI列）
 */
function writePropertyData(sheet, data, startRow, nameCol, monthCol, amountCol) {
  // 物件名（Row 1のB列またはG列）
  sheet.getRow(startRow).getCell(nameCol).value = data.propertyName;

  // 各項目の金額を書き込み（Row 2-14のD列またはI列）
  const items = [
    { row: startRow + 1, amount: data.loanFee },           // ローン事務手数料
    { row: startRow + 2, amount: data.businessFee },       // 事業者事務手数料
    { row: startRow + 3, amount: data.insuranceFee },      // 火災保険料
    { row: startRow + 4, amount: data.transferFee },       // 振込手数料
    { row: startRow + 5, amount: data.registrationFee },   // 登記費用
    { row: startRow + 6, amount: data.propertyTaxBuilding }, // 固都税精算金（建物）
    { row: startRow + 7, amount: data.propertyTaxLand },   // 固都税精算金（土地）
    { row: startRow + 8, amount: data.managementFee },     // 管理費・修繕積立金
    { row: startRow + 9, amount: data.stampSales },        // 収入印紙（売契）
    { row: startRow + 10, amount: data.stampLoan },        // 収入印紙（金消）
    { row: startRow + 11, amount: data.travelFee },        // 出張手数料
    { row: startRow + 12, amount: data.loanTransferFee },  // 融資金振込手数料
    { row: startRow + 13, amount: data.certificateFee },   // 証明書代金
  ];

  items.forEach(item => {
    const row = sheet.getRow(item.row);
    // 金額を書き込み（0の場合も書き込む）
    row.getCell(amountCol).value = item.amount;
  });

  // 管理費・修繕積立金の月数をC列またはH列のRow 9に書き込む
  if (data.managementFeeMonth) {
    const managementFeeRow = sheet.getRow(startRow + 8); // Row 9（管理費・修繕積立金の行）
    // monthCol（C列またはH列）に月数を書き込む
    managementFeeRow.getCell(monthCol).value = data.managementFeeMonth;
  }

  // 合計（Row 15）は数式 =SUM(D2:D14) があるのでそのまま
  console.log(`    合計: ¥${data.total.toLocaleString()}`);
}

module.exports = { writeNewProperties };
