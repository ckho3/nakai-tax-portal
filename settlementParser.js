const pdfParse = require('pdf-parse');
const fs = require('fs').promises;

/**
 * 決済明細書PDFから情報を抽出
 * @param {string} pdfPath - PDFファイルのパス
 * @param {string} filename - ファイル名（オリジナル）
 * @param {string} folderPath - フォルダパス（オプション）
 * @returns {Object} 抽出されたデータ
 */
async function parseSettlementPDF(pdfPath, filename = '', folderPath = '') {
  try {
    const dataBuffer = await fs.readFile(pdfPath);
    const pdfData = await pdfParse(dataBuffer);
    const text = pdfData.text;

    // デバッグ: PDFテキストの最初の1000文字を出力
    console.log('\n===== PDFテキスト (最初の1000文字) =====');
    console.log(text.substring(0, 1000));
    console.log('=====================================\n');

    // 物件名を抽出 - ファイル名から取得
    // ファイル名フォーマット: "アスヴェル京都西京極 703_決済明細書.pdf"
    // → "_" の前の部分が物件名
    let propertyName = '';

    if (filename) {
      const match = filename.match(/^(.+?)_/);
      if (match && match[1]) {
        propertyName = match[1].trim();
        // 「決済明細書」というテキストが含まれている場合は削除
        propertyName = propertyName.replace(/決済明細書/g, '').trim();
        console.log(`\nファイル名から物件名を抽出: "${propertyName}" (ファイル名: "${filename}")`);
      }
    }

    // ファイル名から抽出できなかった場合、フォルダ名を使用
    if (!propertyName && folderPath) {
      // フォルダパスから最後のフォルダ名を取得
      const pathParts = folderPath.split('/');
      const folderName = pathParts[pathParts.length - 1] || pathParts[0] || '';
      if (folderName) {
        propertyName = folderName;
        console.log(`\nファイル名から抽出できなかったため、フォルダ名を使用: "${propertyName}" (フォルダパス: "${folderPath}")`);
      }
    }

    console.log(`\n最終的に抽出された物件名: "${propertyName}"`);

    // 決済日を抽出
    const dateMatch = text.match(/決済日[^\n]*\n([0-9\/]+)/);
    const settlementDate = dateMatch ? dateMatch[1] : '';

    // 管理費・修繕積立金の月数を抽出
    // 例: 「管理費・修繕積立金 (立替) （12月日割、1・2月分）」→ 一番後ろの2を抽出
    let managementFeeMonth = null;
    const managementFeeIndex = text.indexOf('管理費・修繕積立金');

    if (managementFeeIndex !== -1) {
      // 「管理費・修繕積立金」から次の150文字を抽出
      const managementFeeSection = text.substring(managementFeeIndex, managementFeeIndex + 150);

      console.log('\n===== 管理費・修繕積立金セクション =====');
      console.log(managementFeeSection);
      console.log('=====================================\n');

      // 「月」という文字の前にある1-12の数字を全て抽出
      const monthMatches = managementFeeSection.matchAll(/([1-9]|1[0-2])月/g);
      const months = [];

      for (const match of monthMatches) {
        const month = parseInt(match[1], 10);
        if (month >= 1 && month <= 12) {
          months.push(month);
        }
      }

      if (months.length > 0) {
        // 一番後ろの月を取得
        managementFeeMonth = months[months.length - 1];
        console.log(`抽出された月: [${months.join(', ')}] → 一番後ろ: ${managementFeeMonth}`);
      } else {
        console.log('⚠ 管理費月数が抽出できませんでした');
      }
    }

    // 諸費用エリアから金額を順番に抽出
    // PDFのフォーマット: 「諸費用」という文字の後に、各項目の金額が順番に並んでいる
    const expensesSection = text.substring(text.indexOf('諸費用'));
    const amounts = expensesSection.match(/¥([0-9,]+)/g);

    let amountIndex = 0;
    const getNextAmount = () => {
      if (amounts && amountIndex < amounts.length) {
        const amt = amounts[amountIndex++];
        const value = parseInt(amt.replace(/¥|,/g, ''), 10);
        return value;
      }
      return 0;
    };

    // 最初の2つの金額をスキップ（諸費用合計とその他）
    getNextAmount(); // ¥861,878 (諸費用合計)
    getNextAmount(); // ¥0

    // 各項目の金額を順番に取得
    const loanFee = getNextAmount();           // ¥330,000
    getNextAmount(); // "諸費用" の次の金額をスキップ
    const businessFee = getNextAmount();       // ¥198,000
    getNextAmount(); // 自己資金
    const insuranceFee = getNextAmount();      // ¥0
    getNextAmount(); // ¥0
    const transferFee = getNextAmount();       // ¥800,000 (これは融資金額?)
    const registrationFee = getNextAmount();   // ¥0 or 次の金額

    // 土地代金と建物代金を抽出
    const landPrice = extractLandPrice(text);
    const buildingPrice = extractBuildingPrice(text);

    // PDFテキストから金額を位置で抽出
    // PDFの金額の並び: ¥861,878, ¥0, ¥330,000, (¥800,000), ¥198,000, ¥0, ¥0, ¥800,000, ¥0, ¥263,678, ¥25,190, ¥16,507, ¥8,683, ¥25,010, ¥0, ¥20,000, ¥0, ¥0, ¥0...
    const data = {
      propertyName,
      settlementDate,
      landPrice,                                          // 土地代金
      buildingPrice,                                      // 建物代金
      loanFee: extractAmountByPosition(text, 2),          // ¥330,000
      businessFee: extractAmountByPosition(text, 4),      // ¥198,000
      insuranceFee: extractAmountByPosition(text, 5),     // ¥0
      transferFee: extractAmountByPosition(text, 6),      // ¥0
      registrationFee: extractAmountByPosition(text, 9),  // ¥263,678
      propertyTaxBuilding: extractBuildingTax(text),      // ¥16,507
      propertyTaxLand: extractLandTax(text),              // ¥8,683
      managementFee: extractAmountByPosition(text, 13),   // ¥25,010
      managementFeeMonth,                                 // 月数
      stampSales: extractAmountByPosition(text, 14),      // ¥0
      stampLoan: extractAmountByPosition(text, 15),       // ¥20,000
      travelFee: extractAmountByPosition(text, 16),       // ¥0
      loanTransferFee: extractAmountByPosition(text, 17), // ¥0
      certificateFee: extractAmountByPosition(text, 18),  // ¥0
    };

    console.log('\n抽出された金額:');
    console.log(`  ローン事務手数料: ¥${data.loanFee.toLocaleString()}`);
    console.log(`  事業者事務手数料: ¥${data.businessFee.toLocaleString()}`);
    console.log(`  火災保険料: ¥${data.insuranceFee.toLocaleString()}`);
    console.log(`  振込手数料: ¥${data.transferFee.toLocaleString()}`);
    console.log(`  登記費用: ¥${data.registrationFee.toLocaleString()}`);
    console.log(`  固都税精算金（建物）: ¥${data.propertyTaxBuilding.toLocaleString()}`);
    console.log(`  固都税精算金（土地）: ¥${data.propertyTaxLand.toLocaleString()}`);
    console.log(`  管理費・修繕積立金: ¥${data.managementFee.toLocaleString()}`);

    // 合計を計算
    data.total = Object.values(data)
      .filter(v => typeof v === 'number')
      .reduce((sum, v) => sum + v, 0);

    console.log(`決済明細書解析: ${propertyName}`);
    console.log(`  合計: ¥${data.total.toLocaleString()}`);

    return data;
  } catch (error) {
    console.error('決済明細書の解析エラー:', error);
    throw new Error(`決済明細書の解析に失敗しました: ${error.message}`);
  }
}

/**
 * 位置で金額を抽出（「諸費用」セクション内のN番目の金額）
 * @param {string} text - PDFテキスト
 * @param {number} position - 取得する金額の位置（0始まり）
 * @returns {number} 金額
 */
function extractAmountByPosition(text, position) {
  // 「諸費用」セクションから金額を抽出
  const expensesStart = text.indexOf('諸費用');
  if (expensesStart === -1) {
    return 0;
  }

  const expensesSection = text.substring(expensesStart);
  const amounts = expensesSection.match(/¥([0-9,]+)/g);

  if (amounts && position < amounts.length) {
    const amount = parseInt(amounts[position].replace(/¥|,/g, ''), 10);
    if (!isNaN(amount) && amount >= 0) {
      return amount;
    }
  }

  return 0;
}

/**
 * 固都税精算金（建物）を抽出
 * @param {string} text - PDFテキスト
 * @returns {number} 金額
 */
function extractBuildingTax(text) {
  // 「うち建物代金」の行から金額を抽出
  const match = text.match(/うち建物代金[^¥]*¥([0-9,]+)/);
  if (match) {
    return parseInt(match[1].replace(/,/g, ''), 10);
  }
  return 0;
}

/**
 * 固都税精算金（土地）を抽出
 * @param {string} text - PDFテキスト
 * @returns {number} 金額
 */
function extractLandTax(text) {
  // 「うち土地代金」の行から金額を抽出
  const match = text.match(/うち土地代金[^¥]*¥([0-9,]+)/);
  if (match) {
    return parseInt(match[1].replace(/,/g, ''), 10);
  }
  return 0;
}

/**
 * 土地代金を抽出
 * @param {string} text - PDFテキスト
 * @returns {number} 金額
 */
function extractLandPrice(text) {
  // 「土地代金」の行から金額を抽出（「うち土地代金」は除外）
  const match = text.match(/(?<!うち)土地代金[^¥]*¥([0-9,]+)/);
  if (match) {
    return parseInt(match[1].replace(/,/g, ''), 10);
  }
  return 0;
}

/**
 * 建物代金を抽出
 * @param {string} text - PDFテキスト
 * @returns {number} 金額
 */
function extractBuildingPrice(text) {
  // 「建物代金」の行から金額を抽出（「うち建物代金」は除外）
  const match = text.match(/(?<!うち)建物代金[^¥]*¥([0-9,]+)/);
  if (match) {
    return parseInt(match[1].replace(/,/g, ''), 10);
  }
  return 0;
}

module.exports = { parseSettlementPDF };
