const pdfParse = require('pdf-parse');
const fs = require('fs').promises;

/**
 * 譲渡対価証明書PDFから情報を抽出
 * @param {string} pdfPath - PDFファイルのパス
 * @param {string} filename - ファイル名（オリジナル）
 * @param {string} folderPath - フォルダパス（オプション）
 * @returns {Object} 抽出されたデータ
 */
async function parseTransferPDF(pdfPath, filename = '', folderPath = '') {
  try {
    const dataBuffer = await fs.readFile(pdfPath);
    const pdfData = await pdfParse(dataBuffer);
    const text = pdfData.text;

    // デバッグ: PDFテキストの最初の1000文字を出力
    console.log('\n===== 譲渡対価証明書PDFテキスト (最初の1000文字) =====');
    console.log(text.substring(0, 1000));
    console.log('=====================================\n');

    // 物件名を抽出 - ファイル名から取得
    // ファイル名フォーマット: "アスヴェル京都西京極 703_譲渡対価証明書.pdf"
    // → "_" の前の部分が物件名
    let propertyName = '';

    if (filename) {
      const match = filename.match(/^(.+?)_/);
      if (match && match[1]) {
        propertyName = match[1].trim();
        // 「譲渡対価証明書」というテキストが含まれている場合は削除
        propertyName = propertyName.replace(/譲渡対価証明書/g, '').trim();
        console.log(`\nファイル名から物件名を抽出: "${propertyName}" (ファイル名: "${filename}")`);
      }
    }

    // 物件名が取得できない場合、フォルダ名をフォールバックとして使用
    if (!propertyName && folderPath) {
      // フォルダパスから最初のフォルダ名（ユーザーが選択したフォルダ）を取得
      // 例: "中井様20250101/サブフォルダ" → "中井様20250101"
      const pathParts = folderPath.split('/');
      const folderName = pathParts[pathParts.length - 1] || pathParts[0] || '';
      if (folderName) {
        propertyName = folderName;
        console.log(`\nファイル名から物件名を取得できないため、フォルダ名を使用: "${propertyName}" (フォルダパス: "${folderPath}")`);
      }
    }

    // 竣工日を抽出
    // パターン1: "竣工日 2010年2月24日" (スペース区切り)
    // パターン2: "竣工日：2010年2月24日" (コロン)
    // パターン3: "竣工日2010年2月24日" (直接続く)
    let completionDate = '';
    const completionMatch = text.match(/竣工日[：:|\s]*(\d{4})年(\d{1,2})月(\d{1,2})日/);
    if (completionMatch) {
      const year = completionMatch[1];
      const month = completionMatch[2].padStart(2, '0');
      const day = completionMatch[3].padStart(2, '0');
      completionDate = `${year}/${month}/${day}`;
      console.log(`竣工日を抽出: ${completionDate}`);
    } else {
      console.log('⚠ 竣工日が見つかりませんでした');
    }

    // 譲渡日を抽出
    // パターン1: "譲渡日 2024年12月27日" (スペース区切り)
    // パターン2: "譲渡日：2024年12月27日" (コロン)
    // パターン3: "譲渡日2024年12月27日" (直接続く)
    let transferDate = '';
    const transferMatch = text.match(/譲渡日[：:|\s]*(\d{4})年(\d{1,2})月(\d{1,2})日/);
    if (transferMatch) {
      const year = transferMatch[1];
      const month = transferMatch[2].padStart(2, '0');
      const day = transferMatch[3].padStart(2, '0');
      transferDate = `${year}/${month}/${day}`;
      console.log(`譲渡日を抽出: ${transferDate}`);
    } else {
      console.log('⚠ 譲渡日が見つかりませんでした');
    }

    const data = {
      propertyName,
      completionDate,  // 竣工日
      transferDate     // 譲渡日（購入日）
    };

    console.log(`\n譲渡対価証明書解析: ${propertyName}`);
    console.log(`  竣工日: ${completionDate || '(なし)'}`);
    console.log(`  譲渡日: ${transferDate || '(なし)'}`);

    return data;
  } catch (error) {
    console.error('譲渡対価証明書の解析エラー:', error);
    throw new Error(`譲渡対価証明書の解析に失敗しました: ${error.message}`);
  }
}

module.exports = { parseTransferPDF };
