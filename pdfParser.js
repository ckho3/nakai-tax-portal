const pdfParse = require('pdf-parse');

/**
 * PDFから不動産収支データを抽出する
 * @param {Buffer} pdfBuffer - PDFファイルのバッファ
 * @returns {Promise<Object>} 抽出されたデータ
 */
async function parsePDF(pdfBuffer) {
  try {
    const data = await pdfParse(pdfBuffer);
    const text = data.text;

    // 物件名を抽出
    const propertyNameMatch = text.match(/管理物件名\s*(.+?)(?:\n|物件所在地)/m);
    const propertyName = propertyNameMatch ? propertyNameMatch[1].trim() : null;

    // 年度を抽出
    const yearMatch = text.match(/(\d{4})年度/);
    const year = yearMatch ? yearMatch[1] : '2024';

    // 行ごとに分割
    const lines = text.split('\n');

    // 賃料行を抽出（収入セクションの賃料のみ）
    let rentLine = null;
    let managementFeeLine = null;
    let equipmentLine = null;
    let advertisingLine = null;
    const otherExpenseItems = {}; // その他の支払項目を格納

    // 支払セクションを検出するためのフラグ
    let inExpenseSection = false;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];

      // 収入セクションの「収入合計①」または「賃料」行（数値が複数ある行）
      if (line.includes('収入合計') || line.includes('収⼊合計') || line.startsWith('賃料')) {
        // デバッグ: 賃料行候補を出力
        const amounts = line.match(/(\d{1,3}(?:,\d{3})*円|--)/g);
        const amountCount = amounts ? amounts.length : 0;

        if (amountCount >= 12) {
          // 12個以上の金額があればrentLineとして採用
          rentLine = line;
          console.log(`  賃料行検出: ${amountCount}個の金額を発見`);
        } else if (amountCount > 0) {
          // 12個未満だが金額がある場合はログ出力
          console.log(`  [警告] 賃料行候補（金額不足）: ${amountCount}個の金額（12個以上必要）`);
          console.log(`    内容: ${line.substring(0, 100)}...`);
        }
      }

      // 「支払」セクションの開始を検出
      if (line.includes('支払') || line.includes('【B】')) {
        inExpenseSection = true;
      }

      // 支払セクション内の項目を検出
      if (inExpenseSection) {
        // 既知の項目
        if (line.startsWith('管理手数料') && line.match(/(\d{1,3},?\d{3}円.*){10,}/)) {
          managementFeeLine = line;
        } else if (line.startsWith('設備交換費') || (line.includes('設備交換費') && line.includes('円'))) {
          equipmentLine = line;
        } else if (line.startsWith('宣伝広告費') || (line.includes('宣伝広告費') && line.includes('円'))) {
          advertisingLine = line;
        } else {
          // その他の支払項目を検出（月別データがある行）
          const itemMatch = line.match(/^([^\s]+)\s+.*(\d{1,3}(?:,\d{3})*円|--)/);
          if (itemMatch && line.match(/(\d{1,3}(?:,\d{3})*円|--)/g)) {
            const itemName = itemMatch[1];
            // 既知の項目でない場合のみ追加
            if (itemName !== '管理手数料' && itemName !== '設備交換費' && itemName !== '宣伝広告費' && itemName !== '賃料') {
              const amounts = line.match(/(\d{1,3}(?:,\d{3})*円|--)/g);
              if (amounts && amounts.length >= 12) {
                otherExpenseItems[itemName] = line;
              }
            }
          }
        }
      }
    }

    // 月別賃料を抽出
    const monthlyRents = [];
    if (rentLine) {
      // 金額または "--" を抽出（ブランクの月も考慮）
      const amounts = rentLine.match(/(\d{1,3}(?:,\d{3})*円|--)/g);
      if (amounts && amounts.length >= 24) {
        // PDFは2回ずつ重複しているので1個飛ばしで取得 (0, 2, 4, 6, ...)
        for (let i = 0; i < 12; i++) {
          const item = amounts[i * 2];
          if (item === '--') {
            monthlyRents.push(0); // ブランク月は0として扱う
          } else {
            const value = item.replace(/[,円]/g, '');
            monthlyRents.push(parseInt(value));
          }
        }
      } else if (amounts && amounts.length >= 12) {
        // データが12〜23個の場合は最初から12個取得
        for (let i = 0; i < 12; i++) {
          const item = amounts[i];
          if (item === '--') {
            monthlyRents.push(0); // ブランク月は0として扱う
          } else {
            const value = item.replace(/[,円]/g, '');
            monthlyRents.push(parseInt(value));
          }
        }
      }
    }

    // 年間合計賃料を計算
    const totalRent = monthlyRents.length > 0
      ? monthlyRents.reduce((sum, val) => sum + (val || 0), 0)
      : null;

    // 管理手数料を抽出（PDFは2回ずつ重複しているので1個飛ばしで取得）
    const managementFees = [];
    if (managementFeeLine) {
      const amounts = managementFeeLine.match(/(\d{1,3}(?:,\d{3})*円|--)/g);
      if (amounts && amounts.length >= 24) {
        // 1個飛ばしで12個取得 (0, 2, 4, 6, ...)
        for (let i = 0; i < 12; i++) {
          const item = amounts[i * 2];
          if (item === '--') {
            managementFees.push(0);
          } else {
            const value = item.replace(/[,円]/g, '');
            managementFees.push(parseInt(value));
          }
        }
      }
    }

    // 設備交換費を抽出（PDFは2回ずつ重複しているので1個飛ばしで取得）
    const equipmentCosts = [];
    if (equipmentLine) {
      const amounts = equipmentLine.match(/(\d{1,3}(?:,\d{3})*円|--)/g);
      if (amounts && amounts.length >= 24) {
        // 1個飛ばしで12個取得 (0, 2, 4, 6, ...)
        for (let i = 0; i < 12; i++) {
          const item = amounts[i * 2];
          if (item === '--') {
            equipmentCosts.push(0);
          } else {
            const value = item.replace(/[,円]/g, '');
            equipmentCosts.push(parseInt(value));
          }
        }
      } else {
        // データがない場合は0で埋める
        for (let i = 0; i < 12; i++) {
          equipmentCosts.push(0);
        }
      }
    } else {
      // 行が見つからない場合は0で埋める
      for (let i = 0; i < 12; i++) {
        equipmentCosts.push(0);
      }
    }

    // 宣伝広告費を抽出（PDFは2回ずつ重複しているので1個飛ばしで取得）
    const advertisingCosts = [];
    if (advertisingLine) {
      const amounts = advertisingLine.match(/(\d{1,3}(?:,\d{3})*円|--)/g);
      if (amounts && amounts.length >= 12) {
        // データが24個以上ある場合は1個飛ばしで取得
        if (amounts.length >= 24) {
          for (let i = 0; i < 12; i++) {
            const item = amounts[i * 2];
            if (item === '--') {
              advertisingCosts.push(0);
            } else {
              const value = item.replace(/[,円]/g, '');
              advertisingCosts.push(parseInt(value));
            }
          }
        } else {
          // データが少ない場合は最初から12個取得
          for (let i = 0; i < 12; i++) {
            const item = amounts[i];
            if (item === '--') {
              advertisingCosts.push(0);
            } else {
              const value = item.replace(/[,円]/g, '');
              advertisingCosts.push(parseInt(value));
            }
          }
        }
      } else {
        for (let i = 0; i < 12; i++) {
          advertisingCosts.push(0);
        }
      }
    } else {
      for (let i = 0; i < 12; i++) {
        advertisingCosts.push(0);
      }
    }

    // 物件所在地を抽出
    const addressMatch = text.match(/物件所在地\s*〒?\s*\d{3}-?\d{4}\s*(.+?)(?:\n|専有面積)/m);
    const propertyAddress = addressMatch ? addressMatch[1].trim() : null;

    // 賃借人名を抽出（最新の契約者）
    const tenantMatches = text.match(/賃借人名(.+?)賃貸期間/g);
    let tenantName = null;
    if (tenantMatches && tenantMatches.length > 0) {
      // 最後の賃借人（最新の契約）を取得
      const lastTenant = tenantMatches[tenantMatches.length - 1];
      const nameMatch = lastTenant.match(/賃借人名(.+?)賃貸期間/);
      if (nameMatch) {
        tenantName = nameMatch[1].trim();
      }
    }

    // 契約期間を抽出（最新の契約）
    const contractMatches = text.match(/賃貸期間(\d{4}-\d{2}-\d{2})〜(\d{4}-\d{2}-\d{2})/g);
    let contractStartDate = null;
    let contractEndDate = null;
    if (contractMatches && contractMatches.length > 0) {
      // 最後の契約期間（最新の契約）を取得
      const lastContract = contractMatches[contractMatches.length - 1];
      const dateMatch = lastContract.match(/賃貸期間(\d{4}-\d{2}-\d{2})〜(\d{4}-\d{2}-\d{2})/);
      if (dateMatch) {
        contractStartDate = dateMatch[1];
        contractEndDate = dateMatch[2];
      }
    }

    // 専有面積を抽出
    const areaMatch = text.match(/専有面積\s*([\d.]+)\s*m²/);
    const rentalArea = areaMatch ? parseFloat(areaMatch[1]) : null;

    // その他の支払項目を月別データに変換
    const otherExpenseData = {};
    for (const [itemName, line] of Object.entries(otherExpenseItems)) {
      const amounts = line.match(/(\d{1,3}(?:,\d{3})*円|--)/g);
      const monthlyData = [];

      if (amounts && amounts.length >= 12) {
        // データが24個以上ある場合は1個飛ばしで取得（重複パターン）
        if (amounts.length >= 24) {
          for (let i = 0; i < 12; i++) {
            const item = amounts[i * 2];
            if (item === '--') {
              monthlyData.push(0);
            } else {
              const value = item.replace(/[,円]/g, '');
              monthlyData.push(parseInt(value));
            }
          }
        } else {
          // データが12〜23個の場合は最初から12個取得
          for (let i = 0; i < 12; i++) {
            const item = amounts[i];
            if (item === '--') {
              monthlyData.push(0);
            } else {
              const value = item.replace(/[,円]/g, '');
              monthlyData.push(parseInt(value));
            }
          }
        }
      }

      otherExpenseData[itemName] = monthlyData;
    }

    console.log(`Parsed: ${propertyName} - Rents: ${monthlyRents.length} items`);

    // Rentsが0の場合、詳細なデバッグ情報を出力
    if (monthlyRents.length === 0) {
      console.log(`  [デバッグ] 賃料行が見つかりませんでした`);
      console.log(`  rentLine: ${rentLine ? 'あり' : 'なし'}`);
      if (rentLine) {
        console.log(`  rentLineの内容: ${rentLine.substring(0, 150)}`);
      }

      // 収入合計を含む全ての行を表示
      console.log(`  収入関連の行を検索中...`);
      const lines = text.split('\n');
      for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes('収入') && (lines[i].includes('円') || lines[i].includes('--'))) {
          console.log(`    Line ${i}: ${lines[i].substring(0, 100)}`);
        }
      }
    }

    if (Object.keys(otherExpenseData).length > 0) {
      console.log(`Other expense items found: ${Object.keys(otherExpenseData).join(', ')}`);
    }

    return {
      propertyName,
      year,
      monthlyRents,
      totalRent,
      managementFees,
      equipmentCosts,
      advertisingCosts,
      otherExpenseItems: otherExpenseData, // その他の支払項目（月別データ）
      // 追加情報
      propertyAddress,
      tenantName,
      contractStartDate,
      contractEndDate,
      rentalArea,
      rawText: text // デバッグ用
    };

  } catch (error) {
    console.error('PDF parsing error:', error);
    throw new Error('PDFの解析に失敗しました: ' + error.message);
  }
}

module.exports = { parsePDF };
