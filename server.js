const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs').promises;
const { parsePDF } = require('./pdfParser');
const { updateExcel } = require('./excelWriter');
const { parseSettlementPDF } = require('./settlementParser');
const { parseTransferPDF } = require('./transferParser');
const { writeNewProperties } = require('./newPropertyWriter');
// Force reload of transferWriter module
delete require.cache[require.resolve('./transferWriter')];
const { writeTransferDates } = require('./transferWriter');
const { copyDepreciationData } = require('./depreciationCopier');

const app = express();
const PORT = 1919;

// JSONボディパーサーを追加
app.use(express.json());

// アップロードディレクトリの設定
// Vercel環境では /tmp ディレクトリを使用、ローカルでは通常のディレクトリを使用
const isVercel = process.env.VERCEL === '1';
const uploadDir = isVercel ? '/tmp/uploads' : path.join(__dirname, 'uploads');
const outputDir = isVercel ? '/tmp/output' : path.join(__dirname, 'output');

// ディレクトリ作成
(async () => {
  try {
    await fs.mkdir(uploadDir, { recursive: true });
    await fs.mkdir(outputDir, { recursive: true });
    console.log(`アップロードディレクトリ: ${uploadDir}`);
    console.log(`出力ディレクトリ: ${outputDir}`);
  } catch (err) {
    console.error('ディレクトリ作成エラー:', err);
  }
})();

// Multerの設定
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, uniqueSuffix + '-' + Buffer.from(file.originalname, 'latin1').toString('utf8'));
  }
});

const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.pdf' || ext === '.xlsx') {
      cb(null, true);
    } else {
      cb(new Error('PDFファイルまたはExcelファイルのみアップロード可能です'));
    }
  }
});

// 静的ファイルの提供
app.use(express.static(path.join(__dirname, 'public')));
app.use('/output', express.static(outputDir));

// ルート
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// 静的ファイル用の明示的なルート（Vercel用）
app.get('/style.css', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'style.css'));
});

app.get('/script.js', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'script.js'));
});

// PDFアップロードとExcel更新エンドポイント
app.post('/upload', upload.fields([
  { name: 'pdfs', maxCount: 50 },
  { name: 'excel', maxCount: 1 },
  { name: 'settlements', maxCount: 50 },
  { name: 'transfers', maxCount: 50 }
]), async (req, res) => {
  try {
    const pdfFiles = req.files['pdfs'];
    const excelFile = req.files['excel'] ? req.files['excel'][0] : null;
    const settlementFiles = req.files['settlements'] || [];
    const transferFiles = req.files['transfers'] || [];

    // フォルダパス情報を取得
    const pdfPathsMap = req.body.pdfPaths ? JSON.parse(req.body.pdfPaths) : {};
    const settlementPathsMap = req.body.settlementPaths ? JSON.parse(req.body.settlementPaths) : {};
    const transferPathsMap = req.body.transferPaths ? JSON.parse(req.body.transferPaths) : {};

    // デバッグ: フォルダパス情報を出力
    console.log('\n===== フロントエンドから受信したフォルダパス情報 =====');
    console.log('pdfPathsMap:', JSON.stringify(pdfPathsMap, null, 2));
    console.log('settlementPathsMap:', JSON.stringify(settlementPathsMap, null, 2));
    console.log('transferPathsMap:', JSON.stringify(transferPathsMap, null, 2));
    console.log('================================================\n');

    // デバッグ: サーバー側で受信したファイル名を出力
    console.log('\n===== サーバー側で受信したPDFファイル名 =====');
    if (pdfFiles) {
      pdfFiles.forEach(file => {
        console.log(`originalname: "${file.originalname}"`);
      });
    }
    console.log('================================================\n');

    if (!pdfFiles || pdfFiles.length === 0) {
      return res.status(400).json({ error: 'PDFファイルがアップロードされていません' });
    }

    if (!excelFile) {
      return res.status(400).json({ error: 'Excelファイルがアップロードされていません' });
    }

    console.log(`処理開始: ${pdfFiles.length}件のPDFファイル`);
    if (settlementFiles.length > 0) {
      console.log(`決済明細書: ${settlementFiles.length}件`);
    }
    if (transferFiles.length > 0) {
      console.log(`譲渡対価証明書: ${transferFiles.length}件`);
    }

    // 項目マッピングを読み込み
    const mappingPath = path.join(__dirname, 'item-mapping.json');
    let itemMapping = {};
    try {
      const mappingData = await fs.readFile(mappingPath, 'utf8');
      itemMapping = JSON.parse(mappingData);
    } catch (error) {
      console.log('item-mapping.jsonが見つかりません。デフォルトマッピングを使用します。');
      itemMapping = {
        '管理手数料': 'B',
        '宣伝広告費': 'C',
        '設備交換費': 'D'
      };
    }

    // 各PDFファイルを解析
    const pdfDataArray = [];
    const parseErrors = [];
    const unknownItems = new Set(); // 未知の項目を収集

    for (const pdfFile of pdfFiles) {
      try {
        const pdfBuffer = await fs.readFile(pdfFile.path);
        const pdfData = await parsePDF(pdfBuffer);
        const filename = Buffer.from(pdfFile.originalname, 'latin1').toString('utf8');

        // 未知の項目をチェック
        if (pdfData.otherExpenseItems && Object.keys(pdfData.otherExpenseItems).length > 0) {
          Object.keys(pdfData.otherExpenseItems).forEach(itemName => {
            if (!itemMapping[itemName]) {
              unknownItems.add(itemName);
            }
          });
        }

        // originalname を UTF-8 でデコード (filename と同じ処理)
        const decodedOriginalName = Buffer.from(pdfFile.originalname, 'latin1').toString('utf8');
        const folderPath = pdfPathsMap[decodedOriginalName] || '';
        pdfDataArray.push({
          ...pdfData,
          filename: filename,
          folderPath: folderPath // フォルダパスを追加
        });

        console.log(`✓ ${filename} - ${pdfData.propertyName} [フォルダ: "${folderPath}"]`);
      } catch (error) {
        const filename = Buffer.from(pdfFile.originalname, 'latin1').toString('utf8');
        parseErrors.push({
          filename,
          error: error.message
        });
        console.error(`✗ ${filename}:`, error.message);
      }
    }

    if (pdfDataArray.length === 0) {
      return res.status(400).json({
        error: 'PDFファイルの解析に失敗しました',
        parseErrors
      });
    }

    // 未知の項目がある場合は、ユーザーに質問
    if (unknownItems.size > 0) {
      console.log(`未知の項目が見つかりました: ${Array.from(unknownItems).join(', ')}`);

      // 一時データを保存（セッション代わり）
      const tempId = Date.now().toString();
      const tempData = {
        pdfDataArray,
        excelPath,
        pdfFiles: pdfFiles.map(f => f.path),
        parseErrors
      };

      await fs.writeFile(
        path.join(__dirname, `temp-${tempId}.json`),
        JSON.stringify(tempData)
      );

      return res.json({
        needsMapping: true,
        unknownItems: Array.from(unknownItems),
        tempId,
        message: '未知の支払項目が見つかりました。分類を選択してください。'
      });
    }

    // 決済明細書PDFを先に解析
    let propertiesData = [];
    const settlementParseErrors = [];

    if (settlementFiles.length > 0) {
      console.log(`\n===== 決済明細書PDFの解析 =====`);

      for (const settlementFile of settlementFiles) {
        try {
          const originalFilename = Buffer.from(settlementFile.originalname, 'latin1').toString('utf8');
          // originalname を UTF-8 でデコード (filename と同じ処理)
          const decodedOriginalName = Buffer.from(settlementFile.originalname, 'latin1').toString('utf8');
          const folderPath = settlementPathsMap[decodedOriginalName] || '';

          // parseSettlementPDFにフォルダパスも渡す
          const data = await parseSettlementPDF(settlementFile.path, originalFilename, folderPath);
          propertiesData.push(data);
          console.log(`✓ ${originalFilename} [フォルダ: "${folderPath}"]`);
        } catch (error) {
          const filename = Buffer.from(settlementFile.originalname, 'latin1').toString('utf8');
          settlementParseErrors.push({
            filename,
            error: error.message
          });
          console.error(`✗ ${filename}:`, error.message);
        }
      }
    }

    // 譲渡対価証明書PDFを解析
    let transferData = [];
    const transferParseErrors = [];

    if (transferFiles.length > 0) {
      console.log(`\n===== 譲渡対価証明書PDFの解析 =====`);

      for (const transferFile of transferFiles) {
        try {
          const originalFilename = Buffer.from(transferFile.originalname, 'latin1').toString('utf8');
          // originalname を UTF-8 でデコード (filename と同じ処理)
          const decodedOriginalName = Buffer.from(transferFile.originalname, 'latin1').toString('utf8');
          const folderPath = transferPathsMap[decodedOriginalName] || '';

          // parseTransferPDFにフォルダパスも渡す
          const data = await parseTransferPDF(transferFile.path, originalFilename, folderPath);
          transferData.push({
            ...data,
            folderPath: folderPath // フォルダパスを追加
          });
          console.log(`✓ ${originalFilename} [フォルダ: "${folderPath}"]`);
        } catch (error) {
          const filename = Buffer.from(transferFile.originalname, 'latin1').toString('utf8');
          transferParseErrors.push({
            filename,
            error: error.message
          });
          console.error(`✗ ${filename}:`, error.message);
        }
      }
    }

    // 決済明細書のファイル名を抽出（updateExcelに渡す）
    const settlementFileNames = settlementFiles.map(file =>
      Buffer.from(file.originalname, 'latin1').toString('utf8')
    );

    // フォルダ名を抽出（最初のPDFファイルのフォルダパスから取得）
    let folderName = '';
    if (pdfDataArray.length > 0 && pdfDataArray[0].folderPath) {
      const folderPath = pdfDataArray[0].folderPath;
      // フォルダパスから最初のフォルダ名（ユーザーが選択したフォルダ）を取得
      // 例: "中井様20250101/サブフォルダ" → "中井様20250101"
      const pathParts = folderPath.split('/');
      folderName = pathParts[0] || '';
    }
    console.log(`\n使用するフォルダ名: "${folderName}"`);

    // Excelファイルを更新（マッピング情報、決済明細書ファイル名、決済明細書データ、フォルダ名を渡す）
    const excelPath = excelFile.path;
    const result = await updateExcel(excelPath, pdfDataArray, itemMapping, settlementFileNames, propertiesData, folderName);

    let finalOutputPath = result.outputPath;

    // 決済明細書PDFがある場合、【不】新規不動産シートにも書き込み
    if (propertiesData.length > 0) {
      console.log(`\n===== 【不】新規不動産シートへの書き込み =====`);
      const newPropertyResult = await writeNewProperties(finalOutputPath, propertiesData, pdfDataArray.length);
      finalOutputPath = newPropertyResult.outputPath;
      console.log(`【不】新規不動産シートへの書き込み完了: ${propertiesData.length}件`);
    }

    // 譲渡対価証明書PDFがある場合、【不】④耐用年数シートにも書き込み
    if (transferData.length > 0) {
      const transferResult = await writeTransferDates(finalOutputPath, transferData, pdfDataArray);
      finalOutputPath = transferResult.outputPath;
      console.log(`【不】④耐用年数シートへの書き込み完了: ${transferResult.updateCount}件`);
    }

    // 全ての処理が完了した後、【不】④耐用年数 → 【不】③減価償却(JDLよりエクスポート)にデータを転記
    console.log(`\n===== 【不】④耐用年数 → 【不】③減価償却(JDLよりエクスポート) データ転記 =====`);
    const depreciationResult = await copyDepreciationData(finalOutputPath);
    finalOutputPath = depreciationResult.outputPath;
    console.log(`【不】③減価償却(JDLよりエクスポート)シートへの転記完了: ${depreciationResult.copyCount}件`);

    // アップロードされたPDFファイルを削除
    for (const pdfFile of pdfFiles) {
      await fs.unlink(pdfFile.path).catch(err => console.error('ファイル削除エラー:', err));
    }
    // 決済明細書PDFを削除
    for (const settlementFile of settlementFiles) {
      await fs.unlink(settlementFile.path).catch(err => console.error('ファイル削除エラー:', err));
    }
    // 譲渡対価証明書PDFを削除
    for (const transferFile of transferFiles) {
      await fs.unlink(transferFile.path).catch(err => console.error('ファイル削除エラー:', err));
    }
    // 元のExcelファイルを削除
    await fs.unlink(excelPath).catch(err => console.error('ファイル削除エラー:', err));

    // 出力ファイルをoutputディレクトリに移動
    const outputFilename = path.basename(finalOutputPath);
    const finalOutputPathInOutput = path.join(outputDir, outputFilename);
    await fs.rename(finalOutputPath, finalOutputPathInOutput);

    console.log(`処理完了: ${finalOutputPathInOutput}`);

    res.json({
      success: true,
      message: `${pdfDataArray.length}件のPDFを処理しました`,
      outputPath: finalOutputPathInOutput, // フロントエンドに出力パスを返す
      downloadUrl: `/output/${outputFilename}`,
      results: result.results,
      parseErrors: parseErrors.length > 0 ? parseErrors : undefined
    });

  } catch (error) {
    console.error('エラー:', error);
    res.status(500).json({
      error: '処理中にエラーが発生しました',
      message: error.message
    });
  }
});

// 項目分類を保存して処理を続行するエンドポイント
app.post('/save-mapping', async (req, res) => {
  try {
    const { tempId, mapping } = req.body;

    if (!tempId || !mapping) {
      return res.status(400).json({ error: 'tempIdとmappingが必要です' });
    }

    // 一時データを読み込み
    const tempPath = path.join(__dirname, `temp-${tempId}.json`);
    const tempDataStr = await fs.readFile(tempPath, 'utf8');
    const tempData = JSON.parse(tempDataStr);

    // 項目マッピングを更新
    const mappingPath = path.join(__dirname, 'item-mapping.json');
    let itemMapping = {};
    try {
      const mappingData = await fs.readFile(mappingPath, 'utf8');
      itemMapping = JSON.parse(mappingData);
    } catch (error) {
      itemMapping = {
        '管理手数料': 'B',
        '宣伝広告費': 'C',
        '設備交換費': 'D'
      };
    }

    // 新しいマッピングを追加
    Object.assign(itemMapping, mapping);

    // 保存
    await fs.writeFile(mappingPath, JSON.stringify(itemMapping, null, 2));
    console.log('項目マッピングを更新しました:', mapping);

    // Excelファイルを更新（マッピング情報を渡す）
    const result = await updateExcel(tempData.excelPath, tempData.pdfDataArray, itemMapping);

    // 一時ファイルとアップロードファイルを削除
    await fs.unlink(tempPath).catch(err => console.error('temp削除エラー:', err));
    for (const pdfPath of tempData.pdfFiles) {
      await fs.unlink(pdfPath).catch(err => console.error('PDF削除エラー:', err));
    }
    await fs.unlink(tempData.excelPath).catch(err => console.error('Excel削除エラー:', err));

    // 出力ファイルをoutputディレクトリに移動
    const outputFilename = path.basename(result.outputPath);
    const finalOutputPath = path.join(outputDir, outputFilename);
    await fs.rename(result.outputPath, finalOutputPath);

    console.log(`処理完了: ${finalOutputPath}`);

    res.json({
      success: true,
      message: `${tempData.pdfDataArray.length}件のPDFを処理しました`,
      downloadUrl: `/output/${outputFilename}`,
      results: result.results,
      parseErrors: tempData.parseErrors.length > 0 ? tempData.parseErrors : undefined
    });

  } catch (error) {
    console.error('エラー:', error);
    res.status(500).json({
      error: '処理中にエラーが発生しました',
      message: error.message
    });
  }
});

// 新規物件アップロードエンドポイント
app.post('/upload-new-properties', upload.fields([
  { name: 'settlements', maxCount: 50 },
  { name: 'excel', maxCount: 1 }
]), async (req, res) => {
  try {
    const settlementFiles = req.files['settlements'];
    const excelFile = req.files['excel'] ? req.files['excel'][0] : null;

    // フロントエンドから既存のExcelパスを受け取る（年間収支一覧表処理済みの場合）
    const existingExcelPath = req.body.existingExcelPath;

    if (!settlementFiles || settlementFiles.length === 0) {
      return res.status(400).json({ error: '決済明細書PDFがアップロードされていません' });
    }

    // existingExcelPathがある場合はそれを使用、なければアップロードされたExcelを使用
    let excelPath;
    if (existingExcelPath) {
      excelPath = existingExcelPath;
      console.log(`\n既存のExcelファイルを使用: ${excelPath}`);
    } else if (excelFile) {
      excelPath = excelFile.path;
    } else {
      return res.status(400).json({ error: 'Excelファイルがアップロードされていません' });
    }

    console.log(`\n新規物件処理開始: ${settlementFiles.length}件の決済明細書`);

    // 各決済明細書PDFを解析
    const propertiesData = [];
    const parseErrors = [];

    for (const settlementFile of settlementFiles) {
      try {
        const originalFilename = Buffer.from(settlementFile.originalname, 'latin1').toString('utf8');
        const data = await parseSettlementPDF(settlementFile.path, originalFilename);
        propertiesData.push(data);
        console.log(`✓ ${originalFilename}`);
      } catch (error) {
        const filename = Buffer.from(settlementFile.originalname, 'latin1').toString('utf8');
        parseErrors.push({
          filename,
          error: error.message
        });
        console.error(`✗ ${filename}:`, error.message);
      }
    }

    if (propertiesData.length === 0) {
      return res.status(400).json({
        error: '決済明細書の解析に失敗しました',
        parseErrors
      });
    }

    // Excelファイルを更新
    const result = await writeNewProperties(excelPath, propertiesData);

    // アップロードされたファイルを削除
    for (const settlementFile of settlementFiles) {
      await fs.unlink(settlementFile.path).catch(err => console.error('ファイル削除エラー:', err));
    }
    // 新規アップロードのExcelファイルの場合のみ削除（既存パスの場合は削除しない）
    if (!existingExcelPath && excelFile) {
      await fs.unlink(excelPath).catch(err => console.error('ファイル削除エラー:', err));
    }

    // 出力ファイルをoutputディレクトリに移動
    const outputFilename = path.basename(result.outputPath);
    const finalOutputPath = path.join(outputDir, outputFilename);
    await fs.rename(result.outputPath, finalOutputPath);

    console.log(`新規物件処理完了: ${finalOutputPath}\n`);

    res.json({
      success: true,
      message: `${propertiesData.length}件の新規物件を処理しました`,
      downloadUrl: `/output/${outputFilename}`,
      propertyCount: result.propertyCount,
      parseErrors: parseErrors.length > 0 ? parseErrors : undefined
    });

  } catch (error) {
    console.error('エラー:', error);
    res.status(500).json({
      error: '処理中にエラーが発生しました',
      message: error.message
    });
  }
});

// サーバー起動（ローカル環境のみ）
if (process.env.NODE_ENV !== 'production') {
  app.listen(PORT, () => {
    console.log(`\n========================================`);
    console.log(`中井ソリューションズ`);
    console.log(`========================================`);
    console.log(`サーバーが起動しました`);
    console.log(`URL: http://localhost:${PORT}`);
    console.log(`========================================\n`);
  });
}

// Vercel用にappをエクスポート
module.exports = app;
