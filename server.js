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
const { JobManager } = require('./jobManager');

const app = express();
const jobManager = new JobManager();
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

// ========================================
// 非同期処理エンドポイント
// ========================================

// 非同期アップロードエンドポイント（即座にjobIdを返す）
app.post('/upload-async', upload.fields([
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

    if (!pdfFiles || pdfFiles.length === 0) {
      return res.status(400).json({ error: 'PDFファイルがアップロードされていません' });
    }

    if (!excelFile) {
      return res.status(400).json({ error: 'Excelファイルがアップロードされていません' });
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

    // ジョブを作成
    const job = await jobManager.createJob({
      pdfFiles: pdfFiles,
      excelPath: excelFile.path,
      settlementFiles: settlementFiles,
      transferFiles: transferFiles,
      pdfPathsMap,
      settlementPathsMap,
      transferPathsMap,
      itemMapping
    });

    console.log(`非同期ジョブ作成: ${job.id}`);
    console.log(`- PDFファイル: ${pdfFiles.length}件`);
    console.log(`- 決済明細書: ${settlementFiles.length}件`);
    console.log(`- 譲渡対価証明書: ${transferFiles.length}件`);

    // 即座にjobIdを返す（フロントエンドが /process-next を繰り返し呼び出す）
    res.json({
      success: true,
      jobId: job.id,
      message: 'ジョブを受け付けました。処理を開始してください。'
    });
  } catch (error) {
    console.error('非同期アップロードエラー:', error);
    res.status(500).json({
      error: 'アップロードエラーが発生しました',
      message: error.message
    });
  }
});

// ジョブステータス確認エンドポイント
app.get('/job-status/:jobId', async (req, res) => {
  try {
    const { jobId } = req.params;
    const job = await jobManager.getJob(jobId);

    if (!job) {
      return res.status(404).json({
        error: 'ジョブが見つかりません',
        jobId
      });
    }

    res.json({
      jobId: job.id,
      status: job.status,
      progress: job.progress,
      message: job.message,
      error: job.error,
      createdAt: job.createdAt,
      updatedAt: job.updatedAt
    });
  } catch (error) {
    console.error('ジョブステータス取得エラー:', error);
    res.status(500).json({
      error: 'ステータス取得エラー',
      message: error.message
    });
  }
});

// 結果ダウンロードエンドポイント
app.get('/download/:jobId', async (req, res) => {
  try {
    const { jobId } = req.params;
    const job = await jobManager.getJob(jobId);

    if (!job) {
      return res.status(404).json({
        error: 'ジョブが見つかりません',
        jobId
      });
    }

    if (job.status !== 'completed') {
      return res.status(400).json({
        error: 'ジョブがまだ完了していません',
        status: job.status,
        progress: job.progress
      });
    }

    if (!job.resultPath) {
      return res.status(404).json({
        error: '結果ファイルが見つかりません'
      });
    }

    // ファイルが存在するか確認
    try {
      await fs.access(job.resultPath);
    } catch (error) {
      return res.status(404).json({
        error: '結果ファイルが見つかりません',
        message: 'ファイルが削除された可能性があります'
      });
    }

    // ファイル名を取得
    const filename = path.basename(job.resultPath);

    res.download(job.resultPath, filename, (error) => {
      if (error) {
        console.error('ダウンロードエラー:', error);
        if (!res.headersSent) {
          res.status(500).json({
            error: 'ダウンロードエラー',
            message: error.message
          });
        }
      }
    });
  } catch (error) {
    console.error('ダウンロードエラー:', error);
    res.status(500).json({
      error: 'ダウンロードエラー',
      message: error.message
    });
  }
});

// 1件ずつ処理するエンドポイント（Vercel対応）
app.post('/process-next', async (req, res) => {
  try {
    const { jobId } = req.body;

    if (!jobId) {
      return res.status(400).json({ error: 'jobId is required' });
    }

    // ジョブ情報を取得
    const job = await jobManager.getJob(jobId);
    if (!job) {
      return res.status(404).json({ error: 'Job not found' });
    }

    // 処理状態チェック
    if (job.status === 'completed') {
      return res.status(200).json({
        success: true,
        completed: true,
        message: 'Job already completed'
      });
    }

    if (job.status === 'failed') {
      return res.status(400).json({ error: 'Job has failed' });
    }

    const {
      pdfFiles,
      settlementFiles = [],
      transferFiles = [],
      pdfPathsMap = {},
      settlementPathsMap = {},
      transferPathsMap = {},
      itemMapping = {}
    } = job.data;

    // 現在の処理インデックス
    const currentPdfIndex = job.progress || 0;
    const currentSettlementIndex = job.settlementProgress || 0;
    const currentTransferIndex = job.transferProgress || 0;

    // 初期化
    if (!job.pdfResults) {
      await jobManager.updateJob(jobId, {
        status: 'processing',
        pdfResults: [],
        settlementResults: [],
        transferResults: [],
        parseErrors: []
      });
    }

    // Step 1: 年間収支一覧表PDFを1件処理
    if (currentPdfIndex < pdfFiles.length) {
      const pdfFile = pdfFiles[currentPdfIndex];

      try {
        // pdfFile.pathが文字列の場合はそのまま、Bufferの場合はエラー
        let pdfBuffer;
        if (typeof pdfFile.path === 'string') {
          pdfBuffer = await fs.readFile(pdfFile.path);
        } else if (Buffer.isBuffer(pdfFile.buffer)) {
          // Bufferが直接保存されている場合
          pdfBuffer = pdfFile.buffer;
        } else {
          throw new Error('Invalid PDF file format in job data');
        }

        const pdfData = await parsePDF(pdfBuffer);
        const filename = Buffer.from(pdfFile.originalname, 'latin1').toString('utf8');
        const decodedOriginalName = Buffer.from(pdfFile.originalname, 'latin1').toString('utf8');
        const folderPath = pdfPathsMap[decodedOriginalName] || '';

        const pdfResults = job.pdfResults || [];
        pdfResults.push({
          ...pdfData,
          filename: filename,
          folderPath: folderPath
        });

        const progressPercent = Math.round(((currentPdfIndex + 1) / pdfFiles.length) * 30);

        await jobManager.updateJob(jobId, {
          progress: currentPdfIndex + 1,
          pdfResults,
          message: `年間収支一覧表を解析中 (${currentPdfIndex + 1}/${pdfFiles.length})...`
        });

        return res.json({
          success: true,
          completed: false,
          nextStep: 'pdf',
          progress: {
            current: currentPdfIndex + 1,
            total: pdfFiles.length + settlementFiles.length + transferFiles.length,
            percentage: progressPercent,
            message: `年間収支一覧表を解析中 (${currentPdfIndex + 1}/${pdfFiles.length})...`
          }
        });

      } catch (error) {
        const filename = Buffer.from(pdfFile.originalname, 'latin1').toString('utf8');
        const parseErrors = job.parseErrors || [];
        parseErrors.push({
          filename,
          error: error.message
        });

        await jobManager.updateJob(jobId, {
          progress: currentPdfIndex + 1,
          parseErrors
        });

        return res.json({
          success: true,
          completed: false,
          nextStep: 'pdf',
          error: `${filename}: ${error.message}`,
          progress: {
            current: currentPdfIndex + 1,
            total: pdfFiles.length + settlementFiles.length + transferFiles.length,
            percentage: Math.round(((currentPdfIndex + 1) / pdfFiles.length) * 30),
            message: `年間収支一覧表を解析中 (${currentPdfIndex + 1}/${pdfFiles.length})...`
          }
        });
      }
    }

    // Step 2: 決済明細書PDFを1件処理
    if (currentSettlementIndex < settlementFiles.length) {
      const settlementFile = settlementFiles[currentSettlementIndex];

      try {
        const originalFilename = Buffer.from(settlementFile.originalname, 'latin1').toString('utf8');
        const decodedOriginalName = Buffer.from(settlementFile.originalname, 'latin1').toString('utf8');
        const folderPath = settlementPathsMap[decodedOriginalName] || '';

        // ファイルパスを取得（pathがない場合は一時ファイルを作成）
        let settlementPath;
        let isTempFile = false;

        if (typeof settlementFile.path === 'string') {
          settlementPath = settlementFile.path;
        } else if (Buffer.isBuffer(settlementFile.buffer)) {
          // Bufferの場合は一時ファイルを作成
          const tempPath = path.join(uploadDir, `temp_settlement_${Date.now()}.pdf`);
          await fs.writeFile(tempPath, settlementFile.buffer);
          settlementPath = tempPath;
          isTempFile = true;
        } else {
          throw new Error('Invalid settlement PDF file format in job data');
        }

        const propertyData = await parseSettlementPDF(
          settlementPath,
          originalFilename,
          folderPath
        );

        // 一時ファイルを削除
        if (isTempFile) {
          try {
            await fs.unlink(settlementPath);
          } catch (error) {
            console.error('一時ファイル削除エラー:', error);
          }
        }

        if (propertyData) {
          const settlementResults = job.settlementResults || [];
          settlementResults.push(propertyData);

          await jobManager.updateJob(jobId, {
            settlementProgress: currentSettlementIndex + 1,
            settlementResults
          });
        }

        const totalProgress = pdfFiles.length + currentSettlementIndex + 1;
        const totalItems = pdfFiles.length + settlementFiles.length + transferFiles.length;
        const progressPercent = 30 + Math.round(((currentSettlementIndex + 1) / settlementFiles.length) * 20);

        return res.json({
          success: true,
          completed: false,
          nextStep: 'settlement',
          progress: {
            current: totalProgress,
            total: totalItems,
            percentage: progressPercent,
            message: `決済明細書を解析中 (${currentSettlementIndex + 1}/${settlementFiles.length})...`
          }
        });

      } catch (error) {
        const filename = Buffer.from(settlementFile.originalname, 'latin1').toString('utf8');

        await jobManager.updateJob(jobId, {
          settlementProgress: currentSettlementIndex + 1
        });

        return res.json({
          success: true,
          completed: false,
          nextStep: 'settlement',
          error: `${filename}: ${error.message}`,
          progress: {
            current: pdfFiles.length + currentSettlementIndex + 1,
            total: pdfFiles.length + settlementFiles.length + transferFiles.length,
            percentage: 30 + Math.round(((currentSettlementIndex + 1) / settlementFiles.length) * 20),
            message: `決済明細書を解析中 (${currentSettlementIndex + 1}/${settlementFiles.length})...`
          }
        });
      }
    }

    // Step 3: 譲渡対価証明書PDFを1件処理
    if (currentTransferIndex < transferFiles.length) {
      const transferFile = transferFiles[currentTransferIndex];

      try {
        const originalFilename = Buffer.from(transferFile.originalname, 'latin1').toString('utf8');
        const decodedOriginalName = Buffer.from(transferFile.originalname, 'latin1').toString('utf8');
        const folderPath = transferPathsMap[decodedOriginalName] || '';

        // ファイルパスを取得（pathがない場合は一時ファイルを作成）
        let transferPath;
        let isTempFile = false;

        if (typeof transferFile.path === 'string') {
          transferPath = transferFile.path;
        } else if (Buffer.isBuffer(transferFile.buffer)) {
          // Bufferの場合は一時ファイルを作成
          const tempPath = path.join(uploadDir, `temp_transfer_${Date.now()}.pdf`);
          await fs.writeFile(tempPath, transferFile.buffer);
          transferPath = tempPath;
          isTempFile = true;
        } else {
          throw new Error('Invalid transfer PDF file format in job data');
        }

        const transferData = await parseTransferPDF(
          transferPath,
          originalFilename,
          folderPath
        );

        // 一時ファイルを削除
        if (isTempFile) {
          try {
            await fs.unlink(transferPath);
          } catch (error) {
            console.error('一時ファイル削除エラー:', error);
          }
        }

        if (transferData) {
          const transferResults = job.transferResults || [];
          transferResults.push(transferData);

          await jobManager.updateJob(jobId, {
            transferProgress: currentTransferIndex + 1,
            transferResults
          });
        }

        const totalProgress = pdfFiles.length + settlementFiles.length + currentTransferIndex + 1;
        const totalItems = pdfFiles.length + settlementFiles.length + transferFiles.length;
        const progressPercent = 50 + Math.round(((currentTransferIndex + 1) / transferFiles.length) * 20);

        return res.json({
          success: true,
          completed: false,
          nextStep: 'transfer',
          progress: {
            current: totalProgress,
            total: totalItems,
            percentage: progressPercent,
            message: `譲渡対価証明書を解析中 (${currentTransferIndex + 1}/${transferFiles.length})...`
          }
        });

      } catch (error) {
        const filename = Buffer.from(transferFile.originalname, 'latin1').toString('utf8');

        await jobManager.updateJob(jobId, {
          transferProgress: currentTransferIndex + 1
        });

        return res.json({
          success: true,
          completed: false,
          nextStep: 'transfer',
          error: `${filename}: ${error.message}`,
          progress: {
            current: pdfFiles.length + settlementFiles.length + currentTransferIndex + 1,
            total: pdfFiles.length + settlementFiles.length + transferFiles.length,
            percentage: 50 + Math.round(((currentTransferIndex + 1) / transferFiles.length) * 20),
            message: `譲渡対価証明書を解析中 (${currentTransferIndex + 1}/${transferFiles.length})...`
          }
        });
      }
    }

    // Step 4: 全件処理完了 → Excelファイル生成
    console.log(`[Job ${jobId}] すべてのPDF処理完了。Excelファイルを生成中...`);

    const pdfResults = job.pdfResults || [];
    const settlementResults = job.settlementResults || [];
    const transferResults = job.transferResults || [];

    if (pdfResults.length === 0) {
      await jobManager.failJob(jobId, new Error('すべてのPDFファイルの解析に失敗しました'));
      return res.status(400).json({
        error: 'すべてのPDFファイルの解析に失敗しました'
      });
    }

    // Excel更新
    console.log(`[Job ${jobId}] Excelパス確認:`, job.data.excelPath);
    console.log(`[Job ${jobId}] Excelパスの型:`, typeof job.data.excelPath);

    if (!job.data.excelPath || typeof job.data.excelPath !== 'string') {
      throw new Error('Excelファイルのパスが不正です');
    }

    // Excelファイルの存在確認
    try {
      await fs.access(job.data.excelPath);
      console.log(`[Job ${jobId}] Excelファイル存在確認OK`);
    } catch (error) {
      console.error(`[Job ${jobId}] Excelファイルが見つかりません:`, job.data.excelPath);
      throw new Error(`Excelファイルが見つかりません: ${job.data.excelPath}`);
    }

    console.log(`[Job ${jobId}] Excel更新を開始します`);

    // updateExcelの正しい引数順序:
    // (excelPath, pdfDataArray, itemMapping, settlementFileNames, propertiesData, folderName)
    const result = await updateExcel(
      job.data.excelPath,
      pdfResults,
      itemMapping,
      [],  // settlementFileNames (不要、空配列)
      settlementResults,
      ''   // folderName (空文字列)
    );

    const outputPath = result.outputPath;
    console.log(`[Job ${jobId}] Excel更新完了:`, outputPath);

    // 一時ファイルのクリーンアップ（pathが文字列の場合のみ）
    for (const pdfFile of pdfFiles) {
      try {
        if (typeof pdfFile.path === 'string') {
          await fs.unlink(pdfFile.path);
        }
      } catch (error) {
        console.error(`ファイル削除エラー:`, error);
      }
    }

    for (const settlementFile of settlementFiles) {
      try {
        if (typeof settlementFile.path === 'string') {
          await fs.unlink(settlementFile.path);
        }
      } catch (error) {
        console.error(`ファイル削除エラー:`, error);
      }
    }

    for (const transferFile of transferFiles) {
      try {
        if (typeof transferFile.path === 'string') {
          await fs.unlink(transferFile.path);
        }
      } catch (error) {
        console.error(`ファイル削除エラー:`, error);
      }
    }

    try {
      if (typeof job.data.excelPath === 'string') {
        await fs.unlink(job.data.excelPath);
      }
    } catch (error) {
      console.error(`Excelファイル削除エラー:`, error);
    }

    // ジョブ完了
    await jobManager.completeJob(jobId, outputPath);

    return res.json({
      success: true,
      completed: true,
      outputPath,
      outputFilename,
      progress: {
        current: pdfFiles.length + settlementFiles.length + transferFiles.length,
        total: pdfFiles.length + settlementFiles.length + transferFiles.length,
        percentage: 100,
        message: '処理が完了しました'
      },
      stats: {
        pdfCount: pdfResults.length,
        settlementCount: settlementResults.length,
        transferCount: transferResults.length
      }
    });

  } catch (error) {
    console.error('Process next error:', error);

    if (req.body.jobId) {
      await jobManager.failJob(req.body.jobId, error);
    }

    res.status(500).json({
      error: 'Processing failed',
      details: error.message
    });
  }
});

// ========================================
// 同期処理エンドポイント（既存のまま維持）
// ========================================

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
