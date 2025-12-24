const fs = require('fs').promises;
const path = require('path');
const { parsePDF } = require('./pdfParser');
const { parseSettlementPDF } = require('./settlementParser');
const { parseTransferPDF } = require('./transferParser');
const { updateExcel } = require('./excelWriter');
const { JobManager } = require('./jobManager');

const isVercel = process.env.VERCEL === '1';
const outputDir = isVercel ? '/tmp/output' : path.join(__dirname, 'output');

class AsyncProcessor {
  constructor() {
    this.jobManager = new JobManager();
  }

  // 非同期でPDF処理を実行
  async processJob(jobId) {
    try {
      const job = await this.jobManager.getJob(jobId);
      if (!job) {
        throw new Error('ジョブが見つかりません');
      }

      await this.jobManager.updateJob(jobId, { status: 'processing' });

      const {
        pdfFiles,
        excelPath,
        settlementFiles = [],
        transferFiles = [],
        pdfPathsMap = {},
        settlementPathsMap = {},
        transferPathsMap = {},
        itemMapping = {}
      } = job.data;

      // ステップ1: PDFファイルの解析
      await this.jobManager.updateProgress(jobId, 10, '年間収支一覧表PDFを解析中...');
      console.log(`[Job ${jobId}] PDFファイル解析開始: ${pdfFiles.length}件`);

      const pdfDataArray = [];
      const parseErrors = [];
      const unknownItems = new Set();

      for (let i = 0; i < pdfFiles.length; i++) {
        const pdfFile = pdfFiles[i];
        const progress = 10 + Math.floor((i / pdfFiles.length) * 20);
        await this.jobManager.updateProgress(
          jobId,
          progress,
          `年間収支一覧表を解析中 (${i + 1}/${pdfFiles.length})...`
        );

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

          const decodedOriginalName = Buffer.from(pdfFile.originalname, 'latin1').toString('utf8');
          const folderPath = pdfPathsMap[decodedOriginalName] || '';
          pdfDataArray.push({
            ...pdfData,
            filename: filename,
            folderPath: folderPath
          });

          console.log(`[Job ${jobId}] ✓ ${filename} - ${pdfData.propertyName}`);
        } catch (error) {
          const filename = Buffer.from(pdfFile.originalname, 'latin1').toString('utf8');
          parseErrors.push({
            filename,
            error: error.message
          });
          console.error(`[Job ${jobId}] ✗ ${filename}:`, error.message);
        }
      }

      if (pdfDataArray.length === 0) {
        throw new Error('すべてのPDFファイルの解析に失敗しました');
      }

      // ステップ2: 決済明細書PDFの解析
      let propertiesData = [];
      const settlementParseErrors = [];

      if (settlementFiles.length > 0) {
        await this.jobManager.updateProgress(jobId, 30, '決済明細書PDFを解析中...');
        console.log(`[Job ${jobId}] 決済明細書PDFの解析: ${settlementFiles.length}件`);

        for (let i = 0; i < settlementFiles.length; i++) {
          const settlementFile = settlementFiles[i];
          const progress = 30 + Math.floor((i / settlementFiles.length) * 20);
          await this.jobManager.updateProgress(
            jobId,
            progress,
            `決済明細書を解析中 (${i + 1}/${settlementFiles.length})...`
          );

          try {
            const originalFilename = Buffer.from(settlementFile.originalname, 'latin1').toString('utf8');
            const decodedOriginalName = Buffer.from(settlementFile.originalname, 'latin1').toString('utf8');
            const folderPath = settlementPathsMap[decodedOriginalName] || '';

            const settlementBuffer = await fs.readFile(settlementFile.path);
            const propertyData = await parseSettlementPDF(
              settlementBuffer,
              originalFilename,
              folderPath
            );

            if (propertyData) {
              propertiesData.push(propertyData);
              console.log(`[Job ${jobId}] ✓ 決済明細書解析: ${propertyData.propertyName}`);
            }
          } catch (error) {
            const filename = Buffer.from(settlementFile.originalname, 'latin1').toString('utf8');
            settlementParseErrors.push({
              filename,
              error: error.message
            });
            console.error(`[Job ${jobId}] ✗ 決済明細書解析失敗 (${filename}):`, error.message);
          }
        }
      }

      // ステップ3: 譲渡対価証明書PDFの解析
      let transfersData = [];
      const transferParseErrors = [];

      if (transferFiles.length > 0) {
        await this.jobManager.updateProgress(jobId, 50, '譲渡対価証明書PDFを解析中...');
        console.log(`[Job ${jobId}] 譲渡対価証明書PDFの解析: ${transferFiles.length}件`);

        for (let i = 0; i < transferFiles.length; i++) {
          const transferFile = transferFiles[i];
          const progress = 50 + Math.floor((i / transferFiles.length) * 20);
          await this.jobManager.updateProgress(
            jobId,
            progress,
            `譲渡対価証明書を解析中 (${i + 1}/${transferFiles.length})...`
          );

          try {
            const originalFilename = Buffer.from(transferFile.originalname, 'latin1').toString('utf8');
            const decodedOriginalName = Buffer.from(transferFile.originalname, 'latin1').toString('utf8');
            const folderPath = transferPathsMap[decodedOriginalName] || '';

            const transferBuffer = await fs.readFile(transferFile.path);
            const transferData = await parseTransferPDF(
              transferBuffer,
              originalFilename,
              folderPath
            );

            if (transferData) {
              transfersData.push(transferData);
              console.log(`[Job ${jobId}] ✓ 譲渡対価証明書解析: ${transferData.propertyName}`);
            }
          } catch (error) {
            const filename = Buffer.from(transferFile.originalname, 'latin1').toString('utf8');
            transferParseErrors.push({
              filename,
              error: error.message
            });
            console.error(`[Job ${jobId}] ✗ 譲渡対価証明書解析失敗 (${filename}):`, error.message);
          }
        }
      }

      // ステップ4: Excelファイルの更新
      await this.jobManager.updateProgress(jobId, 70, 'Excelファイルを更新中...');
      console.log(`[Job ${jobId}] Excelファイル更新開始`);

      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
      const outputFilename = `年間収支一覧表_更新済_${timestamp}.xlsx`;
      const outputPath = path.join(outputDir, outputFilename);

      await updateExcel(
        excelPath,
        pdfDataArray,
        outputPath,
        itemMapping,
        propertiesData,
        transfersData
      );

      await this.jobManager.updateProgress(jobId, 90, 'ファイルを保存中...');

      // ステップ5: 一時ファイルのクリーンアップ
      await this.jobManager.updateProgress(jobId, 95, '一時ファイルをクリーンアップ中...');
      console.log(`[Job ${jobId}] 一時ファイルのクリーンアップ`);

      // PDFファイルを削除
      for (const pdfFile of pdfFiles) {
        try {
          await fs.unlink(pdfFile.path);
        } catch (error) {
          console.error(`[Job ${jobId}] ファイル削除エラー:`, error);
        }
      }

      // 決済明細書PDFファイルを削除
      for (const settlementFile of settlementFiles) {
        try {
          await fs.unlink(settlementFile.path);
        } catch (error) {
          console.error(`[Job ${jobId}] ファイル削除エラー:`, error);
        }
      }

      // 譲渡対価証明書PDFファイルを削除
      for (const transferFile of transferFiles) {
        try {
          await fs.unlink(transferFile.path);
        } catch (error) {
          console.error(`[Job ${jobId}] ファイル削除エラー:`, error);
        }
      }

      // Excelファイルを削除
      try {
        await fs.unlink(excelPath);
      } catch (error) {
        console.error(`[Job ${jobId}] Excelファイル削除エラー:`, error);
      }

      // ジョブ完了
      await this.jobManager.completeJob(jobId, outputPath);
      console.log(`[Job ${jobId}] 処理完了: ${outputPath}`);

      return {
        success: true,
        outputPath,
        outputFilename,
        stats: {
          pdfCount: pdfDataArray.length,
          settlementCount: propertiesData.length,
          transferCount: transfersData.length,
          parseErrors: parseErrors.length,
          settlementParseErrors: settlementParseErrors.length,
          transferParseErrors: transferParseErrors.length
        }
      };
    } catch (error) {
      console.error(`[Job ${jobId}] エラー:`, error);
      await this.jobManager.failJob(jobId, error);
      throw error;
    }
  }

  // ジョブを非同期で開始（ノンブロッキング）
  startJob(jobId) {
    // Promise を返さずに、バックグラウンドで実行
    this.processJob(jobId).catch(error => {
      console.error(`[Job ${jobId}] バックグラウンド処理エラー:`, error);
    });
  }
}

module.exports = { AsyncProcessor };
