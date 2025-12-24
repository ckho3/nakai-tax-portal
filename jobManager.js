const fs = require('fs').promises;
const path = require('path');

// ジョブストレージディレクトリ
const isVercel = process.env.VERCEL === '1';
const jobsDir = isVercel ? '/tmp/jobs' : path.join(__dirname, 'jobs');

// ジョブのステータス
const JobStatus = {
  QUEUED: 'queued',
  PROCESSING: 'processing',
  COMPLETED: 'completed',
  FAILED: 'failed'
};

class JobManager {
  constructor() {
    this.initializeJobsDir();
  }

  async initializeJobsDir() {
    try {
      await fs.mkdir(jobsDir, { recursive: true });
      console.log(`ジョブディレクトリ: ${jobsDir}`);
    } catch (err) {
      console.error('ジョブディレクトリ作成エラー:', err);
    }
  }

  // 新しいジョブを作成
  async createJob(jobData) {
    const jobId = `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    const job = {
      id: jobId,
      status: JobStatus.QUEUED,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      progress: 0,
      message: 'ジョブを初期化中...',
      data: jobData,
      error: null,
      resultPath: null
    };

    await this.saveJob(job);
    console.log(`ジョブ作成: ${jobId}`);
    return job;
  }

  // ジョブを保存
  async saveJob(job) {
    const jobPath = path.join(jobsDir, `${job.id}.json`);
    await fs.writeFile(jobPath, JSON.stringify(job, null, 2));
  }

  // ジョブを取得
  async getJob(jobId) {
    try {
      const jobPath = path.join(jobsDir, `${jobId}.json`);
      const jobData = await fs.readFile(jobPath, 'utf8');
      return JSON.parse(jobData);
    } catch (error) {
      console.error(`ジョブ取得エラー (${jobId}):`, error.message);
      return null;
    }
  }

  // ジョブのステータスを更新
  async updateJob(jobId, updates) {
    const job = await this.getJob(jobId);
    if (!job) {
      throw new Error(`ジョブが見つかりません: ${jobId}`);
    }

    const updatedJob = {
      ...job,
      ...updates,
      updatedAt: new Date().toISOString()
    };

    await this.saveJob(updatedJob);
    return updatedJob;
  }

  // ジョブの進捗を更新
  async updateProgress(jobId, progress, message) {
    return await this.updateJob(jobId, {
      progress,
      message,
      status: progress === 100 ? JobStatus.COMPLETED : JobStatus.PROCESSING
    });
  }

  // ジョブを完了としてマーク
  async completeJob(jobId, resultPath) {
    return await this.updateJob(jobId, {
      status: JobStatus.COMPLETED,
      progress: 100,
      message: '処理が完了しました',
      resultPath
    });
  }

  // ジョブを失敗としてマーク
  async failJob(jobId, error) {
    return await this.updateJob(jobId, {
      status: JobStatus.FAILED,
      progress: 0,
      message: 'エラーが発生しました',
      error: error.message || error.toString()
    });
  }

  // 古いジョブを削除（1時間以上前のジョブ）
  async cleanupOldJobs() {
    try {
      const files = await fs.readdir(jobsDir);
      const now = Date.now();
      const oneHour = 60 * 60 * 1000;

      for (const file of files) {
        if (!file.endsWith('.json')) continue;

        const jobPath = path.join(jobsDir, file);
        const jobData = await fs.readFile(jobPath, 'utf8');
        const job = JSON.parse(jobData);

        const jobAge = now - new Date(job.createdAt).getTime();
        if (jobAge > oneHour) {
          await fs.unlink(jobPath);
          console.log(`古いジョブを削除: ${job.id}`);
        }
      }
    } catch (error) {
      console.error('ジョブクリーンアップエラー:', error);
    }
  }
}

module.exports = { JobManager, JobStatus };
