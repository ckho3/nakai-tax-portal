// Vercel Serverless Function wrapper for Express app
let app;

// 遅延ロード: 初回リクエスト時にExpressアプリを読み込む
const getApp = () => {
  if (!app) {
    console.log('Loading Express app...');
    app = require('../server');
    console.log('Express app loaded successfully');
  }
  return app;
};

// Vercel Serverless Functionとしてエクスポート
module.exports = async (req, res) => {
  try {
    console.log(`Incoming request: ${req.method} ${req.url}`);
    const expressApp = getApp();
    return expressApp(req, res);
  } catch (error) {
    console.error('Error in serverless function:', error);
    res.status(500).json({
      error: 'サーバーエラーが発生しました',
      message: error.message,
      stack: process.env.NODE_ENV !== 'production' ? error.stack : undefined
    });
  }
};
