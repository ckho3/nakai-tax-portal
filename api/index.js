// Vercel Serverless Function wrapper for Express app
const app = require('../server');

// Vercel Serverless Functionとしてエクスポート
module.exports = (req, res) => {
  // Expressアプリにリクエストを渡す
  return app(req, res);
};
