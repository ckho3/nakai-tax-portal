const { copyDepreciationData } = require('./depreciationCopier');
const path = require('path');
const fs = require('fs');

async function test() {
  // 最新のExcelファイルを使用
  const outputDir = path.join(__dirname, 'output');
  const files = fs.readdirSync(outputDir)
    .filter(f => f.endsWith('.xlsx'))
    .map(f => ({
      name: f,
      path: path.join(outputDir, f),
      time: fs.statSync(path.join(outputDir, f)).mtime.getTime()
    }))
    .sort((a, b) => b.time - a.time);

  if (files.length === 0) {
    console.log('出力ファイルが見つかりません');
    return;
  }

  const excelPath = files[0].path;
  console.log(`使用ファイル: ${files[0].name}\n`);

  try {
    const result = await copyDepreciationData(excelPath);
    console.log('\n処理結果:');
    console.log(`  成功: ${result.success}`);
    console.log(`  転記件数: ${result.copyCount}件`);
    console.log(`  出力ファイル: ${path.basename(result.outputPath)}`);
  } catch (error) {
    console.error('エラー:', error.message);
  }
}

test();
