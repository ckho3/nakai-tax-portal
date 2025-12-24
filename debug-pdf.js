const fs = require('fs');
const pdfParse = require('pdf-parse');

async function debugPDF() {
  // outputディレクトリからPDFファイルを探す
  const uploadsDir = './uploads';
  const files = fs.readdirSync(uploadsDir);

  // ハーモニーレジデンスを含むファイルを検索
  const targetFile = files.find(f => f.includes('ハーモニー') || f.includes('ファステート') || f.includes('プランドール'));

  if (!targetFile) {
    console.log('対象PDFファイルが見つかりません');
    console.log('利用可能なファイル:', files.filter(f => f.endsWith('.pdf')));
    return;
  }

  console.log(`デバッグ対象: ${targetFile}\n`);

  const pdfPath = `${uploadsDir}/${targetFile}`;
  const dataBuffer = fs.readFileSync(pdfPath);
  const data = await pdfParse(dataBuffer);
  const text = data.text;

  console.log('===== PDF全文 =====');
  console.log(text);
  console.log('\n===== 行ごとの解析 =====\n');

  const lines = text.split('\n');

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // 収入合計または賃料を含む行
    if (line.includes('収入合計') || line.includes('収⼊合計') || line.includes('賃料')) {
      console.log(`Line ${i}: [収入関連]`);
      console.log(`  内容: ${line}`);

      // 数値パターンをチェック
      const amounts = line.match(/(\d{1,3}(?:,\d{3})*円|--)/g);
      console.log(`  抽出された金額: ${amounts ? amounts.length : 0}個`);
      if (amounts) {
        console.log(`  金額リスト: ${amounts.join(', ')}`);
      }

      // 既存のパターンマッチ
      const oldPattern = line.match(/(\d{1,3},?\d{3}円.*){10,}/);
      console.log(`  既存パターンマッチ: ${oldPattern ? 'OK' : 'NG'}`);
      console.log('');
    }
  }
}

debugPDF().catch(err => console.error('エラー:', err));
