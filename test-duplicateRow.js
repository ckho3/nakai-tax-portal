/**
 * ExcelJS の duplicateRow() メソッドの動作を説明
 */

console.log('========================================');
console.log('duplicateRow()の動作説明');
console.log('========================================\n');

console.log('構文:');
console.log('  sheet.duplicateRow(rowNumber, amount, insert)');
console.log('');

console.log('パラメータ:');
console.log('  rowNumber: 複製元の行番号');
console.log('  amount: 複製する行数（通常は1）');
console.log('  insert: true = 元の行の後に挿入、false = 元の行を上書き');
console.log('');

console.log('========================================');
console.log('例: sheet.duplicateRow(75, 1, true)');
console.log('========================================\n');

console.log('実行前の状態:');
console.log('  Row 55: ヘッダー');
console.log('  Row 56-74: データ行1-19');
console.log('  Row 75: データ行20（最終行）← これを複製');
console.log('  Row 76: （次のセクション）');
console.log('');

console.log('実行内容:');
console.log('  1. Row 75の内容（セルの値、書式、数式など）を全てコピー');
console.log('  2. Row 75の後に新しい行を挿入（Row 76に挿入）');
console.log('  3. Row 76以降の既存の行を全て1行下にずらす');
console.log('  4. 新しく挿入されたRow 76にコピーした内容を貼り付け');
console.log('');

console.log('実行後の状態:');
console.log('  Row 55: ヘッダー');
console.log('  Row 56-74: データ行1-19');
console.log('  Row 75: データ行20（元の行）');
console.log('  Row 76: データ行20のコピー（新しく挿入された行）✅');
console.log('  Row 77: （次のセクション）← 元々Row 76だった行が1行下にずれた');
console.log('');

console.log('========================================');
console.log('コピーされる内容');
console.log('========================================\n');

console.log('duplicateRow()は以下を全てコピーします:');
console.log('  ✅ セルの値（数値、文字列など）');
console.log('  ✅ セルの書式（フォント、色、罫線など）');
console.log('  ✅ 数式（相対参照は自動調整される）');
console.log('  ✅ セルの結合情報');
console.log('  ✅ 行の高さ');
console.log('  ✅ 条件付き書式');
console.log('');

console.log('例えば、Row 75に以下の数式があった場合:');
console.log('  V75: =SUM(I75:T75)');
console.log('');
console.log('Row 76に複製されると:');
console.log('  V76: =SUM(I76:T76)  ← 自動的に行番号が調整される');
console.log('');

console.log('========================================');
console.log('複数行追加する場合のループ');
console.log('========================================\n');

console.log('extraRows = 3の場合:');
console.log('');

console.log('ループ1回目 (i=0):');
console.log('  duplicateRow(75, 1, true)');
console.log('  → Row 75を複製してRow 76に挿入');
console.log('  → 結果: Row 55-76 (22行)');
console.log('');

console.log('ループ2回目 (i=1):');
console.log('  duplicateRow(76, 1, true)  ← 75 + 1');
console.log('  → Row 76を複製してRow 77に挿入');
console.log('  → 結果: Row 55-77 (23行)');
console.log('');

console.log('ループ3回目 (i=2):');
console.log('  duplicateRow(77, 1, true)  ← 75 + 2');
console.log('  → Row 77を複製してRow 78に挿入');
console.log('  → 結果: Row 55-78 (24行)');
console.log('');

console.log('なぜ「75 + i」とするのか:');
console.log('  - 1回目でRow 75を複製すると、新しいRow 76が追加される');
console.log('  - 2回目は最新の最終行（Row 76）を複製する必要がある');
console.log('  - 3回目は最新の最終行（Row 77）を複製する必要がある');
console.log('  - よって、「75 + i」で常に最新の最終行を指定する');
console.log('');

console.log('========================================');
console.log('後ろから前に向かって追加する理由');
console.log('========================================\n');

console.log('コード例:');
console.log('  for (let i = 0; i < extraRows; i++) {');
console.log('    sheet.duplicateRow(145 + i, 1, true);  // 【D】');
console.log('    sheet.duplicateRow(122 + i, 1, true);  // 【C】');
console.log('    sheet.duplicateRow(99 + i, 1, true);   // 【B】');
console.log('    sheet.duplicateRow(75 + i, 1, true);   // 【A】');
console.log('  }');
console.log('');

console.log('理由:');
console.log('  1. 【D】に行を追加 → 【D】のみ影響（後ろにセクションがない）');
console.log('  2. 【C】に行を追加 → 【C】と【D】が下にずれる');
console.log('  3. 【B】に行を追加 → 【B】【C】【D】が下にずれる');
console.log('  4. 【A】に行を追加 → 【A】【B】【C】【D】が下にずれる');
console.log('');
console.log('後ろから追加することで、前のセクションの行番号を');
console.log('気にせず安全に追加できる。');
console.log('');

console.log('========================================');
console.log('まとめ');
console.log('========================================\n');

console.log('duplicateRow(75, 1, true) は:');
console.log('  ❌ 単に空の行を挿入するだけではない');
console.log('  ✅ Row 75の全ての内容（値、書式、数式）をコピーして新しい行として挿入する');
console.log('');
console.log('これにより:');
console.log('  - デザイン（罫線、色、フォント）が保持される');
console.log('  - 数式が自動調整される');
console.log('  - テンプレートの構造が維持される');
