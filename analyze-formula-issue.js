/**
 * 行追加による数式のずれ問題を分析
 */

console.log('========================================');
console.log('行追加による数式のずれ問題');
console.log('========================================\n');

console.log('【E】サブリースセクション（Row 147）の元の数式:');
console.log('  H147: =H124  （【D】セクションの1個目を参照）');
console.log('  I147: =IF(I$146>=$U147,IF(I$53="サブリース",J147,0),0)');
console.log('  T147: =IF(T$146>=$U147,IF(T$53="サブリース",ROUNDDOWN((T55-T78)*$B$146,-2),0),0)');
console.log('');

console.log('========================================');
console.log('問題: 3行追加した場合');
console.log('========================================\n');

console.log('【A】セクションに3行追加:');
console.log('  Row 55-77 (23行) ← 元々Row 55-74だった');
console.log('  → Row 78以降が全て+3行ずれる');
console.log('');

console.log('【B】セクションに3行追加:');
console.log('  Row 81-103 (23行) ← 元々Row 78-99だった');
console.log('  → Row 101以降が全て+3行ずれる（累計+6行）');
console.log('');

console.log('【C】セクションに3行追加:');
console.log('  Row 107-129 (23行) ← 元々Row 101-122だった');
console.log('  → Row 124以降が全て+3行ずれる（累計+9行）');
console.log('');

console.log('【D】セクションに3行追加:');
console.log('  Row 133-155 (23行) ← 元々Row 124-145だった');
console.log('  → Row 147以降が全て+3行ずれる（累計+12行）');
console.log('');

console.log('【E】サブリースセクション:');
console.log('  元々Row 147 → Row 159に移動（+12行）');
console.log('');

console.log('========================================');
console.log('Row 159の数式がどうなるか？');
console.log('========================================\n');

console.log('相対参照の数式は自動調整される:');
console.log('  元: =IF(I$146>=$U147,...)');
console.log('  → =IF(I$158>=$U159,...)  ← 行番号が+12される ✅');
console.log('');

console.log('しかし、T147の数式の問題:');
console.log('  元: =IF(T$146>=$U147,IF(T$53="サブリース",ROUNDDOWN((T55-T78)*$B$146,-2),0),0)');
console.log('  → =IF(T$158>=$U159,IF(T$53="サブリース",ROUNDDOWN((T67-T90)*$B$158,-2),0),0)');
console.log('');

console.log('問題点:');
console.log('  ❌ T55 → T67に変更される');
console.log('     本来はT55（【A】の1行目）を参照したいが、ずれてしまう');
console.log('');
console.log('  ❌ T78 → T90に変更される');
console.log('     本来はT81（【B】の1行目、行追加後）を参照したいが、間違っている');
console.log('');

console.log('========================================');
console.log('H147の数式の問題');
console.log('========================================\n');

console.log('元の数式:');
console.log('  H147: =H124  （【D】の1行目を参照）');
console.log('');

console.log('3行追加後、Row 159に移動すると:');
console.log('  H159: =H136  ← 自動的に+12される');
console.log('');

console.log('しかし、実際の【D】の1行目は:');
console.log('  Row 133（元々Row 124）');
console.log('');

console.log('問題:');
console.log('  ❌ H159は H136を参照するが、');
console.log('     H136は【D】セクションの4行目（Row 133の+3行目）');
console.log('  ✅ 本来は H133（【D】の1行目）を参照すべき');
console.log('');

console.log('========================================');
console.log('解決策');
console.log('========================================\n');

console.log('Option 1: 絶対参照を使う（$を付ける）');
console.log('  H147: =$H$124  ← 行追加してもH124を参照し続ける');
console.log('  問題: これでも元のテンプレートがH124固定なので不完全');
console.log('');

console.log('Option 2: INDIRECT関数を使う');
console.log('  H147: =INDIRECT("H124")');
console.log('  問題: 行追加後もH124を参照してしまう（133を参照すべき）');
console.log('');

console.log('Option 3: 行追加後に数式を手動で修正する ✅');
console.log('  1. 行追加を実行');
console.log('  2. 【E】セクションの数式を正しい位置に修正');
console.log('  3. 例: H159 = H133（【D】の1行目）に修正');
console.log('');

console.log('Option 4: 名前付き範囲を使う');
console.log('  1. 【D】の1行目に名前を付ける（例: D_First_Row）');
console.log('  2. H147: =D_First_Row');
console.log('  問題: 動的な範囲の管理が複雑');
console.log('');

console.log('========================================');
console.log('推奨される対応');
console.log('========================================\n');

console.log('行追加後、【E】サブリースセクションの数式を修正:');
console.log('');
console.log('修正が必要な箇所:');
console.log('  1. H列: =H124 → =H133（【D】の1行目に修正）');
console.log('  2. T列の数式内: (T55-T78) → (T55-T81) に修正');
console.log('     ※T55は【A】の1行目（変わらない）');
console.log('     ※T78は【B】の元の位置、T81は行追加後の【B】の1行目');
console.log('');

console.log('これをコードで実装する必要があります。');
