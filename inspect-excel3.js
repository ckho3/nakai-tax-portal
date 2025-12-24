const ExcelJS = require('exceljs');
const path = require('path');

async function inspectExcel() {
  try {
    const workbook = new ExcelJS.Workbook();
    const excelPath = path.join(__dirname, '【原本】R7確定申告フォーマット.xlsx');
    await workbook.xlsx.readFile(excelPath);

    const sheet = workbook.getWorksheet('【不】①不動産収入');

    if (!sheet) {
      console.log('シート【不】①不動産収入が見つかりません');
      return;
    }

    console.log('========================================');
    console.log('C列とG-H列の確認 (Row 50-150)');
    console.log('========================================\n');

    const sections = [];

    for (let rowNum = 50; rowNum <= 150; rowNum++) {
      const row = sheet.getRow(rowNum);
      const cellC = row.getCell(3).value; // C列
      const cellG = row.getCell(7).value; // G列
      const cellH = row.getCell(8).value; // H列

      const formatCell = (val) => {
        if (!val) return '';
        if (typeof val === 'object' && val.formula) {
          return `[Formula]`;
        }
        return val.toString().substring(0, 60);
      };

      // C列に【】が含まれる行、またはG列/H列に文字列がある行を表示
      const cellCText = formatCell(cellC);
      const cellGText = formatCell(cellG);
      const cellHText = formatCell(cellH);

      if (cellCText.includes('【') || cellCText.includes('収入') ||
          cellCText.includes('管理') || cellCText.includes('広告') ||
          cellCText.includes('修繕') || cellCText.includes('支払')) {
        console.log(`Row ${rowNum}:`);
        console.log(`  C="${cellCText}"`);
        console.log(`  G="${cellGText}"`);
        console.log(`  H="${cellHText}"`);
        console.log('');

        // セクション情報を保存
        if (cellCText.includes('【A】')) sections.push({ section: 'A', row: rowNum, name: cellCText });
        if (cellCText.includes('【B】')) sections.push({ section: 'B', row: rowNum, name: cellCText });
        if (cellCText.includes('【C】')) sections.push({ section: 'C', row: rowNum, name: cellCText });
        if (cellCText.includes('【D】')) sections.push({ section: 'D', row: rowNum, name: cellCText });
      }
    }

    console.log('\n========================================');
    console.log('検出されたセクション一覧');
    console.log('========================================\n');

    for (const sec of sections) {
      console.log(`【${sec.section}】 Row ${sec.row}: ${sec.name}`);
    }

    if (sections.length >= 2) {
      console.log('\n========================================');
      console.log('セクション間の行数');
      console.log('========================================\n');

      for (let i = 0; i < sections.length - 1; i++) {
        const current = sections[i];
        const next = sections[i + 1];
        const gap = next.row - current.row;
        console.log(`【${current.section}】(Row ${current.row}) → 【${next.section}】(Row ${next.row}): ${gap}行の差`);
      }
    }

    console.log('\n========================================');
    console.log('各セクションのデータ行（最初の3行のG-H列）');
    console.log('========================================\n');

    for (const sec of sections) {
      console.log(`\n【${sec.section}】セクション (Row ${sec.row}から) のデータ行:`);
      for (let i = 1; i <= 5; i++) {
        const dataRow = sheet.getRow(sec.row + i);
        const cellG = dataRow.getCell(7).value;
        const cellH = dataRow.getCell(8).value;
        const cellGText = formatCell(cellG);
        const cellHText = formatCell(cellH);

        if (cellGText || cellHText) {
          console.log(`  Row ${sec.row + i}: G="${cellGText}" H="${cellHText}"`);
        }
      }
    }

  } catch (error) {
    console.error('エラー:', error);
  }
}

inspectExcel();
