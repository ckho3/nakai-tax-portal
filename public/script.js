// DOM要素の取得
const excelInput = document.getElementById('excelInput');
const folderInput = document.getElementById('folderInput');
const excelFileName = document.getElementById('excelFileName');
const clearExcelBtn = document.getElementById('clearExcelBtn');
const pdfFileList = document.getElementById('pdfFileList');
const settlementFileList = document.getElementById('settlementFileList');
const transferFileList = document.getElementById('transferFileList');
const folderScanResult = document.getElementById('folderScanResult');
const uploadBtn = document.getElementById('uploadBtn');
const progressSection = document.getElementById('progressSection');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const resultSection = document.getElementById('resultSection');
const resultContent = document.getElementById('resultContent');

let excelFile = null;
let pdfFiles = [];
let settlementFiles = [];
let transferFiles = [];

// localStorageのキー
const EXCEL_STORAGE_KEY = 'nakai_solutions_excel_file';

// ページ読み込み時に保存されたExcelファイルを復元
window.addEventListener('DOMContentLoaded', async () => {
  const savedExcelData = localStorage.getItem(EXCEL_STORAGE_KEY);

  if (savedExcelData) {
    try {
      const { name, dataUrl } = JSON.parse(savedExcelData);

      // Base64データをBlobに変換
      const response = await fetch(dataUrl);
      const blob = await response.blob();

      // FileオブジェクトとしてexcelFileに設定
      excelFile = new File([blob], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

      // UI更新
      excelFileName.textContent = `✓ ${name} (保存済み)`;
      excelFileName.style.color = '#28a745';
      clearExcelBtn.style.display = 'inline-block';

      console.log(`保存されたExcelファイルを復元しました: ${name}`);
      checkUploadButton();
    } catch (error) {
      console.error('Excelファイルの復元に失敗しました:', error);
      localStorage.removeItem(EXCEL_STORAGE_KEY);
    }
  }
});

// Excelファイル選択
excelInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (file) {
    excelFile = file;
    excelFileName.textContent = `✓ ${file.name}`;
    excelFileName.style.color = '#28a745';
    clearExcelBtn.style.display = 'inline-block';

    // localStorageに保存（Base64エンコード）
    try {
      const reader = new FileReader();
      reader.onload = (event) => {
        const excelData = {
          name: file.name,
          dataUrl: event.target.result
        };
        localStorage.setItem(EXCEL_STORAGE_KEY, JSON.stringify(excelData));
        console.log(`Excelファイルを保存しました: ${file.name}`);
      };
      reader.readAsDataURL(file);
    } catch (error) {
      console.error('Excelファイルの保存に失敗しました:', error);
    }

    checkUploadButton();
  }
});

// Excelファイルクリアボタン
clearExcelBtn.addEventListener('click', () => {
  if (confirm('保存されたExcelファイルをクリアしますか？')) {
    excelFile = null;
    excelInput.value = '';
    excelFileName.textContent = '';
    clearExcelBtn.style.display = 'none';
    localStorage.removeItem(EXCEL_STORAGE_KEY);
    console.log('Excelファイルをクリアしました');
    checkUploadButton();
  }
});

// フォルダ選択（年間収支一覧表PDFと決済明細書PDFを自動検索）
folderInput.addEventListener('change', (e) => {
  const allFiles = Array.from(e.target.files);

  // 「年間収支一覧表」というテキストを含むPDFファイルをフィルタ
  const annualIncomeFiles = allFiles.filter(file =>
    file.name.toLowerCase().endsWith('.pdf') &&
    file.name.includes('年間収支一覧表')
  );

  // 「決済明細書」というテキストを含むPDFファイルをフィルタ
  const settlementMatchingFiles = allFiles.filter(file =>
    file.name.toLowerCase().endsWith('.pdf') &&
    file.name.includes('決済明細書')
  );

  // 「譲渡対価証明書」というテキストを含むPDFファイルをフィルタ
  const transferMatchingFiles = allFiles.filter(file =>
    file.name.toLowerCase().endsWith('.pdf') &&
    file.name.includes('譲渡対価証明書')
  );

  // 年間収支一覧表PDFを設定
  pdfFiles = annualIncomeFiles;

  // 決済明細書PDFを設定
  settlementFiles = settlementMatchingFiles;

  // 譲渡対価証明書PDFを設定
  transferFiles = transferMatchingFiles;

  // スキャン結果のサマリーを表示
  let resultHTML = '';
  if (pdfFiles.length > 0 || settlementFiles.length > 0 || transferFiles.length > 0) {
    resultHTML = `<div style="color: #28a745; font-weight: bold;">✓ スキャン完了</div>`;
    resultHTML += `<div style="margin-top: 5px;">年間収支一覧表: ${pdfFiles.length}件</div>`;
    resultHTML += `<div>決済明細書: ${settlementFiles.length}件</div>`;
    resultHTML += `<div>譲渡対価証明書: ${transferFiles.length}件</div>`;
  } else {
    resultHTML = '<div style="color: #dc3545;">⚠ PDFファイルが見つかりませんでした</div>';
  }
  folderScanResult.innerHTML = resultHTML;

  // アップロードボタンの有効化チェック
  checkUploadButton();
});

// アップロードボタンの有効/無効を切り替え
function checkUploadButton() {
  // Excelと年間収支一覧表PDFがあれば有効（決済明細書は任意）
  if (excelFile && pdfFiles.length > 0) {
    uploadBtn.disabled = false;
  } else {
    uploadBtn.disabled = true;
  }
}

// アップロード処理（非同期版）
uploadBtn.addEventListener('click', async () => {
  // UIをリセット
  progressSection.style.display = 'block';
  resultSection.style.display = 'none';
  progressFill.style.width = '0%';
  progressText.textContent = 'ファイルをアップロード中...';
  uploadBtn.disabled = true;

  try {
    // 1. ファイルをアップロードしてジョブを開始
    const formData = new FormData();
    formData.append('excel', excelFile);

    // PDFファイルのフォルダパス情報も送信
    const pdfPathsMap = {};
    const settlementPathsMap = {};
    const transferPathsMap = {};

    pdfFiles.forEach((file, index) => {
      formData.append('pdfs', file);
      // webkitRelativePathからフォルダパスを抽出（親フォルダのパス）
      if (file.webkitRelativePath) {
        const folderPath = file.webkitRelativePath.substring(0, file.webkitRelativePath.lastIndexOf('/'));
        pdfPathsMap[file.name] = folderPath;
      }
    });

    // 決済明細書PDFも一緒に送信（ある場合）
    settlementFiles.forEach((file, index) => {
      formData.append('settlements', file);
      if (file.webkitRelativePath) {
        const folderPath = file.webkitRelativePath.substring(0, file.webkitRelativePath.lastIndexOf('/'));
        settlementPathsMap[file.name] = folderPath;
      }
    });

    // 譲渡対価証明書PDFも一緒に送信（ある場合）
    transferFiles.forEach((file, index) => {
      formData.append('transfers', file);
      if (file.webkitRelativePath) {
        const folderPath = file.webkitRelativePath.substring(0, file.webkitRelativePath.lastIndexOf('/'));
        transferPathsMap[file.name] = folderPath;
      }
    });

    // フォルダパス情報をJSON文字列として送信
    formData.append('pdfPaths', JSON.stringify(pdfPathsMap));
    formData.append('settlementPaths', JSON.stringify(settlementPathsMap));
    formData.append('transferPaths', JSON.stringify(transferPathsMap));

    progressFill.style.width = '5%';
    progressText.textContent = 'サーバーにアップロード中...';

    // 非同期エンドポイントを使用
    const uploadResponse = await fetch('/upload-async', {
      method: 'POST',
      body: formData
    });

    if (!uploadResponse.ok) {
      const errorData = await uploadResponse.json();
      throw new Error(errorData.error || 'アップロードに失敗しました');
    }

    const uploadResult = await uploadResponse.json();
    const jobId = uploadResult.jobId;

    console.log(`ジョブID: ${jobId}`);
    progressFill.style.width = '10%';
    progressText.textContent = '処理を開始しました...';

    // 2. ジョブのステータスをポーリング
    await pollJobStatus(jobId);

  } catch (error) {
    console.error('エラー:', error);
    progressSection.style.display = 'none';
    displayError('処理中にエラーが発生しました: ' + error.message);
    uploadBtn.disabled = false;
  }
});

// ジョブステータスをポーリング
async function pollJobStatus(jobId) {
  const pollInterval = 1000; // 1秒ごとにチェック
  const maxAttempts = 600; // 最大10分
  let attempts = 0;

  const poll = async () => {
    attempts++;

    if (attempts > maxAttempts) {
      throw new Error('処理がタイムアウトしました');
    }

    try {
      const statusResponse = await fetch(`/job-status/${jobId}`);

      if (!statusResponse.ok) {
        throw new Error('ステータス取得に失敗しました');
      }

      const status = await statusResponse.json();
      console.log(`ジョブステータス: ${status.status} (${status.progress}%) - ${status.message}`);

      // 進捗を更新
      progressFill.style.width = `${status.progress}%`;
      progressText.textContent = status.message;

      if (status.status === 'completed') {
        // 完了
        progressFill.style.width = '100%';
        progressText.textContent = '処理完了！ファイルをダウンロード中...';

        // ダウンロード
        await downloadResult(jobId);

        // 成功メッセージを表示
        setTimeout(() => {
          progressSection.style.display = 'none';
          displaySuccess('Excelファイルの更新が完了しました！');
          uploadBtn.disabled = false;
        }, 1000);

      } else if (status.status === 'failed') {
        // 失敗
        throw new Error(status.error || '処理に失敗しました');

      } else {
        // まだ処理中 - 次のポーリング
        setTimeout(poll, pollInterval);
      }

    } catch (error) {
      console.error('ポーリングエラー:', error);
      throw error;
    }
  };

  // ポーリング開始
  await poll();
}

// 結果をダウンロード
async function downloadResult(jobId) {
  try {
    const downloadUrl = `/download/${jobId}`;

    // ダウンロードリンクを作成してクリック
    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = ''; // サーバー側のファイル名を使用
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    console.log('ダウンロード開始:', downloadUrl);
  } catch (error) {
    console.error('ダウンロードエラー:', error);
    throw error;
  }
}

// 成功メッセージを表示
function displaySuccess(message) {
  resultSection.style.display = 'block';
  resultSection.innerHTML = `
    <h3>✅ 処理完了</h3>
    <div class="success-message">${message}</div>
    <p>ダウンロードが開始されない場合は、ブラウザの設定を確認してください。</p>
  `;
}

// 未知の項目の分類を聞くダイアログを表示
function showMappingDialog(unknownItems, tempId) {
  const mapping = {};

  // ダイアログHTMLを作成
  const dialogHTML = `
    <div class="mapping-dialog-overlay" id="mappingDialogOverlay">
      <div class="mapping-dialog">
        <h2>🔍 未知の支払項目が見つかりました</h2>
        <p>以下の項目をどのセクションに分類するか選択してください:</p>
        <div class="mapping-items" id="mappingItems">
          ${unknownItems.map(item => `
            <div class="mapping-item">
              <label class="mapping-label">${item}:</label>
              <select class="mapping-select" data-item="${item}">
                <option value="">選択してください</option>
                <option value="B">【B】管理手数料セクション</option>
                <option value="C">【C】広告費等セクション</option>
                <option value="D">【D】修繕費・設備費セクション</option>
              </select>
            </div>
          `).join('')}
        </div>
        <div class="mapping-actions">
          <button class="mapping-btn mapping-btn-save" id="saveMappingBtn">保存して続行</button>
          <button class="mapping-btn mapping-btn-cancel" id="cancelMappingBtn">キャンセル</button>
        </div>
      </div>
    </div>
  `;

  // ダイアログを挿入
  document.body.insertAdjacentHTML('beforeend', dialogHTML);

  // 保存ボタンのイベント
  document.getElementById('saveMappingBtn').addEventListener('click', async () => {
    const selects = document.querySelectorAll('.mapping-select');
    let allSelected = true;

    selects.forEach(select => {
      const itemName = select.dataset.item;
      const section = select.value;

      if (!section) {
        allSelected = false;
      } else {
        mapping[itemName] = section;
      }
    });

    if (!allSelected) {
      alert('すべての項目を選択してください');
      return;
    }

    // ダイアログを閉じる
    document.getElementById('mappingDialogOverlay').remove();

    // 保存して処理を続行
    await saveMappingAndContinue(tempId, mapping);
  });

  // キャンセルボタンのイベント
  document.getElementById('cancelMappingBtn').addEventListener('click', () => {
    document.getElementById('mappingDialogOverlay').remove();
    uploadBtn.disabled = false;
  });
}

// マッピングを保存して処理を続行
async function saveMappingAndContinue(tempId, mapping) {
  progressSection.style.display = 'block';
  progressFill.style.width = '50%';
  progressText.textContent = 'マッピングを保存して処理を続行しています...';

  try {
    const response = await fetch('/save-mapping', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ tempId, mapping })
    });

    const result = await response.json();

    progressFill.style.width = '100%';
    progressText.textContent = '処理完了！';

    setTimeout(() => {
      progressSection.style.display = 'none';
      displayResult(result);
      uploadBtn.disabled = false;
    }, 1000);

  } catch (error) {
    console.error('エラー:', error);
    progressSection.style.display = 'none';
    displayError('処理中にエラーが発生しました: ' + error.message);
    uploadBtn.disabled = false;
  }
}

// 処理結果の表示（統合版）
function displayCombinedResult(annualIncomeResult, settlementResult) {
  resultSection.style.display = 'block';

  let html = '';

  // 年間収支一覧表の処理結果
  if (annualIncomeResult && annualIncomeResult.success) {
    html += `
      <div class="success-message">
        <h4>✓ 年間収支一覧表の処理が完了しました！</h4>
      </div>
    `;
  }

  // ダウンロードボタン（settlementResultを優先、なければannualIncomeResult）
  const downloadUrl = (settlementResult && settlementResult.downloadUrl) ||
                     (annualIncomeResult && annualIncomeResult.downloadUrl);

  resultContent.innerHTML = html;

  // エラーがある場合
  if ((annualIncomeResult && !annualIncomeResult.success) ||
      (settlementResult && !settlementResult.success)) {
    const errorMsg = (annualIncomeResult && annualIncomeResult.error) ||
                    (settlementResult && settlementResult.error);
    if (errorMsg) {
      displayError(errorMsg);
      return;
    }
  }

  // 自動ダウンロード
  if (downloadUrl) {
    setTimeout(() => {
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = '';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }, 500);
  }

  // スクロール
  resultSection.scrollIntoView({ behavior: 'smooth' });
}

// 処理結果の表示（単一結果用 - マッピング後に使用）
function displayResult(result) {
  resultSection.style.display = 'block';

  if (result.success) {
    let html = `
      <div class="success-message">
        <h4>✓ 処理が完了しました！</h4>
        <p>${result.message}</p>
      </div>
    `;

    // ダウンロードボタン
    if (result.downloadUrl) {
      html += `
        <a href="${result.downloadUrl}" download class="download-btn">
          📥 更新されたExcelファイルをダウンロード
        </a>
      `;
    }

    // 処理結果の詳細テーブル
    if (result.results && result.results.length > 0) {
      html += `
        <table class="result-table">
          <thead>
            <tr>
              <th>物件名</th>
              <th>ステータス</th>
              <th>メッセージ</th>
            </tr>
          </thead>
          <tbody>
      `;

      result.results.forEach(item => {
        const statusClass = item.status === 'success' ? 'status-success' : 'status-error';
        const statusIcon = item.status === 'success' ? '✓' : '✗';
        html += `
          <tr>
            <td>${item.propertyName}</td>
            <td class="${statusClass}">${statusIcon} ${item.status}</td>
            <td>${item.message}</td>
          </tr>
        `;
      });

      html += `
          </tbody>
        </table>
      `;
    }

    // パースエラーがある場合
    if (result.parseErrors && result.parseErrors.length > 0) {
      html += `
        <div class="error-message" style="margin-top: 20px;">
          <h4>⚠ 解析できなかったファイル</h4>
          <ul>
      `;
      result.parseErrors.forEach(err => {
        html += `<li>${err.filename}: ${err.error}</li>`;
      });
      html += `
          </ul>
        </div>
      `;
    }

    resultContent.innerHTML = html;
  } else {
    displayError(result.error || result.message);
  }

  // スクロール
  resultSection.scrollIntoView({ behavior: 'smooth' });
}

// エラー表示
function displayError(message) {
  resultSection.style.display = 'block';
  resultContent.innerHTML = `
    <div class="error-message">
      <h4>✗ エラーが発生しました</h4>
      <p>${message}</p>
    </div>
  `;
  resultSection.scrollIntoView({ behavior: 'smooth' });
}
