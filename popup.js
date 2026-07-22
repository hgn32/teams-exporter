const statusEl = document.getElementById('status');
const detectBtn = document.getElementById('detectBtn');
const extractBtn = document.getElementById('extractBtn');
const formatSelect = document.getElementById('formatSelect');

function log(line) {
  statusEl.textContent += '\n' + line;
  statusEl.scrollTop = statusEl.scrollHeight;
}

function setStatus(line) {
  statusEl.textContent = line;
}

async function getActiveTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab;
}

function sendToContent(tabId, message) {
  return new Promise((resolve) => {
    chrome.tabs.sendMessage(tabId, message, (response) => {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, error: chrome.runtime.lastError.message });
        return;
      }
      resolve(response);
    });
  });
}

detectBtn.addEventListener('click', async () => {
  setStatus('検出中...');
  const tab = await getActiveTab();
  if (!tab || !tab.url || !/teams\.(microsoft\.com|live\.com|cloud\.microsoft)/.test(tab.url)) {
    setStatus('Teams のタブ（teams.microsoft.com 等）を開いた状態で実行してください。');
    return;
  }
  const res = await sendToContent(tab.id, { type: 'DETECT' });
  if (!res || !res.ok) {
    setStatus(
      'content scriptと通信できませんでした。ページを再読み込みしてから再実行してください。\n詳細: ' +
        (res && res.error ? res.error : '不明なエラー')
    );
    return;
  }
  setStatus(
    [
      `コンテナ検出: ${res.containerFound ? 'OK' : 'NG'}`,
      `コンテナセレクタ: ${res.containerSelector || '(なし)'}`,
      `メッセージセレクタ: ${res.messageSelector || '(なし)'}`,
      `現在表示中のメッセージ数: ${res.visibleMessageCount}`,
      '',
      res.visibleMessageCount > 0
        ? '→ 検出できています。抽出を実行できます。'
        : '→ 検出できていません。content.js のセレクタ候補を、DevToolsで確認した実際の値に修正してください。',
    ].join('\n')
  );
});

function toCSV(messages) {
  const header = ['送信者', '日時(ISO)', '表示日時', '本文'];
  const escape = (s) => '"' + String(s || '').replace(/"/g, '""') + '"';
  const lines = [header.map(escape).join(',')];
  for (const m of messages) {
    lines.push(
      [m.author, m.isoTime, m.displayTime, m.text].map(escape).join(',')
    );
  }
  // Excelで文字化けしないようUTF-8 BOMを付与
  return '\uFEFF' + lines.join('\r\n');
}

function escapeHtml(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function toHTML(messages, tabUrl) {
  const rows = messages
    .map(
      (m) => `
    <div class="msg">
      <div class="meta"><span class="author">${escapeHtml(m.author)}</span>
        <span class="time">${escapeHtml(m.displayTime || m.isoTime)}</span></div>
      <div class="body">${escapeHtml(m.text).replace(/\n/g, '<br>')}</div>
    </div>`
    )
    .join('\n');

  return `<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8">
<title>Teams チャット抽出結果</title>
<style>
  body { font-family: sans-serif; max-width: 800px; margin: 20px auto; color: #222; }
  .msg { border-bottom: 1px solid #eee; padding: 10px 0; }
  .meta { font-size: 12px; color: #666; margin-bottom: 4px; }
  .author { font-weight: bold; color: #333; margin-right: 8px; }
  .body { white-space: pre-wrap; }
  .source { font-size: 11px; color: #999; margin-bottom: 16px; }
</style></head>
<body>
  <h2>Teams チャット抽出結果</h2>
  <div class="source">抽出元: ${escapeHtml(tabUrl)} / 抽出日時: ${new Date().toLocaleString('ja-JP')} / 件数: ${messages.length}</div>
  ${rows}
</body></html>`;
}

function download(content, filename, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

extractBtn.addEventListener('click', async () => {
  const tab = await getActiveTab();
  if (!tab || !tab.url || !/teams\.(microsoft\.com|live\.com|cloud\.microsoft)/.test(tab.url)) {
    setStatus('Teams のタブを開いた状態で実行してください。');
    return;
  }

  extractBtn.disabled = true;
  detectBtn.disabled = true;
  setStatus('抽出開始（自動スクロール中。チャットの長さによっては数十秒〜数分かかります）...');

  const progressListener = (msg) => {
    if (msg.type === 'PROGRESS') {
      log(`収集済み: ${msg.count}件 (走査${msg.round}周目)`);
    }
  };
  chrome.runtime.onMessage.addListener(progressListener);

  const res = await sendToContent(tab.id, { type: 'START_EXTRACT' });

  chrome.runtime.onMessage.removeListener(progressListener);
  extractBtn.disabled = false;
  detectBtn.disabled = false;

  if (!res || !res.ok) {
    log('エラー: ' + (res && res.error ? res.error : '不明なエラー（content scriptと通信できない可能性）'));
    return;
  }

  const messages = res.messages;
  log(`完了。合計 ${messages.length} 件のメッセージを取得しました。`);

  if (messages.length === 0) {
    log('0件でした。検出テストでセレクタが正しくヒットしているか確認してください。');
    return;
  }

  const format = formatSelect.value;
  const stamp = new Date().toISOString().replace(/[:T]/g, '-').slice(0, 19);
  if (format === 'csv') {
    download(toCSV(messages), `teams_export_${stamp}.csv`, 'text/csv;charset=utf-8');
  } else {
    download(toHTML(messages, tab.url), `teams_export_${stamp}.html`, 'text/html;charset=utf-8');
  }
  log('ダウンロードを開始しました。');
});
