const statusEl = document.getElementById('status');
const actionBtn = document.getElementById('actionBtn');

let logLines = [];

function render() {
  statusEl.textContent = logLines.join('\n');
  statusEl.scrollTop = statusEl.scrollHeight;
}

function log(line) {
  logLines.push(line);
  render();
}

function setStatus(line) {
  logLines = [line];
  render();
}

// 走査の進捗は行を積み上げず、直前の進捗行を上書きする
// （長時間の抽出でログが際限なく伸びるのを防ぐ）
function updateProgress(line) {
  if (logLines.length && logLines[logLines.length - 1].startsWith('収集済み')) {
    logLines[logLines.length - 1] = line;
  } else {
    logLines.push(line);
  }
  render();
}

// どの版が動いているか一目で分かるよう、常にバージョンを表示する
// （古い版のまま「直っていない」となる事故を防ぐ）
setStatus(`待機中... (v${chrome.runtime.getManifest().version})`);

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

// 属性値(src/href/alt)にも使うため、"と'も必ずエスケープする
function escapeHtml(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// 拡張子ごとにそれらしいアイコンを添える程度の見た目調整（厳密な
// ファイル種別判定ではない）
function fileIcon(name) {
  const ext = ((String(name).match(/\.([a-z0-9]+)$/i) || [])[1] || '').toLowerCase();
  const map = {
    xlsx: '📊', xls: '📊', csv: '📊',
    docx: '📄', doc: '📄', pdf: '📄', txt: '📄',
    md: '📝',
    pptx: '📈', ppt: '📈',
    svg: '🖼️', png: '🖼️', jpg: '🖼️', jpeg: '🖼️', gif: '🖼️',
    zip: '🗜️',
  };
  return map[ext] || '📎';
}

function toHTML(messages, tabUrl, pageTitle) {
  const rows = messages
    .map((m) => {
      const isReply = m.kind === '返信';
      const kindLabel = m.kind ? `<span class="kind">${escapeHtml(m.kind)}</span>` : '';
      const embeddedImageHtml = (m.images || [])
        .filter((img) => img.dataUri)
        .map((img) => `<div class="embedded-image"><img src="${img.dataUri}" alt="${escapeHtml(img.alt)}"></div>`)
        .join('\n');
      // 埋め込みに失敗した画像・本文外のファイル添付カードを、
      // 同じ見た目のチップとしてまとめて表示する
      const chipHtml = [
        ...(m.images || [])
          .filter((img) => !img.dataUri)
          .map(
            (img) =>
              `<a class="file-chip" href="${escapeHtml(img.src)}" target="_blank" rel="noopener noreferrer"><span class="file-chip-icon">🖼️</span><span class="file-chip-name">${escapeHtml(img.alt)}</span></a>`
          ),
        ...(m.files || []).map(
          (f) =>
            `<a class="file-chip" href="${escapeHtml(f.href)}" target="_blank" rel="noopener noreferrer"><span class="file-chip-icon">${fileIcon(f.name)}</span><span class="file-chip-name">${escapeHtml(f.name)}</span></a>`
        ),
      ].join('\n');
      const attachmentsHtml = chipHtml ? `<div class="attachments">${chipHtml}</div>` : '';
      return `
    <div class="msg${isReply ? ' reply' : ''}">
      <div class="meta">${kindLabel}<span class="author">${escapeHtml(m.author)}</span>
        <span class="time">${escapeHtml(m.displayTime || m.isoTime)}</span></div>
      <div class="body">${m.bodyHtml || escapeHtml(m.text).replace(/\n/g, '<br>')}</div>
      ${embeddedImageHtml}
      ${attachmentsHtml}
    </div>`;
    })
    .join('\n');

  const titleText = pageTitle || 'Teams チャット抽出結果';

  return `<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8">
<title>${escapeHtml(titleText)} - Teams抽出結果</title>
<style>
  body { font-family: sans-serif; max-width: 800px; margin: 20px auto; color: #222; }
  .msg { border: 1px solid #ddd; border-radius: 8px; padding: 12px 14px; margin-bottom: 10px; }
  .msg.reply { margin-left: 32px; background: #f7f9fc; border-color: #cdd7e6; }
  .meta { font-size: 12px; color: #666; margin-bottom: 4px; }
  .kind { display: inline-block; font-size: 11px; color: #555; background: #eee; border-radius: 4px; padding: 1px 6px; margin-right: 6px; }
  .author { font-weight: bold; color: #333; margin-right: 8px; }
  .body { white-space: pre-wrap; }
  .embedded-image { margin-top: 8px; }
  .embedded-image img { max-width: 100%; border-radius: 4px; }
  .attachments { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 8px; }
  .file-chip {
    display: inline-flex; align-items: center; gap: 6px;
    padding: 5px 10px; max-width: 100%;
    border: 1px solid #d7dee8; border-radius: 14px;
    background: #f3f6fb; color: #2b579a;
    font-size: 12px; text-decoration: none;
  }
  .file-chip:hover { background: #e8eef7; border-color: #b9c7db; }
  .file-chip-icon { font-size: 13px; line-height: 1; }
  .file-chip-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 320px; }
  .source { font-size: 11px; color: #999; margin-bottom: 16px; }
</style></head>
<body>
  <h2>${escapeHtml(titleText)}</h2>
  <div class="source">抽出元: ${escapeHtml(tabUrl)} / 抽出日時: ${new Date().toLocaleString('ja-JP')} / 件数: ${messages.length}</div>
  ${rows}
</body></html>`;
}

function downloadBlobToPath(content, relPath, mime) {
  return new Promise((resolve) => {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    chrome.downloads.download({ url, filename: relPath, saveAs: false }, () => {
      setTimeout(() => URL.revokeObjectURL(url), 10000);
      resolve(!chrome.runtime.lastError);
    });
  });
}

let extractTabId = null;
let extracting = false;

async function cancelExtraction() {
  if (!extractTabId) return;
  actionBtn.disabled = true;
  actionBtn.textContent = '中止中...';
  log('中止を要求しました。現在のラウンドの処理が終わり次第停止します...');
  await sendToContent(extractTabId, { type: 'CANCEL_EXTRACT' });
}

async function startExtraction() {
  const tab = await getActiveTab();
  if (!tab || !tab.url || !/teams\.(microsoft\.com|live\.com|cloud\.microsoft)/.test(tab.url)) {
    setStatus('Teams のタブを開いた状態で実行してください。');
    return;
  }

  extracting = true;
  extractTabId = tab.id;
  actionBtn.textContent = '抽出を中止';
  setStatus('抽出開始（自動スクロール中。チャットの長さによっては数十秒〜数分かかります。最大10分で自動終了します）...');

  const progressListener = (msg) => {
    if (msg.type === 'PROGRESS') {
      updateProgress(`収集済み: ${msg.count}件 (走査${msg.round}周目)`);
    } else if (msg.type === 'PROGRESS_STAGE') {
      log(msg.stage);
    }
  };
  chrome.runtime.onMessage.addListener(progressListener);

  const res = await sendToContent(tab.id, { type: 'START_EXTRACT', embedImages: true });

  chrome.runtime.onMessage.removeListener(progressListener);
  extracting = false;
  extractTabId = null;
  actionBtn.disabled = false;
  actionBtn.textContent = '自動スクロール抽出を開始';

  if (!res || !res.ok) {
    log('エラー: ' + (res && res.error ? res.error : '不明なエラー（content scriptと通信できない可能性）'));
    return;
  }

  const messages = res.messages;

  if (messages.length === 0) {
    log('0件でした。Teamsの画面構造が変わり、content.js のセレクタが合わなくなっている可能性があります。');
    return;
  }

  // 通常時は結果1行だけ。問題があったときだけ詳細を追加で出す
  // （このログを貼れば原因が特定できるように、警告の中身は残す）
  const stats = res.stats;
  const summary = [`完了: ${messages.length}件`];
  if (stats) {
    summary.push(`画像 ${stats.imagesEmbedded}/${stats.imagesTotal}`, `添付 ${stats.fileCards}件`);
  }
  log(summary.join(' / '));

  if (stats) {
    const linkGap = stats.bodyAnchors - stats.bodyAnchorsWithHref;
    const warnings = [];
    if (stats.bodyMissing > 0) warnings.push(`本文未特定 ${stats.bodyMissing}件`);
    if (linkGap > 0) warnings.push(`本文内リンク未解決 ${linkGap}件`);
    if ((stats.imageFailures || []).length > 0) warnings.push(`画像取得失敗 ${stats.imageFailures.length}件`);
    if (warnings.length > 0) {
      log('警告: ' + warnings.join(' / '));
      for (const f of stats.imageFailures || []) {
        log(`  画像失敗: ${f.src} → ${f.error}`);
      }
    }
  }

  // 出力は1つのHTMLファイルだけ（画像はbase64埋め込み、添付はリンク）
  // なので、フォルダは掘らずDownloads直下にそのまま保存する
  const stamp = new Date().toISOString().replace(/[:T]/g, '-').slice(0, 19);
  const filename = `teams_export_${stamp}.html`;

  const saved = await downloadBlobToPath(toHTML(messages, tab.url, res.pageTitle), filename, 'text/html;charset=utf-8');
  if (saved) {
    log(`ダウンロードを開始しました（Downloads/${filename}）。`);
  } else {
    log(`ダウンロードに失敗しました（${filename}）。`);
  }
}

actionBtn.addEventListener('click', () => {
  if (extracting) {
    cancelExtraction();
  } else {
    startExtraction();
  }
});
