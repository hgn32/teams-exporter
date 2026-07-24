// ============================================================
// Teams メッセージ抽出ツール - background service worker
//
// 役割は2つ:
// 1) 画像のbase64化。content script内のfetchはページ側のCORS制約を
//    受けるため、asyncgw.teams.microsoft.com 等の別オリジンから配信される
//    画像は取得できない。service worker内のfetchはhost_permissionsを
//    与えたホストに対してCORSの制約を受けずに実行できるため、
//    content scriptで取得に失敗した画像はここで代わりに取得する。
// 2) 抽出処理そのものの進行管理（状態・ログの保持とHTML生成・ダウンロード）。
//    popup.jsで完結させると、popupが閉じられた瞬間にJS実行が止まり
//    ダウンロードまで到達できなくなる（popupは非表示になると破棄される）。
//    そのためcontent scriptとのやり取り・完了後のHTML生成・ダウンロードは
//    すべてこのservice worker側で行い、popupは「開いている間だけ状態を
//    表示するビュー」として扱う。popupを閉じても抽出自体は継続する。
// 外部サーバーへの送信は一切行わない（GETで画像を読むだけ、生成物はローカル保存のみ）。
// ============================================================

function arrayBufferToBase64(buf) {
  const bytes = new Uint8Array(buf);
  let binary = '';
  const CHUNK = 0x8000; // String.fromCharCodeの引数上限対策
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
  }
  return btoa(binary);
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!msg || msg.type !== 'FETCH_IMAGE_AS_DATA_URI') return;

  (async () => {
    try {
      const res = await fetch(msg.url, { credentials: 'include' });
      if (!res.ok) {
        sendResponse({ ok: false, error: 'HTTP ' + res.status });
        return;
      }
      const contentType = res.headers.get('content-type') || 'image/png';
      // service workerにはFileReaderが無いためarrayBuffer経由で変換する
      const buf = await res.arrayBuffer();
      const dataUri = `data:${contentType};base64,${arrayBufferToBase64(buf)}`;
      sendResponse({ ok: true, dataUri });
    } catch (e) {
      sendResponse({ ok: false, error: String(e && e.message ? e.message : e) });
    }
  })();

  return true; // 非同期でsendResponseを呼ぶために必須
});

// ============================================================
// 抽出処理の状態管理（popupが閉じていても継続する）
// ============================================================

const state = {
  running: false,
  cancelling: false,
  tabId: null,
  log: [],
};

// popupが開いていなければ受け手が無く失敗するだけなので、エラーは無視してよい
function broadcast(msg) {
  try {
    chrome.runtime.sendMessage(msg, () => {
      void chrome.runtime.lastError;
    });
  } catch (e) {
    /* noop */
  }
}

function broadcastStatus() {
  broadcast({ type: 'BG_STATUS', running: state.running, cancelling: state.cancelling });
}

function setLog(line) {
  state.log = [line];
  broadcast({ type: 'BG_LOG', log: state.log });
}

function pushLog(line) {
  state.log.push(line);
  broadcast({ type: 'BG_LOG', log: state.log });
}

// 走査の進捗は行を積み上げず、直前の進捗行を上書きする
// （長時間の抽出でログが際限なく伸びるのを防ぐ）
function updateProgressLog(line) {
  if (state.log.length && state.log[state.log.length - 1].startsWith('収集済み')) {
    state.log[state.log.length - 1] = line;
  } else {
    state.log.push(line);
  }
  broadcast({ type: 'BG_LOG', log: state.log });
}

function sendToContentTab(tabId, message) {
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

async function runExtraction(tabId, tabUrl) {
  state.tabId = tabId;
  state.cancelling = false;
  state.running = true;
  broadcastStatus();
  setLog(
    '抽出開始（自動スクロール中。チャットの長さによっては数十秒〜数分かかります。最大10分で自動終了します。popupを閉じても処理は継続します）...'
  );

  try {
    const res = await sendToContentTab(tabId, { type: 'START_EXTRACT', embedImages: true });

    if (!res || !res.ok) {
      pushLog('エラー: ' + (res && res.error ? res.error : '不明なエラー（content scriptと通信できない可能性）'));
      return;
    }

    const messages = res.messages;

    if (messages.length === 0) {
      pushLog('0件でした。Teamsの画面構造が変わり、content.js のセレクタが合わなくなっている可能性があります。');
      return;
    }

    // 通常時は結果1行だけ。問題があったときだけ詳細を追加で出す
    // （このログを貼れば原因が特定できるように、警告の中身は残す）
    const stats = res.stats;
    const summary = [`完了: ${messages.length}件`];
    if (stats) {
      summary.push(`画像 ${stats.imagesEmbedded}/${stats.imagesTotal}`, `添付 ${stats.fileCards}件`);
    }
    pushLog(summary.join(' / '));

    if (stats) {
      const linkGap = stats.bodyAnchors - stats.bodyAnchorsWithHref;
      const warnings = [];
      if (stats.bodyMissing > 0) warnings.push(`本文未特定 ${stats.bodyMissing}件`);
      if (linkGap > 0) warnings.push(`本文内リンク未解決 ${linkGap}件`);
      if ((stats.imageFailures || []).length > 0) warnings.push(`画像取得失敗 ${stats.imageFailures.length}件`);
      if (warnings.length > 0) {
        pushLog('警告: ' + warnings.join(' / '));
        for (const f of stats.imageFailures || []) {
          pushLog(`  画像失敗: ${f.src} → ${f.error}`);
        }
      }
    }

    // 出力は1つのHTMLファイルだけ（画像はbase64埋め込み、添付はリンク）
    // なので、フォルダは掘らずDownloads直下にそのまま保存する
    const stamp = new Date().toISOString().replace(/[:T]/g, '-').slice(0, 19);
    const filename = `teams_export_${stamp}.html`;

    const saved = await downloadBlobToPath(toHTML(messages, tabUrl, res.pageTitle), filename, 'text/html;charset=utf-8');
    if (saved) {
      pushLog(`ダウンロードを開始しました（Downloads/${filename}）。`);
    } else {
      pushLog(`ダウンロードに失敗しました（${filename}）。`);
    }
  } finally {
    state.running = false;
    state.cancelling = false;
    state.tabId = null;
    broadcastStatus();
  }
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!msg) return;

  // content scriptからの走査進捗のブロードキャスト。popupが閉じていても
  // ここで受け取ってログに積み、次にpopupが開いたときに再現できるようにする
  if (msg.type === 'PROGRESS') {
    if (state.running) updateProgressLog(`収集済み: ${msg.count}件 (走査${msg.round}周目)`);
    return;
  }
  if (msg.type === 'PROGRESS_STAGE') {
    if (state.running) pushLog(msg.stage);
    return;
  }

  if (msg.type === 'GET_STATE') {
    sendResponse({ running: state.running, cancelling: state.cancelling, log: state.log });
    return;
  }

  if (msg.type === 'START_EXTRACT_REQUEST') {
    if (state.running) {
      sendResponse({ ok: false, error: '既に他のタブで抽出処理が実行中です。' });
      return;
    }
    sendResponse({ ok: true });
    runExtraction(msg.tabId, msg.tabUrl);
    return;
  }

  if (msg.type === 'CANCEL_EXTRACT_REQUEST') {
    if (!state.running || !state.tabId) {
      sendResponse({ ok: false });
      return;
    }
    state.cancelling = true;
    broadcastStatus();
    pushLog('中止を要求しました。現在のラウンドの処理が終わり次第停止します...');
    sendToContentTab(state.tabId, { type: 'CANCEL_EXTRACT' });
    sendResponse({ ok: true });
    return;
  }
});
