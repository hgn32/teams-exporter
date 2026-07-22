// ============================================================
// 社内用 Teams メッセージ抽出ツール - content script
//
// 注意: Teams WebのDOM構造は非公開であり、バージョンアップで
// 頻繁に変わる。ここに書いたセレクタは公開されている類似OSS
// (ingo/microsoft-teams-chat-extractor, devnix/msteams-exporter)
// を参考にした「候補」であり、実際の社内テナントのTeamsで動く
// 保証はない。ポップアップの「検出テスト」で実際にヒットする
// 候補を確認し、ヒットしない場合はこのファイルのセレクタを
// DevToolsで確認した実際の値に書き換えてから使うこと。
// ============================================================

const SCROLL_CONTAINER_CANDIDATES = [
  '#chat-pane-list',
  '[data-tid="message-pane-list-viewport"]',
  '[data-tid="channel-pane-list"]',
  '[role="list"][data-tid*="message"]',
];

const MESSAGE_ITEM_CANDIDATES = [
  '[data-tid="chat-pane-message"]',
  '[data-tid="channel-replies-pane-message"]',
  '[data-tid^="message-item"]',
  '[id^="chat-pane-item"]',
  '[data-mid]',
];

const AUTHOR_SELECTOR_CANDIDATES = [
  '[data-tid="message-author-name"]',
  'span.fui-StyledText',
  '[itemprop="author"]',
];

const BODY_SELECTOR_CANDIDATES = [
  '[data-tid="message-body"]',
  '[data-tid="message-body-content"]',
  '.fui-ChatMyMessage__body, .fui-ChatMessage__body',
];

const TIME_SELECTOR_CANDIDATES = ['time', '[data-tid="message-timestamp"]'];

function firstMatch(root, selectors) {
  for (const sel of selectors) {
    try {
      const el = root.querySelector(sel);
      if (el) return el;
    } catch (e) {
      /* 不正なセレクタは無視 */
    }
  }
  return null;
}

function firstMatchingSelector(root, selectors) {
  for (const sel of selectors) {
    try {
      const list = root.querySelectorAll(sel);
      if (list && list.length > 0) return { selector: sel, count: list.length };
    } catch (e) {
      /* noop */
    }
  }
  return null;
}

function findScrollContainer() {
  for (const sel of SCROLL_CONTAINER_CANDIDATES) {
    const el = document.querySelector(sel);
    if (el) return el;
  }
  return null;
}

function findMessageSelector(container) {
  const found = firstMatchingSelector(container || document, MESSAGE_ITEM_CANDIDATES);
  return found ? found.selector : null;
}

function extractOne(node) {
  const authorEl = firstMatch(node, AUTHOR_SELECTOR_CANDIDATES);
  const bodyEl = firstMatch(node, BODY_SELECTOR_CANDIDATES);
  const timeEl = firstMatch(node, TIME_SELECTOR_CANDIDATES);

  const author = authorEl ? authorEl.textContent.trim() : '(不明な送信者)';
  const text = bodyEl ? bodyEl.innerText.trim() : node.innerText.trim();
  const isoTime = timeEl ? timeEl.getAttribute('datetime') || '' : '';
  const displayTime = timeEl ? timeEl.textContent.trim() : '';

  // 安定IDが取れない場合のフォールバック用シグネチャ（重複排除に使用）
  const signature = [author, isoTime || displayTime, text].join('||');

  return { author, text, isoTime, displayTime, signature };
}

function collectVisible(container, messageSelector) {
  const nodes = container.querySelectorAll(messageSelector);
  const results = [];
  nodes.forEach((n) => results.push(extractOne(n)));
  return results;
}

async function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

async function autoScrollAndCollect(onProgress) {
  const container = findScrollContainer();
  if (!container) {
    throw new Error(
      'スクロール対象のコンテナが見つかりません（SCROLL_CONTAINER_CANDIDATES を要確認）'
    );
  }
  const messageSelector = findMessageSelector(container);
  if (!messageSelector) {
    throw new Error(
      'メッセージ要素が見つかりません（MESSAGE_ITEM_CANDIDATES を要確認）'
    );
  }

  const seen = new Map(); // signature -> message
  let idleRounds = 0;
  let lastScrollHeight = -1;
  const MAX_IDLE_ROUNDS = 5; // これ以上高さが変わらなければ終端とみなす
  const MAX_ROUNDS = 2000; // 無限ループ防止の安全弁

  for (let round = 0; round < MAX_ROUNDS; round++) {
    // 現在見えている分を収集
    for (const msg of collectVisible(container, messageSelector)) {
      if (!seen.has(msg.signature)) seen.set(msg.signature, msg);
    }

    onProgress && onProgress(seen.size, round);

    // 先頭（過去方向）へスクロール
    container.scrollTop = 0;
    await sleep(600); // 読み込み待ち。回線が遅い場合は増やす必要あり

    // もう一度収集（スクロール直後に生成された要素を拾う）
    for (const msg of collectVisible(container, messageSelector)) {
      if (!seen.has(msg.signature)) seen.set(msg.signature, msg);
    }

    const currentScrollHeight = container.scrollHeight;
    if (currentScrollHeight === lastScrollHeight && container.scrollTop === 0) {
      idleRounds++;
    } else {
      idleRounds = 0;
    }
    lastScrollHeight = currentScrollHeight;

    if (idleRounds >= MAX_IDLE_ROUNDS) break;
  }

  return Array.from(seen.values());
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.type === 'PING') {
    sendResponse({ ok: true, url: location.href });
    return true;
  }

  if (msg.type === 'DETECT') {
    const container = findScrollContainer();
    const containerSelectorHit = SCROLL_CONTAINER_CANDIDATES.find(
      (s) => document.querySelector(s)
    );
    const messageSelector = container ? findMessageSelector(container) : null;
    const count =
      container && messageSelector
        ? container.querySelectorAll(messageSelector).length
        : 0;
    sendResponse({
      ok: true,
      containerFound: !!container,
      containerSelector: containerSelectorHit || null,
      messageSelector: messageSelector || null,
      visibleMessageCount: count,
    });
    return true;
  }

  if (msg.type === 'START_EXTRACT') {
    (async () => {
      try {
        const messages = await autoScrollAndCollect((count, round) => {
          chrome.runtime.sendMessage({
            type: 'PROGRESS',
            count,
            round,
          });
        });
        // 時系列順に並べる（ISO時刻があればそれで、なければ収集順のまま）
        messages.sort((a, b) => {
          if (a.isoTime && b.isoTime) return a.isoTime.localeCompare(b.isoTime);
          return 0;
        });
        sendResponse({ ok: true, messages });
      } catch (e) {
        sendResponse({ ok: false, error: String(e && e.message ? e.message : e) });
      }
    })();
    return true; // 非同期でsendResponseを呼ぶために必須
  }
});
