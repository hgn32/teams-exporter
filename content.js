// ============================================================
// Teams メッセージ抽出ツール - content script
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
  '[data-tid="channel-pane-viewport"]',
  '[data-tid="message-pane-list-viewport"]',
  '[role="list"][data-tid*="message"]',
];

// チャネルでは、投稿(channel-pane-message)と返信(id="message-body-*"かつ
// reply-message-headerを含むもの)の2種類をまとめて1つのセレクタで拾う。
// :has()はChrome/Edge双方の対象バージョンでサポート済み
const MESSAGE_ITEM_CANDIDATES = [
  '[data-tid="chat-pane-message"]',
  '[data-tid="channel-pane-message"], [id^="message-body-"]:has([data-tid="reply-message-header"])',
  '[data-tid^="message-item"]',
  '[id^="chat-pane-item"]',
  '[data-mid]',
];

// 診断機能の実測で判明した構造: 送信者名・時刻は個々のメッセージ要素
// ([data-tid="chat-pane-message"]) の「中」ではなく、その2階層上にある
// グループコンテナ内の、メッセージ本体より前にある兄弟divに入っている。
// GROUP_CONTAINER_SELECTOR_CANDIDATESでそのコンテナを見つけてから、
// その中でAUTHOR/TIMEを検索する。
//
// 重要: Teamsは「自分が送信したメッセージ」と「他人が送信したメッセージ」で
// クラス名が異なり、自分のメッセージ側には"My"が挟まる
// （例: fui-ChatMessage__author ⇔ fui-ChatMyMessage__author）。
// 両方をカバーする必要がある。
const GROUP_CONTAINER_SELECTOR_CANDIDATES = [
  '.fui-ChatMessage, .fui-ChatMyMessage',
];

const AUTHOR_SELECTOR_CANDIDATES = [
  '.fui-ChatMessage__author, .fui-ChatMyMessage__author',
  '[data-tid="message-author-name"]',
  '[itemprop="author"]',
];
// 注意: 'span.fui-StyledText' は以前候補に含めていたが、Fluent UIの汎用スタイル
// クラスでメッセージ本文中の任意のspanにも付与されるため、送信者名の代わりに
// 本文の断片やリンクプレビューのタイトルを誤って拾ってしまう実害を確認した。
// 候補としては使わないこと。

// 本文はメッセージ要素自身に'fui-ChatMessage__body'クラスが付いているだけ
// で（querySelectorは要素自身にはマッチしないため検出不可）、実体は
// data-message-content属性を持つ内側のdivにある
const BODY_SELECTOR_CANDIDATES = [
  '[data-message-content]',
  '[data-tid="message-body"]',
  '[data-tid="message-body-content"]',
];

const TIME_SELECTOR_CANDIDATES = [
  '.fui-ChatMessage__timestamp, .fui-ChatMyMessage__timestamp',
  'time[id^="timestamp-"]',
  'time',
  '[data-tid="message-timestamp"]',
];

// チャット/チャネル名（HTML出力のヘッダーに使う）。Teams Web側のヘッダー
// 要素は未確認のため候補は推測込み。どれもヒットしない場合はタブタイトル
// を使う。実測で判明した形式は「(未読件数) チャット | <実際の会話名>」
// （例:「(6) チャット | ADサーバのRMF管理策検討のご依頼」）で、
// Microsoft Teamsという末尾は付かない。最後の"|"より後ろが実際の名前
const TITLE_SELECTOR_CANDIDATES = [
  '[data-tid="pageHeaderTitle"]',
  '[data-tid="channel-header-title"]',
  '[data-tid="chat-header-title"]',
  '[data-tid="thread-title"]',
  'h1[data-tid="title"]',
];

function extractPageTitle() {
  for (const sel of TITLE_SELECTOR_CANDIDATES) {
    try {
      const el = document.querySelector(sel);
      if (el && el.textContent.trim()) {
        return { title: el.textContent.trim(), source: sel };
      }
    } catch (e) {
      /* noop */
    }
  }
  const raw = (document.title || '').trim()
    .replace(/\s*[|\-–]\s*Microsoft Teams\s*$/i, '') // 末尾に付く場合のみ除去
    .trim();
  // "(6) チャット | 会話名" のように"|"区切りなら、最後の"|"より後ろが
  // 実際の会話名。区切りが無ければ先頭の"(未読件数) "だけ除去する
  const pipeIdx = raw.lastIndexOf('|');
  const cleaned =
    pipeIdx >= 0 ? raw.slice(pipeIdx + 1).trim() : raw.replace(/^\(\d+\)\s*/, '').trim();
  return { title: cleaned || raw || '(タイトル取得失敗)', source: 'document.title' };
}

function findGroupContainer(node) {
  for (const sel of GROUP_CONTAINER_SELECTOR_CANDIDATES) {
    try {
      const el = node.closest(sel);
      if (el) return el;
    } catch (e) {
      /* noop */
    }
  }
  return null;
}

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

// 実測で判明: 添付ファイルは本文中に挿入された通常の
// <a class="fui-Link" href="...">（表示テキスト=ファイル名）として
// 出ることが多いが、本文要素の外側にファイルカードとして描画される
// パターンもある。SharePoint/OneDriveの共有リンクかどうかで、
// 本文に貼られただけの外部URL（単なるハイパーリンク）と区別する
function isSharePointFileLink(href) {
  try {
    const url = new URL(href);
    return (
      /(^|\.)sharepoint\.com$/i.test(url.hostname) ||
      /(^|\.)1drv\.ms$/i.test(url.hostname) ||
      /\/:[a-z]:\//i.test(url.pathname)
    );
  } catch (e) {
    return false;
  }
}

// Teamsの遅延読み込みは、本物のURLに差し替わる前の一瞬、1x1透明GIF等の
// ごく短いプレースホルダをsrcに入れている。これをそのまま「解決済み」と
// 扱うと、以降の全ラウンド・fetch処理がスキップされてプレースホルダの
// まま固定されてしまう（実物のsrcに差し替わった後の再スクレイプが
// 生かされない）。実物のdata URIは通常もっと長いため、極端に短い
// data:URIはプレースホルダとみなし「未解決」扱いにする
function isPlaceholderDataUri(src) {
  return src.length < 200;
}

// アニメーション絵文字等は、テーマ・モーション設定違いの複数バリアント
// （静止画/アニメーション、ライト/ダーク等）が同時にDOM上へ描画され、
// CSSで表示側だけを残す実装になっていることがある。抽出時にCSSの
// 出し分けは反映されないため、非表示のバリアントも構わず拾ってしまうと
// 「同じ絵文字が何十個も並ぶ」ような重複が発生する。実際に表示されている
// 要素だけを対象にすることでこれを防ぐ
function isRenderedVisible(el) {
  if (!el.isConnected) return false;
  const style = window.getComputedStyle(el);
  if (style.display === 'none' || style.visibility === 'hidden' || parseFloat(style.opacity) === 0) {
    return false;
  }
  const rect = el.getBoundingClientRect();
  return rect.width > 0 && rect.height > 0;
}

// 画面上に描画済みの<img>要素をその場でcanvas経由でdata URIにする。
// blob: URL（後で失効するためfetchでは間に合わないことがある）や
// 読み込み済み同一オリジン画像はこれで即座に確保できる。
// クロスオリジン画像はcanvasが汚染されtoDataURLが例外を投げるため
// nullを返し、後段のfetch（content script→background）に委ねる
function imgToDataUriViaCanvas(img) {
  try {
    if (!img.complete || !img.naturalWidth) return null;
    const canvas = document.createElement('canvas');
    canvas.width = img.naturalWidth;
    canvas.height = img.naturalHeight;
    canvas.getContext('2d').drawImage(img, 0, 0);
    return canvas.toDataURL('image/png');
  } catch (e) {
    return null;
  }
}

// 画像は本文中に埋め込まれた img として表示されるだけで、本文の
// innerText では拾えない（画像はテキストを持たない）。
//
// 添付ファイルのリンクは実測で3パターン確認されている:
//  (A) 本文(bodyEl)の中に <a href> として挿入されるもの
//      → bodyToSafeHtml が本文HTMLにそのまま復元するので別途出さない
//  (B) 本文の外側に、通常の<a href>を持つファイルカードとして描画されるもの
//      → bodyHtmlには乗らないので、ここで拾ってfiles[]に入れる
//  (C) 本文の外側の[data-tid="file-attachment-grid"]配下に、<a href>を
//      一切持たない独自コンポーネントとして描画されるもの
//      （ファイル名とURLは[data-testid="content-card-custom-title"]の
//      aria-labelに"ファイル名\nURL"の形式で埋め込まれているだけ）
//      → 別途DOMを検証して確認した実際の構造。ここで拾ってfiles[]に入れる
// searchRoot（=メッセージ要素全体）を走査し、bodyElの中にあるものは
// 除外して(B)(C)だけを返す（(A)を二重表示しないため）
function extractAttachments(bodyEl, searchRoot) {
  const images = [];
  const files = [];
  // 添付リンクが消える原因切り分け用の診断カウンタ。本文中に<a>が
  // 何個あり、そのうちhrefを持つものが何個か（未hydrateで0になって
  // いないか）を集計する
  const linkStats = { bodyAnchors: 0, bodyAnchorsWithHref: 0 };

  if (bodyEl) {
    bodyEl.querySelectorAll('a').forEach((a) => {
      linkStats.bodyAnchors++;
      if (a.href) linkStats.bodyAnchorsWithHref++;
    });
  }

  if (bodyEl) {
    const seenImageSrc = new Set();
    bodyEl.querySelectorAll('img').forEach((img) => {
      if (!isRenderedVisible(img)) return; // アニメーション絵文字等の非表示バリアントは除外
      const alt = img.getAttribute('alt') || img.getAttribute('title') || '';
      if (!img.src) return;
      if (seenImageSrc.has(img.src)) return; // 同じ画像の重複要素は1つにまとめる
      seenImageSrc.add(img.src);
      const entry = { src: img.src, alt: alt || '(ファイル名不明)' };
      if (img.src.startsWith('data:')) {
        // 遅延読み込みのプレースホルダはまだ「未解決」として扱い、
        // dataUriを確定させない（確定させると以降のラウンドで実物の
        // srcに差し替わっても再取得されなくなるため）
        if (!isPlaceholderDataUri(img.src)) entry.dataUri = img.src;
      } else {
        const dataUri = imgToDataUriViaCanvas(img);
        if (dataUri) entry.dataUri = dataUri;
      }
      images.push(entry);
    });
  }

  const root = searchRoot || bodyEl;
  if (root) {
    const seenHref = new Set();

    root.querySelectorAll('a[href]').forEach((a) => {
      if (a.closest('[data-person-mri]')) return; // @メンションは添付ではない
      if (!isSharePointFileLink(a.href)) return; // 本文に貼られただけの外部URLは対象外
      if (bodyEl && bodyEl.contains(a)) return; // 本文内のリンクは(A)なので二重表示しない
      if (seenHref.has(a.href)) return;
      seenHref.add(a.href);
      files.push({ name: a.textContent.trim() || '(ファイル名不明)', href: a.href });
    });

    // 実測で判明した(C): ファイル添付カード([data-tid="file-attachment-grid"]
    // 配下)は<a href>を持たない独自コンポーネントで、ファイル名とURLは
    // [data-testid="content-card-custom-title"]のaria-labelに
    // "ファイル名\nURL" の形式で埋め込まれている（見た目のテキストには出ない）
    root
      .querySelectorAll('[data-tid="file-attachment-grid"] [data-testid="content-card-custom-title"]')
      .forEach((el) => {
        const label = el.getAttribute('aria-label') || '';
        const lines = label.split('\n').map((s) => s.trim()).filter(Boolean);
        if (lines.length < 2) return; // URLが乗っていない場合は対象外
        const name = lines[0] || '(ファイル名不明)';
        const href = lines[lines.length - 1];
        if (!/^https?:\/\//i.test(href)) return;
        if (seenHref.has(href)) return;
        seenHref.add(href);
        files.push({ name, href });
      });
  }

  return { images, files, linkStats };
}

// 本文中の通常のリンク（外部URL等）はHTML出力でもクリックできる
// ようにする。安全のため、許可した最小限のタグ（a/br）だけHTMLとして
// 復元し、それ以外は全てエスケープ済みテキストとして流し込む
function escapeHtmlText(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escapeHtmlAttr(s) {
  return escapeHtmlText(s).replace(/"/g, '&quot;');
}

const BLOCK_TAGS = new Set(['p', 'div', 'li', 'tr', 'blockquote']);

function bodyToSafeHtml(bodyEl) {
  if (!bodyEl) return '';
  function walk(node) {
    if (node.nodeType === 3) {
      return escapeHtmlText(node.textContent);
    }
    if (node.nodeType !== 1) return '';
    const tag = node.tagName.toLowerCase();
    if (tag === 'img') return ''; // 画像は別枠(images)で表示するため本文中はスキップ
    if (tag === 'br') return '<br>';
    const childHtml = Array.from(node.childNodes).map(walk).join('');
    if (tag === 'a' && node.href) {
      return `<a href="${escapeHtmlAttr(node.href)}" target="_blank" rel="noopener noreferrer">${childHtml}</a>`;
    }
    return BLOCK_TAGS.has(tag) ? childHtml + '<br>' : childHtml;
  }
  return Array.from(bodyEl.childNodes).map(walk).join('');
}

function blobToDataUri(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

// 画像URLをbase64のdata URIに変換する。2段構え:
// 1) content script内のfetch（同一オリジン・CORS許可済みの画像用）
// 2) 失敗したらbackground service worker経由のfetch
//    （content scriptのfetchはページ側のCORS制約を受けるため、
//     asyncgw等の別オリジン画像はここでしか取れない）
// 失敗時は{ error }を返し、理由を必ず呼び出し側に伝える（無言で
// 握りつぶすと利用者側で原因調査ができないため）
async function fetchImageAsDataUri(url) {
  // 遅延読み込みのプレースホルダのまま最後まで残った画像（スクロールが
  // 速すぎて実物のsrcに一度も差し替わらなかったケース）。これをfetchすると
  // プレースホルダ自身が「成功」扱いで埋め込まれてしまうため、実際には
  // 未解決だったと明示するエラーにする
  if (url.startsWith('data:') && isPlaceholderDataUri(url)) {
    return { error: '遅延読み込みのプレースホルダのまま実画像に解決されませんでした' };
  }

  let firstError = '';
  try {
    const res = await fetch(url, { credentials: 'include' });
    if (res.ok) {
      const blob = await res.blob();
      return { dataUri: await blobToDataUri(blob) };
    }
    firstError = 'HTTP ' + res.status;
  } catch (e) {
    firstError = String(e && e.message ? e.message : e);
  }

  // blob:/data: はページ内でしか解決できないのでbackgroundには投げない
  if (/^(blob|data):/.test(url)) {
    return { error: firstError };
  }

  const bg = await new Promise((resolve) => {
    chrome.runtime.sendMessage({ type: 'FETCH_IMAGE_AS_DATA_URI', url }, (r) => {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, error: chrome.runtime.lastError.message });
        return;
      }
      resolve(r || { ok: false, error: 'backgroundから応答なし' });
    });
  });
  if (bg.ok) return { dataUri: bg.dataUri };
  return { error: `page: ${firstError} / bg: ${bg.error}` };
}

// 抽出済みメッセージ群の画像をまとめてbase64化する。スクレイプ時点で
// canvas経由で確保済みのもの（dataUri設定済み）はスキップ。
// 取得できなかった画像はdataUriを付けず失敗理由をerrorに記録し、
// 呼び出し側（HTML生成側）で元URLへのリンク表示にフォールバックさせる
async function embedImagesAsBase64(messages) {
  const tasks = [];
  for (const m of messages) {
    for (const img of m.images || []) {
      if (img.dataUri) continue;
      tasks.push(
        fetchImageAsDataUri(img.src).then((result) => {
          if (result.dataUri) img.dataUri = result.dataUri;
          else img.error = result.error || '不明なエラー';
        })
      );
    }
  }
  await Promise.all(tasks);
}

// チャネル（投稿・スレッド返信）の場合、送信者名(span[id="author-{msgId}"])・
// 時刻([data-tid="timestamp"] id="timestamp-{msgId}")・本文
// (div[data-message-content] id="content-{msgId}")が、DOM上の親子関係では
// なく「同じmsgId」で結び付いているだけで、離れた場所に存在する。
// そのためメッセージ要素自身のid末尾の数字を取り出し、
// document.getElementByIdで直接該当要素を引く
function isChannelMessageNode(node) {
  return (
    node.matches('[data-tid="channel-pane-message"]') ||
    (node.id && node.id.startsWith('message-body-'))
  );
}

function extractChannelMessage(node) {
  const idMatch = node.id && node.id.match(/(\d+)$/);
  const msgId = idMatch ? idMatch[1] : null;

  // msgIdでの直接参照が基本だが、投稿要素のidに数字が付かない・
  // 対応するid付き要素がまだ描画されていない等で取れないことがある。
  // その場合はメッセージ要素内をセレクタ候補で検索するフォールバックを
  // 使う（bodyElがnullのままだと本文がinnerText化されてリンク・画像が
  // 全て失われるため、ここが添付リンク消失の直接原因になる）
  const authorEl =
    (msgId ? document.getElementById('author-' + msgId) : null) ||
    firstMatch(node, AUTHOR_SELECTOR_CANDIDATES);
  const timeEl =
    (msgId ? document.getElementById('timestamp-' + msgId) : null) ||
    firstMatch(node, TIME_SELECTOR_CANDIDATES);
  const bodyEl =
    (msgId ? document.getElementById('content-' + msgId) : null) ||
    firstMatch(node, BODY_SELECTOR_CANDIDATES);

  const author = authorEl ? authorEl.textContent.trim() : '(不明な送信者)';
  const { images, files, linkStats } = extractAttachments(bodyEl, node);
  const text = bodyEl ? bodyEl.innerText.trim() : node.innerText.trim();
  const bodyHtml = bodyEl ? bodyToSafeHtml(bodyEl) : escapeHtmlText(text);
  const displayTime = timeEl ? timeEl.textContent.trim() : '';

  // チャネルのtimestamp要素にはdatetime属性が無いが、msgId自体が
  // ミリ秒単位のUnixタイムスタンプになっているため、それを時系列の
  // 並べ替えキーとして使う（スクロールで遡って見つけた順ではなく、
  // Teams表示と同じ「古い→新しい」の順に並べるため）
  let isoTime = '';
  if (msgId) {
    const asMillis = Number(msgId);
    if (!Number.isNaN(asMillis) && asMillis > 1000000000000) {
      try {
        isoTime = new Date(asMillis).toISOString();
      } catch (e) {
        isoTime = '';
      }
    }
  }
  if (!isoTime && timeEl) {
    isoTime = timeEl.getAttribute('datetime') || '';
  }

  // 投稿(channel-pane-message)か返信(reply-message-headerを含むmessage-body-*)かを区別
  const kind = node.matches('[data-tid="channel-pane-message"]') ? '投稿' : '返信';

  const signature = [author, isoTime || displayTime, text].join('||');
  return {
    author, text, bodyHtml, isoTime, displayTime, kind, images, files, signature,
    noBody: !bodyEl, linkStats,
  };
}

function extractOne(node, groupState) {
  if (isChannelMessageNode(node)) {
    return extractChannelMessage(node);
  }

  // 送信者名・時刻はメッセージ要素の中ではなく、2階層上の
  // グループコンテナ側にあるため、そちらを検索範囲にする
  const groupContainer = findGroupContainer(node);
  const headerSearchRoot = groupContainer || node;

  const authorEl = firstMatch(headerSearchRoot, AUTHOR_SELECTOR_CANDIDATES);
  const bodyEl = firstMatch(node, BODY_SELECTOR_CANDIDATES);
  const timeEl = firstMatch(headerSearchRoot, TIME_SELECTOR_CANDIDATES);

  const { images, files, linkStats } = extractAttachments(bodyEl, node);
  const text = bodyEl ? bodyEl.innerText.trim() : node.innerText.trim();
  const bodyHtml = bodyEl ? bodyToSafeHtml(bodyEl) : escapeHtmlText(text);

  // グループコンテナ側にヘッダーが無い（＝連続メッセージの2件目以降、
  // または想定外の構造）場合のフォールバックとして、直前メッセージの
  // 値を引き継ぐ
  let author = authorEl ? authorEl.textContent.trim() : '';
  if (author) {
    groupState.lastAuthor = author;
  } else {
    author = groupState.lastAuthor || '(不明な送信者)';
  }

  let isoTime = timeEl ? timeEl.getAttribute('datetime') || '' : '';
  let displayTime = timeEl ? timeEl.textContent.trim() : '';
  if (isoTime || displayTime) {
    groupState.lastIsoTime = isoTime;
    groupState.lastDisplayTime = displayTime;
  } else {
    isoTime = groupState.lastIsoTime || '';
    displayTime = groupState.lastDisplayTime || '';
  }

  // 安定IDが取れない場合のフォールバック用シグネチャ（重複排除に使用）
  const signature = [author, isoTime || displayTime, text].join('||');

  return {
    author, text, bodyHtml, isoTime, displayTime, kind: '', images, files, signature,
    noBody: !bodyEl, linkStats,
  };
}

function collectVisible(container, messageSelector) {
  const nodes = container.querySelectorAll(messageSelector);
  const results = [];
  // 送信者名・時刻の引き継ぎはDOM順（=表示順）に依存するため、
  // 収集のたびにこのラウンド内で先頭からリセットして辿る
  const groupState = { lastAuthor: '', lastIsoTime: '', lastDisplayTime: '' };
  nodes.forEach((n) => results.push(extractOne(n, groupState)));
  return results;
}

// Teamsの仮想リストは、要素の描画直後は添付リンクのhrefが未設定
// （後から非同期に解決される）ことがある。シグネチャは本文テキスト
// ベースなので、href未設定のスナップショットとhref解決後のスナップ
// ショットは同一メッセージ扱いになる。「先勝ち」で固定すると
// リンク・画像の無い方が永久に残るため、情報量の多い方を採用する
// dataUri確定済み、またはプレースホルダではないsrc（https/blob/十分に
// 長いdata URI）を持つ画像は「解決済み」とみなす。プレースホルダの
// ままの画像は、後のラウンドで実物に差し替わったスナップショットに
// 確実に負けるよう、richnessの加点をごく小さくする
function isResolvedImage(img) {
  if (img.dataUri) return true;
  if (typeof img.src === 'string' && img.src.startsWith('data:')) {
    return !isPlaceholderDataUri(img.src);
  }
  return true;
}

function snapshotRichness(m) {
  const hrefCount = m.linkStats ? m.linkStats.bodyAnchorsWithHref : 0;
  const imgs = m.images || [];
  const resolvedImages = imgs.filter(isResolvedImage).length;
  const unresolvedImages = imgs.length - resolvedImages;
  return (
    (m.bodyHtml ? m.bodyHtml.length : 0) +
    resolvedImages * 100 +
    unresolvedImages * 1 +
    (m.files ? m.files.length : 0) * 50 +
    hrefCount * 100 +
    (m.noBody ? 0 : 1)
  );
}

function addMessage(seen, msg) {
  const prev = seen.get(msg.signature);
  if (!prev) {
    seen.set(msg.signature, msg);
    return;
  }
  if (snapshotRichness(msg) > snapshotRichness(prev)) {
    // canvas等で先に確保済みのdataUriは、同じsrcなら引き継ぐ
    for (const img of msg.images || []) {
      if (img.dataUri) continue;
      const prevImg = (prev.images || []).find((p) => p.src === img.src && p.dataUri);
      if (prevImg) img.dataUri = prevImg.dataUri;
    }
    seen.set(msg.signature, msg);
  }
}

async function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

// スクロール後、Teams側の過去メッセージ取得（ネットワーク通信）が
// 反映されるのを待つ。すぐに高さが変わらなくても、timeoutMs内に
// 変化があれば「まだ読み込み中」とみなして待ち続ける。
async function waitForContainerGrowth(container, timeoutMs, pollMs) {
  const start = Date.now();
  const initialHeight = container.scrollHeight;
  while (Date.now() - start < timeoutMs) {
    await sleep(pollMs);
    if (container.scrollHeight !== initialHeight) return true;
  }
  return false;
}

let extractionCancelled = false;

async function autoScrollAndCollect(onProgress) {
  extractionCancelled = false;
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
  // 終了判定は「新規メッセージが拾えたか」の実件数ベースなので、
  // scrollHeightの揺れに惑わされていた頃と違い、閾値を高くする必要はない。
  // 少なすぎると遅い回線で本当に読み込み中なのに打ち切ってしまうので、
  // 数ラウンド分の余裕は残しつつ、無駄な周回を減らす
  const MAX_IDLE_ROUNDS = 2;
  const MAX_ROUNDS = 2000; // 無限ループ防止の安全弁（回数）
  const MAX_TOTAL_MS = 10 * 60 * 1000; // 無限ループ防止の安全弁（合計時間・10分）
  const GROWTH_TIMEOUT_MS = 2000; // 過去メッセージ取得の通信待ち上限
  const GROWTH_POLL_MS = 250;
  const startedAt = Date.now();

  // 開いた時点でチャット/チャネルが最新メッセージ（一番下）まで
  // 表示されているとは限らない。以降は「上（過去方向）」にしかスクロール
  // しないため、先に一番下まで進めておかないと最新メッセージを
  // 取りこぼす。3通りの方法を併用するのは上方向スクロールと同じ理由
  for (let i = 0; i < 3; i++) {
    if (extractionCancelled) break;
    const visible = container.querySelectorAll(messageSelector);
    const bottomMessage = visible[visible.length - 1];
    if (bottomMessage && bottomMessage.scrollIntoView) {
      bottomMessage.scrollIntoView({ block: 'end', behavior: 'instant' });
    }
    container.scrollTop = container.scrollHeight;
    container.dispatchEvent(new Event('scroll', { bubbles: true }));
    container.dispatchEvent(
      new WheelEvent('wheel', { deltaY: 800, bubbles: true, cancelable: true })
    );
    const grew = await waitForContainerGrowth(container, GROWTH_TIMEOUT_MS, GROWTH_POLL_MS);
    if (!grew) break; // これ以上下に新しい内容が読み込まれなくなったら終了
  }

  for (let round = 0; round < MAX_ROUNDS; round++) {
    if (extractionCancelled) break;
    if (Date.now() - startedAt > MAX_TOTAL_MS) break;

    // 現在見えている分を収集
    for (const msg of collectVisible(container, messageSelector)) {
      addMessage(seen, msg);
    }

    onProgress && onProgress(seen.size, round);

    // 先頭（過去方向）へ複数の方法でスクロールを試みる。
    // container.scrollTop の直接代入だけでは、Teamsの仮想リストが
    // wheelイベント等でしか反応しない場合に何も起きないことがあるため、
    // 現在の最上部メッセージへのscrollIntoViewとwheelイベント発火も併用する
    const visibleNow = container.querySelectorAll(messageSelector);
    const topMessage = visibleNow[0];
    if (topMessage && topMessage.scrollIntoView) {
      topMessage.scrollIntoView({ block: 'start', behavior: 'instant' });
    }
    container.scrollTop = 0;
    container.dispatchEvent(new Event('scroll', { bubbles: true }));
    container.dispatchEvent(
      new WheelEvent('wheel', { deltaY: -800, bubbles: true, cancelable: true })
    );

    // 高さが変わるまで（＝過去メッセージの読み込みが反映されるまで）待つ。
    // Teams側のAPI応答が遅い場合もあるため、固定時間ではなくポーリングする。
    // ただしscrollHeightは無関係な再描画でも微妙に揺れることがあるため、
    // 終了判定そのものには使わず、あくまで「待つ時間の目安」としてのみ使う
    await waitForContainerGrowth(container, GROWTH_TIMEOUT_MS, GROWTH_POLL_MS);

    // もう一度収集（読み込み後に生成された要素を拾う）
    const sizeBefore = seen.size;
    for (const msg of collectVisible(container, messageSelector)) {
      addMessage(seen, msg);
    }

    // 終了判定は「実際に新しいメッセージを拾えたか」で行う。
    // scrollHeightは無関係なUIの再描画でも変化しうるし、逆に本当に
    // 読み込みが終わっていても微妙に揺れて終了判定に届かないことがある
    if (seen.size === sizeBefore) {
      idleRounds++;
    } else {
      idleRounds = 0;
    }

    if (idleRounds >= MAX_IDLE_ROUNDS) break;
  }

  return Array.from(seen.values());
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.type === 'CANCEL_EXTRACT') {
    extractionCancelled = true;
    sendResponse({ ok: true });
    return true;
  }

  if (msg.type === 'START_EXTRACT') {
    // 抽出には数十秒〜最大10分かかる。background(service worker)は1回の
    // 応答をそんなに長時間待ち続けられない（アイドル判定で再起動されうる）
    // ため、開始を受理した旨だけ即座に返し、実際の結果はPROGRESSと同じ
    // 「待たれない通知」としてEXTRACT_RESULTで別途送る
    sendResponse({ ok: true });

    (async () => {
      try {
        const pageTitleInfo = extractPageTitle();
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
        chrome.runtime.sendMessage({ type: 'PROGRESS_STAGE', stage: '画像を取得中...' });
        await embedImagesAsBase64(messages);

        // 診断情報: 画像の埋め込み成否と失敗理由、本文要素を特定でき
        // なかったメッセージ数（＝リンク・画像が欠落した可能性のある件数）
        // をポップアップのログに出せるよう集計して返す
        const stats = {
          imagesTotal: 0,
          imagesEmbedded: 0,
          imageFailures: [],
          bodyMissing: 0,
          fileCards: 0,
          bodyAnchors: 0,
          bodyAnchorsWithHref: 0,
        };
        for (const m of messages) {
          if (m.noBody) stats.bodyMissing++;
          stats.fileCards += (m.files || []).length;
          if (m.linkStats) {
            stats.bodyAnchors += m.linkStats.bodyAnchors;
            stats.bodyAnchorsWithHref += m.linkStats.bodyAnchorsWithHref;
          }
          for (const img of m.images || []) {
            stats.imagesTotal++;
            if (img.dataUri) stats.imagesEmbedded++;
            else if (stats.imageFailures.length < 10) {
              stats.imageFailures.push({
                src: String(img.src).slice(0, 120),
                error: img.error || '不明なエラー',
              });
            }
          }
        }

        chrome.runtime.sendMessage({
          type: 'EXTRACT_RESULT',
          ok: true,
          messages,
          stats,
          pageTitle: pageTitleInfo.title,
          titleSource: pageTitleInfo.source,
        });
      } catch (e) {
        chrome.runtime.sendMessage({
          type: 'EXTRACT_RESULT',
          ok: false,
          error: String(e && e.message ? e.message : e),
        });
      }
    })();

    return false; // sendResponseは既に同期的に呼び終えている
  }
});
