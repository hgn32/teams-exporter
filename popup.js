const statusEl = document.getElementById('status');
const actionBtn = document.getElementById('actionBtn');
const openTeamsBtn = document.getElementById('openTeamsBtn');

const TEAMS_URL_RE = /^https:\/\/teams\.microsoft\.com\//;
const TEAMS_WEB_URL = 'https://teams.microsoft.com/v2/';

let running = false;
let cancelling = false;

function render(log) {
  statusEl.textContent = log.join('\n');
  statusEl.scrollTop = statusEl.scrollHeight;
}

// 実際の抽出処理はbackground service worker側で行っている（popupを
// 閉じても継続させるため）。popupはbackgroundが持つログ・状態をただ
// 表示するだけのビューとして振る舞う
function sendToBackground(message) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage(message, (response) => {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, error: chrome.runtime.lastError.message });
        return;
      }
      resolve(response);
    });
  });
}

function updateButtonLabel() {
  if (cancelling) {
    actionBtn.disabled = true;
    actionBtn.textContent = '中止中...';
  } else if (running) {
    actionBtn.disabled = false;
    actionBtn.textContent = '抽出を中止';
  } else {
    actionBtn.disabled = false;
    actionBtn.textContent = '自動スクロール抽出を開始';
  }
}

chrome.runtime.onMessage.addListener((msg) => {
  if (msg.type === 'BG_LOG') {
    render(msg.log);
  } else if (msg.type === 'BG_STATUS') {
    running = msg.running;
    cancelling = msg.cancelling;
    updateButtonLabel();
  }
});

async function getActiveTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab;
}

openTeamsBtn.addEventListener('click', () => {
  chrome.tabs.create({ url: TEAMS_WEB_URL });
  window.close();
});

actionBtn.addEventListener('click', async () => {
  if (running) {
    actionBtn.disabled = true;
    actionBtn.textContent = '中止中...';
    await sendToBackground({ type: 'CANCEL_EXTRACT_REQUEST' });
    return;
  }

  const tab = await getActiveTab();
  if (!tab || !tab.url || !TEAMS_URL_RE.test(tab.url)) {
    render(['Teams のタブを開いた状態で実行してください。']);
    return;
  }

  const res = await sendToBackground({ type: 'START_EXTRACT_REQUEST', tabId: tab.id, tabUrl: tab.url });
  if (!res || !res.ok) {
    render(['エラー: ' + (res && res.error ? res.error : '開始できませんでした。')]);
    return;
  }
  running = true;
  updateButtonLabel();
});

async function init() {
  const tab = await getActiveTab();
  const isTeamsTab = !!(tab && tab.url && TEAMS_URL_RE.test(tab.url));

  const state = await sendToBackground({ type: 'GET_STATE' });
  running = !!(state && state.running);
  cancelling = !!(state && state.cancelling);
  updateButtonLabel();

  // どの版が動いているか一目で分かるよう、待機中は常にバージョンを表示する
  // （古い版のまま「直っていない」となる事故を防ぐ）
  if (state && state.log && state.log.length) {
    render(state.log);
  } else {
    render([`待機中... (v${chrome.runtime.getManifest().version})`]);
  }

  // Teams以外のタブでは抽出を始めようがないので、代わりにTeams Web版を
  // 開くボタンを出す。ただし他タブで抽出処理が進行中の場合は、中止操作を
  // できるようにするため通常の（中止）ボタンを優先して表示する
  if (!isTeamsTab && !running) {
    actionBtn.style.display = 'none';
    openTeamsBtn.style.display = 'block';
  } else {
    actionBtn.style.display = 'block';
    openTeamsBtn.style.display = 'none';
  }
}

init();
