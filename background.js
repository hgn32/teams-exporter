// ============================================================
// Teams メッセージ抽出ツール - background service worker
//
// 画像のbase64化専用。content script内のfetchはページ側のCORS制約を
// 受けるため、asyncgw.teams.microsoft.com 等の別オリジンから配信される
// 画像は取得できない。service worker内のfetchはhost_permissionsを
// 与えたホストに対してCORSの制約を受けずに実行できるため、
// content scriptで取得に失敗した画像はここで代わりに取得する。
// 外部サーバーへの送信は一切行わない（GETで画像を読むだけ）。
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
