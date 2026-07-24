# Changelog

このプロジェクトの変更履歴。フォーマットは [Keep a Changelog](https://keepachangelog.com/ja/1.0.0/) に、
バージョニングは [Semantic Versioning](https://semver.org/lang/ja/) に準拠。

## [Unreleased]

## [0.1.4] - 2026-07-24

### Fixed

- popupを閉じた状態だと抽出完了後のダウンロードが無言で失敗することがある不具合を修正。
  background service worker内でダウンロード用URLをBlob/URL.createObjectURLで
  作っていたが、service worker環境ではこれが失敗しうるため、DOM APIに依存しない
  data:URIを直接組み立てる方式に変更。失敗時は必ずログに理由を出すようにした。
- background(service worker)がcontent scriptからの応答を長時間（最大10分）待ち
  続ける設計だったため、抽出の途中でservice workerが再起動されると結果を
  取りこぼす場合があった。content script側は開始受理のみ即座に返し、
  抽出結果は完了時に別途通知する方式に変更。あわせて実行中の状態を
  chrome.storage.sessionにも保存し、service workerが再起動されても
  途中経過・結果を正しく引き継げるようにした。
- 抽出対象のタブが閉じられた場合に、状態が「実行中」のまま固まらないようにした。

## [0.1.3] - 2026-07-24

### Added

- Teams以外のタブで拡張機能を開いた場合に「Teams Web版を開く」ボタンを表示するようにした。

### Changed

- 抽出処理（content scriptとの通信・完了後のHTML生成・ダウンロード）をpopup.jsから
  background service worker側に移動。popupを閉じても抽出処理・ダウンロードが
  継続するようになった（従来はpopupを閉じるとJS実行が止まり、ダウンロードまで
  到達できなかった）。
- popupを開いたときに、進行中または直近の抽出処理のログ履歴を表示するようにした。
  処理に時間がかかりpopupを閉じてしまっても、再度開けば経過を確認できる。
