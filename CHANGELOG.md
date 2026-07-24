# Changelog

このプロジェクトの変更履歴。フォーマットは [Keep a Changelog](https://keepachangelog.com/ja/1.0.0/) に、
バージョニングは [Semantic Versioning](https://semver.org/lang/ja/) に準拠。

## [Unreleased]

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
