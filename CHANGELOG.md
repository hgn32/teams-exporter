# Changelog

このプロジェクトの変更履歴。フォーマットは [Keep a Changelog](https://keepachangelog.com/ja/1.0.0/) に、
バージョニングは [Semantic Versioning](https://semver.org/lang/ja/) に準拠。

## [Unreleased]

### Changed

- リポジトリのファイル整理（拡張機能の動作に変更なし）:
  ストア掲載用の素材（STORE_LISTING.md・スクリーンショット・プロモ画像）を
  `store-assets/` フォルダにまとめ、`icons/` は実行に必要なアイコンのみにした。
- STORE_LISTING.md の見出しをEdgeアドオン向けからChromeウェブストア向けに修正。
- READMEの誤記修正（フォルダ名 `teams-extractor` → `teams-exporter`、
  完了ログの文言例を実際の出力形式に一致させた）。

## [0.1.5] - 2026-07-24

### Fixed

- アニメーション絵文字などで、テーマ・モーション設定違いの非表示バリアントまで
  画像として拾ってしまい、同じ絵文字が何十個も並んで出力される不具合を修正。
  実際に画面上に表示されている要素だけを対象にするようにした。

### Changed

- 添付ファイルのチップアイコンを、拡張子ごとの絵文字（📊/📄/📝等）出し分けから
  単一の📎アイコンに簡略化。

## [0.1.4] - 2026-07-24

### Fixed

- popupを閉じた状態だと抽出完了後のダウンロードが無言で失敗することがある不具合を修正。
  background service worker内でダウンロード用URLをBlob/URL.createObjectURLで
  作っていたが、service worker環境ではこれが失敗しうるため、DOM APIに依存しない
  data:URIを直接組み立てる方式に変更。失敗時は必ずログに理由を出すようにした。
- background(service worker)がcontent scriptからの応答を長時間（最大10分）待ち
  続ける設計だったため、抽出の途中でservice workerが再起動されると結果を
  取りこぼす場合があった。content script側は開始受理のみ即座に返し、
  抽出結果は完了時に別途通知する方式に変更（権限追加を避けるため、状態の
  永続化はせず、数秒おきに届くPROGRESS通知でservice workerの生存を維持する
  設計とした）。
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
