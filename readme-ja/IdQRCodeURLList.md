# ドキュメント内のQRコードのURLを一覧

[![Direct](https://img.shields.io/badge/Direct%20Link-IdQRCodeURLList.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdQRCodeURLList.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdQRCodeURLList.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

アクティブなドキュメントにある QR コード（［オブジェクト］→［QR コードを生成］で作ったもの）の URL を、ページ順に一覧します。

QR コードは見た目からは URL が分からず、［QR コードを編集］を1つずつ開いて確かめるしかありません。このスクリプトなら、ドキュメント内のすべての QR コードの URL を1回で確認でき、入稿前のリンク先のチェックやチェックリスト作りに使えます。

### 主な機能

- 「ページ｜種類｜URL」の3列のリストで表示
- ドキュメントページ・親ページ・ペーストボード・グループの中・アンカー付きオブジェクト、ロック・非表示のレイヤーにある QR コードも対象
- 種類が「Web ハイパーリンク」の QR コードに加え、中身が URL（`http://` `https://` `ftp://` で始まる）の「プレーンテキスト」の QR コードも一覧に入れる
- 並びはドキュメントページ → 親ページ → ペーストボードの順。同じページの中では上から、同じ高さなら左から
- ［すべてコピー］で「ページ<タブ>URL」の行をクリップボードにコピー（表計算アプリにそのまま貼り付けられる）
- ［選択］またはダブルクリックで、その QR コードのページを表示してフレームを選択
- ダイアログのタイトルに件数を表示

### 使い方

1. ドキュメントを開いてスクリプトを実行する
2. 一覧で URL を確認する
3. 必要なら［すべてコピー］でクリップボードにコピーする
4. 場所を確かめたい QR コードを選んで［選択］をクリックする（ダブルクリックでも可）。ダイアログボックスが閉じ、そのページが表示されてフレームが選択される

Enter／Return で［選択］、Esc で［閉じる］と同じ動作になります。

### 制限事項・メモ

- QR コードの中身はスクリプトから読めるプロパティに無いため、QR コードのフレームを一時的にスニペット（.idms）に書き出して読み取ります。書き出したファイルはすぐに削除します。QR コード1つにつき 0.01 秒ほどかかります。
- ドキュメントには変更を加えません（［選択］で選択とページの表示が変わるだけです）。
- オーバーセットのテキストに入った QR コードは、スニペットへの書き出しで InDesign が落ちるおそれがあるため読み取らず、URL の欄に「（オーバーセットのため読み取れません）」と表示します（未検証）。
- 電子メール・テキストメッセージ・名刺の QR コードは一覧に入れません。
- 配置したリンクの EPS ファイルは対象外です（QR コードは埋め込みの EPS として入っているため）。
- ロック・非表示のレイヤーにある QR コードは一覧には出ますが、［選択］で選択できないことがあります。
- 一覧は実行した時点のものです。QR コードを編集したら、スクリプトを実行し直してください。
- クリップボードへのコピーは、Mac では AppleScript、Windows では VBScript を経由します。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/document/IdQRCodeURLList.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-11-05 |
| 最終更新 | 2026-11-05 |
| 紹介記事 | https://note.com/dtp_tranist/n/ne2cdd5f60080 |
| 動作確認 | InDesign 2026（v21.6）／ macOS |

### 更新履歴

- v1.0.0（2026-11-05）：初版

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
