# 目印の文字列を削除して段落スタイルを適用（シンプル版）

[![Direct](https://img.shields.io/badge/Direct%20Link-IdRemoveMarkerApplyStyleSimple.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdRemoveMarkerApplyStyleSimple.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdRemoveMarkerApplyStyleSimple.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

指定した検索文字列（目印）を含む段落に段落スタイルを適用し、その目印を削除します。[IdRemoveMarkerApplyStyle](IdRemoveMarkerApplyStyle.md) から、検索文字列・段落スタイル・検索対象の 3 項目だけを残したシンプル版です。

### 主な機能

- 検索文字列（初期値 `###`）を含む段落に、選んだ段落スタイルを適用して目印を削除
- 目印に続く半角／全角スペース・タブも一緒に削除
- 同じ文字が続く並びの一部には一致しない（`##` は `### 見出し` に一致しません）
- 検索対象はストーリー／ドキュメント／すべてのドキュメント。テキストを選択していれば「ストーリー」、なければ「ドキュメント」で開きます
- 「対象箇所」に、現在の設定で見つかる件数を表示。検索文字列や検索対象を変えるたびに数え直します
- スタイルグループ内の段落スタイルも選択可能
- 一連の処理は 1 回の取り消し（Cmd+Z）で元に戻せます

### 使い方

1. 目印を含むテキストにカーソルを置く（ドキュメント全体を対象にする場合は開くだけで可）
2. スクリプトを実行する
3. 検索文字列・置換スタイル・検索対象を指定して［OK］

### 制限事項・メモ

- 検索文字列は行頭に固定しません。段落の途中にある目印も削除します。
- 検索は「半角と全角を区別」「ひらがなとカタカナを区別」をオンにして行います。前回の［検索と置換］の設定は引き継ぎません。
- 脚注・マスターページ・非表示レイヤー・ロックされたレイヤー・ロックされたストーリーは検索しません（スクリプト冒頭の設定で変更できます）。
- 「すべてのドキュメント」で、選んだ段落スタイルが無いドキュメントはスキップし、完了時に一覧で知らせます。
- 行頭の記号の自動判別、文字スタイルの適用、GREP での検索が必要な場合は [IdRemoveMarkerApplyStyle](IdRemoveMarkerApplyStyle.md) を使ってください。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdRemoveMarkerApplyStyleSimple.jsx` |
| バージョン | v1.1.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-09-09 |
| 最終更新 | 2026-09-30 |

### 更新履歴

#### v1.1.0（2026-09-30）

- 対象箇所が0件のときは［OK］を押せないように変更（スペースだけの検索文字列で本文のスペースを消してしまうおそれがあった）
- 検索文字列の末尾のスペース・タブを無視するように変更
- 検索まわりの処理を IdRemoveMarkerApplyStyle と共通化し、ローカライズ・余白を共通部品に統一

- v1.0（2026-09-10）初期バージョン

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
