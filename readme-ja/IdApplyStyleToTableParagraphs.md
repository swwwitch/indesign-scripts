# 表だけの段落に段落スタイルを適用

[![Direct](https://img.shields.io/badge/Direct%20Link-IdApplyStyleToTableParagraphs.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdApplyStyleToTableParagraphs.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdApplyStyleToTableParagraphs.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択範囲・ストーリー・ドキュメントの中で、表だけが入っている段落を探し、指定した段落スタイルを適用します。

### 主な機能

- 対象を選択範囲・ストーリー・ドキュメントから選ぶ（セルの中の表も対象）
- 適用する段落スタイルをメニューから選ぶ（初期値は「p.table」）
- 段落の中身が表だけ（末尾の改行は除く）のときに適用する。表が2つ以上並んでいる段落も対象
- 適用した段落の手動の書式（オーバーライド）は消去する
- 1回の取り消しで元に戻せる

### 使い方

1. 対象にするテキスト・テキストフレーム・表を選択する（ドキュメント全体が対象なら不要）
2. スクリプトを実行する
3. ［対象］と［適用する段落スタイル］を指定して［OK］をクリックする
4. 適用した段落の数が表示される

### 制限事項・メモ

- ［選択範囲］は、テキスト・テキストフレーム・表・セルを選択しているときに選べます。カーソルだけのときは［ストーリー］が初期値です。
- ［ストーリー］は、テキストまたは表を選択しているか、カーソルがあるときに選べます。
- 選択した文字の範囲に表の目印の文字が入っていれば対象です。セルの中の文字だけを選択しているときは、その表自体は対象になりません。
- 「p.table」が無いドキュメントでは、段落スタイルを選ぶまで［OK］を押せません。初期値はスクリプト冒頭の `DEFAULT_PARAGRAPH_STYLE_NAME` で変更できます。
- 表と同じ段落に文字・空白・アンカー付きオブジェクト・分割文字（改ページなど）があるときは対象外です。
- 強制改行で区切っただけの行は、段落全体に別の文字があるので対象外です。
- ロックされたストーリーなどで適用できなかった段落は、数を表示して飛ばします。ロックしたレイヤーのテキストには適用されます。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdApplyStyleToTableParagraphs.jsx` |
| バージョン | v1.1.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-11-06 |
| 最終更新 | 2026-11-06 |

### 更新履歴

- v1.1.0（2026-11-06）ダイアログボックスを追加し、対象（選択範囲・ストーリー・ドキュメント）と適用する段落スタイルを選べるようにした
- v1.0.0（2026-11-06）初版

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
