# 段落境界線をパレットでプレビューしながら設定

[![Direct](https://img.shields.io/badge/Direct%20Link-IdParagraphRulesPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdParagraphRulesPalette.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdParagraphRulesPalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

InDesign の［段落境界線］ダイアログボックスを常駐パレットで再現し、選択したテキスト・テキストフレーム・表のセルの段落に前境界線・後境界線をプレビューしながら設定します。

### 主な機能

- 前境界線と後境界線を1つの画面に並べて設定（標準のダイアログボックスのように切り替える必要がない）
- 線幅・種類・カラー・濃淡・線オーバープリント、間隔カラー・間隔の濃淡・間隔オーバープリント、幅（段／テキスト）、オフセット、インデント、フレームに収める（前境界線のみ）
- 開いたときから変えた項目だけを書き込むので、ほかの設定や複数段落で異なる値はそのまま残る
- ［プレビュー］をオンにすると、設定をドキュメントに仮に適用して表示
- ［OK］で1回の取り消し（⌘Z）にまとめて適用
- カラーの横に色見本を表示。［(テキストカラー)］は段落の文字の色で表示
- 設定に応じて項目をディム表示（［境界線を挿入］がオフなら全体、実線なら間隔の項目、テキストカラー・なし・紙色なら濃淡とオーバープリント）
- インデントの項目名は横組みで「左／右」、縦組みで「上／下」
- 数値欄は ∧∨ ボタンと ↑↓ キーで増減（shift で10の倍数へ、option で0.1ずつ）。四則演算や異なる単位での入力にも対応（mm の欄に「1in」→「25.4 mm」）
- 長さの単位はドキュメントの単位設定（線・横・縦）に合わせる
- 日本語／英語 UI に対応

### 使い方

1. テキスト、テキストフレーム、または表のセルを選択する
2. スクリプトを実行する
3. パレットで前境界線・後境界線を設定する（［プレビュー］で仕上がりを確認）
4. ［OK］で適用する。［キャンセル］・Esc・クローズボックスで閉じると元に戻る

### 制限事項・メモ

- パレットに表示する値は、選択範囲の先頭の段落から読み込みます
- パレットを開いたまま再実行すると、前のパレットを閉じて（プレビューを戻して）から開き直します
- パレットを開いている間に対象のテキストが削除されると、プレビューや適用ができずに警告を表示します

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/text/IdParagraphRulesPalette.jsx` |
| バージョン | v1.3.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-10-19 |
| 最終更新 | 2026-10-19 |

### 更新履歴

#### v1.3.0（2026-10-19）

- 公開

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
