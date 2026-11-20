# 流し込んだタグ付きテキストの記述を書式に変換

[![Direct](https://img.shields.io/badge/Direct%20Link-IdRecomposeTaggedText.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdRecomposeTaggedText.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdRecomposeTaggedText.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

テキストフレームに文字のまま入っているタグ付きテキストの記述（`<ParaStyle:見出し>`・`<cSize:20>` など）を、タグ付きテキストとして読み込み直して書式に変換します。何も選択していないときは、ドキュメント内でタグの記述を含むストーリーをすべて変換します。

### 主な機能

- 文字列をタグ付きテキストのファイルに書き出し、InDesign の読み込み機能で同じ位置に読み直すので、段落スタイル・文字スタイルに限らずすべてのタグが効く
- 選択に応じて対象を決める
  - テキストフレーム（グループの中も含む）・カーソル：そのストーリー全体（表のセルならそのセル）
  - 文字の選択：その範囲だけ
  - 表・セルの選択：選んだセル
  - 何も選択していない：タグの記述を含むストーリーすべて
- タグの無い段落は、変換前の段落スタイルのまま残す
- 処理全体が 1 回の取り消しにまとまる

### 使い方

1. タグの記述を含むテキストフレーム・文字・セルを選択する（ドキュメント全体なら選択を解除する）
2. スクリプトを実行する

### 制限事項・メモ

- 表・アンカー付きオブジェクト・脚注・ページ番号などの特殊文字を含むテキストは、読み直すと失われるので変換しません。件数を表示します。
- タグの無い部分も読み直すため、変換前に手動で付けた書式（文字単位の上書き）は消えます。段落スタイルは残ります。
- 定義されていない文字スタイルを指定した `<CharStyle:名前>` は適用されません。段落スタイル・文字スタイルは先に作っておくか、`<DefineParaStyle:…>` などの定義をテキストに含めてください。
- 先頭の `<SJIS-MAC>` などのエンコーディング指定は取り除いて読み直します。
- 読み込みの設定（引用符の自動変換・テキストの書式を削除）は処理中だけ変更し、終わったら元に戻します。
- 全部変換できたときは何も表示しません。ロックされたストーリーなどで変換できなかったものは件数を表示します。
- 発想は id_recomposeTaggedText_a.jsx（z-、2012）です。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/text/IdRecomposeTaggedText.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-11-21 |
| 最終更新 | 2026-11-21 |

### 更新履歴

- v1.0.0（2026-11-21）初版

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
