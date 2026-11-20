# 選択したテキストをタグ付きテキストとして編集

[![Direct](https://img.shields.io/badge/Direct%20Link-IdEditAsTaggedText.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdEditAsTaggedText.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdEditAsTaggedText.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択したテキストをタグ付きテキストに書き出してダイアログボックスに表示します。タグや文字を編集して［OK］を押すと、同じ位置に読み込み直して書式に反映します。

### 主な機能

- 選択に応じて編集する範囲を決める
  - 文字の選択：その範囲
  - カーソル・テキストフレーム：ストーリー全体（表のセルならそのセル）
- 日本語などは `<0x6587>` ではなく文字のまま表示する（スペースの種類や制御文字は `<0x2002>` などのまま）
- 1行が1段落。強制改行は `<0x000A>`
- 段落の先頭の `<ParaStyle:…>` を消しても、元の段落スタイルのまま読み込む
- ［リセット］で編集前の状態に戻せる
- 処理全体が 1 回の取り消しにまとまる

### 使い方

1. 編集したい文字を選択する（またはテキストフレームを選択するか、カーソルを置く）
2. スクリプトを実行する
3. ダイアログボックスでタグ付きテキストを編集し、［OK］を押す

### 制限事項・メモ

- 表・アンカー付きオブジェクト・脚注・ページ番号などの特殊文字を含むテキストは編集できません（読み込み直すと失われるため）。
- 文字の `<` `>` `\` は、タグ付きテキストの書き方どおり `\<` `\>` `\\` と書きます。
- 定義されていない段落スタイル・文字スタイルを指定したタグは適用されません。
- タグは詳細形式（`<cSize:>` など）で表示します。略式（`<cs:>` など）にするにはスクリプト冒頭の `TAG_FORM` を `TagTextForm.ABBREVIATED` に変えます。
- 何も変えずに［OK］を押したときは何もしません。
- 書き出し・読み込みの設定（文字セット・タグの形式・引用符の自動変換など）は処理中だけ変更し、終わったら元に戻します。
- 読み込みに失敗したときは、編集前のテキストを読み込み直して元に戻します。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/text/IdEditAsTaggedText.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-11-21 |
| 最終更新 | 2026-11-21 |

### 更新履歴

- v1.0.0（2026-11-21）初版

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
