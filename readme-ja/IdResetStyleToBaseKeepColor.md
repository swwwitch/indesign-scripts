# 派生スタイルのオーバーライドをカラー以外削除

[![Direct](https://img.shields.io/badge/Direct%20Link-IdResetStyleToBaseKeepColor.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdResetStyleToBaseKeepColor.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdResetStyleToBaseKeepColor.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

ほかのスタイルを基準にした段落スタイル・文字スタイルから、カラー以外のオーバーライドを削除して基準スタイルの設定に戻します。文字カラー・段落背景色・下線や境界線のカラーなどは残します。

### 主な機能

- 段落スタイル・文字スタイルを対象にできる（スタイルグループの中も含む）
- 残すカラーを分類ごとに選べる：文字カラー（塗り・線・濃淡・線幅・オーバープリント）、段落背景色（オン／オフ・カラー・濃淡・オーバープリント）、下線・打ち消し線のカラー、段落境界線・段落囲みのカラー、その他のカラー（圏点・ルビなど）
- ［基本段落］を直接の基準にしたスタイルは、既定では処理しない（オプションで含められる）
- ダイアログボックスに、処理するスタイル数と項目数を表示
- 一度の取り消しで元に戻せる

### 使い方

1. ドキュメントを開く
2. スクリプトを実行し、対象と残すカラーを選んで［OK］をクリック

### 制限事項・メモ

- 段落スタイルは、基準スタイルと同じ値を入れてオーバーライドを外します（同じ値を入れると継承に戻ることを実測で確認）。文字スタイルは未設定に戻します
- ［次のスタイル］と書き出しタグは書式ではないので、そのまま残します
- 先頭文字スタイル・正規表現スタイル・行スタイルは対象外です
- 段落背景色のオフセットや幅など、カラー以外の寸法は基準に戻ります

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdResetStyleToBaseKeepColor.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-10-21 |
| 最終更新 | 2026-10-21 |

### 更新履歴

#### v1.0.0（2026-10-21）

- 初回リリース

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
