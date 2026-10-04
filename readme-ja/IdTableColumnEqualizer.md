# 表幅を親テキストフレームの幅にそろえる

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTableColumnEqualizer.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdTableColumnEqualizer.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableColumnEqualizer.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

カーソルのある表の幅を、親テキストフレームの幅にそろえます。

### 主な機能

- 親方向にたどってテキストフレームを探し、その幅を表に適用
- 列幅は InDesign の既定動作に従って比例配分される

### 使い方

1. 表内にカーソルを置く
2. スクリプトを実行する

### 制限事項・メモ

- 表がテキストフレーム内にない場合はメッセージを表示して終了します。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/table/IdTableColumnEqualizer.jsx` |
| バージョン | v1.0.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-17 |
| 最終更新 | 2026-09-30 |

### 更新履歴
- v1.0.2（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）

#### v1.0.1（2026-09-30）

- セル内の単語や段落を選んでいるときに、表が見つからず警告が出た問題を修正
- 表の選択とローカライズの処理を共通部品に統一

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
