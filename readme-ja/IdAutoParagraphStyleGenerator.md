# IdAutoParagraphStyleGenerator

[![Direct](https://img.shields.io/badge/Direct%20Link-IdAutoParagraphStyleGenerator.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/IdAutoParagraphStyleGenerator.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAutoParagraphStyleGenerator.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

［基本段落］などスタイル未設定の段落を、フォント・文字サイズ・行送りの組み合わせごとにまとめ、段落スタイルを自動生成して適用します。

## 主な機能

- 参照する属性はフォント（ファミリー／スタイル）＋文字サイズ＋行送りのみ
- 内部の判定は pt に正規化するため、pt／Q 換算の誤差でグループが分かれにくい
- 環境設定［単位と増減値］のテキストサイズ単位に合わせてスタイル名の表記（pt / Q）が変わる
- 段落内で書式が混在している段落はスキップし、件数を報告

## 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する
3. 結果ダイアログで作成数と適用数を確認する

## 制限事項・メモ

- 生成されるスタイル名は `AutoStyle_1_10pt` のような形式です。
- 同名スタイルがある場合は連番を付けて重複を避けます。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/IdAutoParagraphStyleGenerator.jsx` |
| バージョン | v3.4 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-02-13 |
| 最終更新 | 2026-03-14 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
