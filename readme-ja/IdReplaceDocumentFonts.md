# ドキュメントで使用中のフォントを一括置換

[![Direct](https://img.shields.io/badge/Direct%20Link-IdReplaceDocumentFonts.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/font/IdReplaceDocumentFonts.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdReplaceDocumentFonts.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

ドキュメントで使用中のフォントをファミリー／スタイル単位で一覧し、選んだフォントを別のフォントへまとめて置き換えます。

## 主な機能

- 使用中のフォントを 2 つのリスト（置換元・置換先）に並べ、置換元は複数選択が可能
- ファミリー名の行を選ぶと、そのファミリーのスタイルがすべて選ばれる
- ［PostScript名で表示］で、ファミリー＋スタイルと PostScript 名の表示を切り替え
- ［段落・文字スタイルも更新］で、テキストだけでなく段落スタイル・文字スタイルのフォントも置き換え
- 各行に使用箇所の件数を表示。環境にないフォントには「未インストール」と表示
- ［全置換］で、使用中のすべてのフォントを 1 つのフォントにそろえる
- 親ページ、非表示レイヤー、脚注、表のセルも対象（InDesign の検索エンジンを使用）

## 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する
3. 左のリストで置換元フォント、右のリストで置換先フォントを選び、［フォントを置換］
4. 続けて別のフォントを置き換える場合は、そのままリストを選び直す（一覧は置換のたびに更新されます）

［全置換］は、置換先を選んでいればそのフォントに、選んでいなければ置換元の 1 つ目のフォントに、使用中のすべてのフォントをそろえます。

## 制限事項・メモ

- 置換先に選べるのは、ドキュメントで使用中のフォントだけです。
- ロックされたレイヤー・ストーリーは置換の対象外です（件数には含まれます）。
- 取り消しは置換元フォント 1 つにつき 1 ステップになります。
- ［段落・文字スタイルも更新］をオフのまま置換すると、テキストにはオーバーライドがかかり、段落スタイルを再適用すると元のフォントに戻ります。
- 合成フォントの構成メンバーは置換できません。
- 使用箇所の件数はフォントごとに検索して求めるため、フォント数・ページ数の多いドキュメントでは一覧の作成に時間がかかります。スクリプト冒頭の `SHOW_USAGE_COUNT` を `false` にすると件数の表示を省けます。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/font/IdReplaceDocumentFonts.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-09-20 |
| 最終更新 | 2026-09-20 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
