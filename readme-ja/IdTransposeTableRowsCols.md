# 表の行と列を入れ替え

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTransposeTableRowsCols.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdTransposeTableRowsCols.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTransposeTableRowsCols.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択した表の行と列を入れ替えます。ヘッダー行の扱いとセル結合の処理方法をダイアログで選べます。

## 主な機能

- テキスト内容に加えて、文字サイズ・フォント・文字色・セルの塗り色・ティントも入れ替え
- セル結合がある場合は「しない（終了）」「転置前に解除」を選択
- ヘッダー行を転置対象にするかをチェックボックスで指定
- ヘッダー行・フッター行の設定を可能な範囲で復元

## 使い方

1. 表、セル、または表内のテキストを選択する
2. スクリプトを実行する
3. ヘッダー行とセル結合の扱いを指定して［OK］

## 制限事項・メモ

- ヘッダー行もセル結合もない場合はダイアログを省略してそのまま転置します。
- 転置の途中で正方形になるよう行または列を一時的に追加し、処理後に削除します。

## オリジナル

Table Transpose（堅牢性を高めるために改変）

Original: Table Transpose v1.0 by Iain Anderson

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/table/IdTransposeTableRowsCols.jsx` |
| バージョン | v1.0.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2025-11-25 |
| 最終更新 | 2026-09-07 |
| 紹介記事 | https://note.com/dtp_tranist/n/nc6dbdb3af6a1 |

## 更新履歴

### v1.0.1（2026-09-07）

- 紹介記事へのリンクを追加
- ハウスルールに沿ってヘッダー・命名・JSDocを整理（動作は変わりません）
- セルの書式を入れ替える処理を共通化し、転置処理を「正方形に揃える」「三角を入れ替える」「余りを取り除く」に分割

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
