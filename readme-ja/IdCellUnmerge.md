# 表の結合セルを解除

[![Direct](https://img.shields.io/badge/Direct%20Link-IdCellUnmerge.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdCellUnmerge.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCellUnmerge.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

表の結合セルを解除します。解除後のセルへ元のテキストを分配するかどうかと、対象範囲をダイアログで選べます。

## 主な機能

- 「デフォルト（分配なし）」と「テキストを分配」を切り替え
- 対象を「表全体」「選択したセルのみ」から選択
- 分配先は、その結合セルの解除で生じたセルだけに限定
- 重複選択されたセルは 1 回だけ処理

## 使い方

1. 表のセルを選択する
2. スクリプトを実行する
3. 分配の有無と対象範囲を指定して［OK］

## 制限事項・メモ

- 「テキストを分配」を選ぶと、その結合セルの解除で生じたすべてのセルに同じテキストが入ります。隣接する別のセルには影響しません。
- 分配されるのは、結合セルが表示していたテキストです。書式は引き継がれず、プレーンテキストとして入ります。
- 1 セルの処理に失敗しても全体は止まりません。
- 全処理が 1 回の取り消しにまとまります。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/table/IdCellUnmerge.jsx` |
| バージョン | v1.0.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-17 |
| 最終更新 | 2026-08-27 |

## 更新履歴

### v1.0.1（2026-08-27）

- ［選択したセルのみ］で結合が解除されなかったり、別のセルのテキストが混ざって入る問題を修正。結合セルの `contents` が構成セルごとのテキストの配列で返るのを、文字列として扱っていたのが原因です
- 「テキストを分配」の対象を、その結合セルの解除で生じたセルだけに限定
- ［表全体］で結合セルを取りこぼすことがあった問題を修正。解除するたびにセル数が変わり、後続セルの参照がずれていました
- 全体を即時関数（IIFE）で囲み、グローバル変数の流出をなくした

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
