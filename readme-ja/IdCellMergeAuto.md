# 内容が同じセルを自動で結合

[![Direct](https://img.shields.io/badge/Direct%20Link-IdCellMergeAuto.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdCellMergeAuto.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCellMergeAuto.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択した表について、内容が同じ隣接セルを自動で結合します。結合する方向と対象をダイアログで選べます。

## 主な機能

- 結合方向を「単一方向」または「両方向（優先付き）」から選択
- 方向は水平／垂直を指定（両方向の場合は先に処理する側になる）
- 対象を「表全体」または「選択セルのみ」から選択
- 比較時は前後の空白を除いた内容で判定
- 各パネルにツールチップで補足を表示

## 使い方

1. 表、または表の中のセルを選択する
2. スクリプトを実行する
3. 結合方法・方向・対象を指定して［OK］

## 制限事項・メモ

- 「選択セルのみ」は、選択されたセルを囲む矩形範囲が対象になります。
- すでに結合されていて行・列のまたぎ方が揃わないセル同士は、結合できないため飛ばします。
- 全処理が 1 回の取り消しにまとまります。


## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/table/IdCellMergeAuto.jsx` |
| バージョン | v1.0.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-17 |
| 最終更新 | 2026-08-27 |
| 紹介記事 | https://note.com/dtp_tranist/n/na84f68305844 |

## 更新履歴

### v1.0.1（2026-08-27）

- ［選択セルのみ］が常に「有効なセル選択が見つかりませんでした」となり、使えなかった問題を修正。選択セルの取得方法が誤っていました
- ［結合］［方向］［対象］の各パネルにツールチップを追加
- ファイル名を `MergeCell-Auto.jsx` から `IdCellMergeAuto.jsx` に変更
- 水平方向・垂直方向の処理を1つにまとめて整理（動作は変わりません）

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
