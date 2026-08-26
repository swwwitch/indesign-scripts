# 表の結合セルを解除

[![Direct](https://img.shields.io/badge/Direct%20Link-IdCellUnmerge.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdCellUnmerge.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCellUnmerge.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

表の結合セルを解除します。解除後のセルへ元のテキストを複製するかどうかと、対象範囲をダイアログで選べます。

## 主な機能

- 「デフォルト（分配なし）」と「テキストを分配」を切り替え
- 対象を「表全体」「選択したセルのみ」から選択
- 重複選択されたセルは 1 回だけ処理

## 使い方

1. 表のセルを選択する
2. スクリプトを実行する
3. 分配の有無と対象範囲を指定して［OK］

## 制限事項・メモ

- 「テキストを分配」を選ぶと、解除後のすべてのセルに元のテキストが複製されます。
- 1 セルの処理に失敗しても全体は止まりません。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/table/IdCellUnmerge.jsx` |
| バージョン | v1.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-17 |
| 最終更新 | 2026-04-17 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
