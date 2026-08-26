# IdGroupHorizontal

[![Direct](https://img.shields.io/badge/Direct%20Link-IdGroupHorizontal.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/group/IdGroupHorizontal.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdGroupHorizontal.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択したオブジェクトを縦位置の近さで「行」に分け、行ごとにグループ化します。

## 主な機能

- 中心 Y 座標の差が許容値以内のオブジェクトを同じ行とみなす
- 1 行に 2 つ以上あるときだけグループ化
- 作成したグループ数をダイアログで報告

## 使い方

1. グループ化したいオブジェクトを選択する
2. スクリプトを実行する

## 制限事項・メモ

- 許容値はスクリプト冒頭の `ROW_TOLERANCE`（現在のルーラー単位）で変更できます。
- より細かく指定したい場合は IdSmartGroup.jsx を使ってください。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/group/IdGroupHorizontal.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-11 |
| 最終更新 | 2026-04-17 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
