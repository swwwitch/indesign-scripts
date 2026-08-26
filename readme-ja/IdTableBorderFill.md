# IdTableBorderFill

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTableBorderFill.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/IdTableBorderFill.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableBorderFill.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択した表セルの罫線を、適用範囲・線幅・カラー・濃淡を指定しながらプレビュー付きで調整します。

## 主な機能

- 適用範囲は「すべて／境界線のみ／内部のみ／水平線のみ／垂直線のみ／上下端／左右端／左右の罫線を消去／すべて消去」
- 適用範囲のショートカットキー（A/E/I/H/V/U/L/R/C）に対応
- 線幅はドキュメントの線幅単位に追従し、プリセット選択と ↑↓ キーでの増減が可能
- スウォッチによるカラー指定と、0〜100% の濃淡（Tint）調整
- ［プレビュー］トグルボタンで画面モードを切り替え

## 使い方

1. 罫線を調整したい表セルを選択する
2. スクリプトを実行する
3. 適用範囲・線幅・カラー・濃淡を指定して［OK］

## 制限事項・メモ

- 隣接する選択セルの向かい合う辺もあわせて処理するため、二重線になりません。
- 「適用前に消去」を OFF にすると既存の罫線を残したまま上書きできます。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/IdTableBorderFill.jsx` |
| バージョン | v1.5.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-11 |
| 最終更新 | 2026-04-13 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
