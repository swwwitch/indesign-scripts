# 表の罫線をプレビュー付きで描画・消去

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSmartBorderBuilder.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdSmartBorderBuilder.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSmartBorderBuilder.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択した表セルに対して、モード・線幅・カラー・濃淡を指定しながら罫線をプレビュー付きで描画・消去します。

## 主な機能

- モードは「すべて／境界線のみ／内部のみ／水平線のみ／垂直線のみ／最下辺のみ／最右辺のみ／見出し行／見出し列／左右の境界線を消去／すべて消去」
- モード切り替えのショートカットキー（A/E/I/H/V/B/U/L/R/C）に対応
- 線幅はドキュメントの線幅単位に追従し、プリセット選択と ↑↓ キーでの増減が可能
- スウォッチによるカラー指定と、0〜100 の濃淡（Tint）調整
- ［標準モード］／［プレビュー］トグルボタンで画面モードを切り替え
- 前回 OK で確定した設定を記憶し、次回起動時に復元

## 使い方

1. 罫線を設定したい表セルを選択する
2. スクリプトを実行する
3. モード・線幅・カラー・濃淡を指定して［OK］

## 制限事項・メモ

- 表の一部を選択した場合も、選択範囲の矩形からセルを再構築して処理します。結合セルは各座標を覆っているセルを探索して再構築します。
- 「描画前に消去」を OFF にすると既存の罫線を残したまま上書きできます。
- ダイアログ終了後は実行前の選択状態を復元します。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/table/IdSmartBorderBuilder.jsx` |
| バージョン | v1.6.7 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-11 |
| 最終更新 | 2026-04-17 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
