# IdTableWidthColumnWidthManager

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTableWidthColumnWidthManager.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdTableWidthColumnWidthManager.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableWidthColumnWidthManager.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択位置から対象の表を特定し、表全体の幅と列の幅をプレビュー付きでまとめて調整します。

## 主な機能

- 表全体の幅は「変更しない／自動調整／親フレームいっぱいに／指定」
- 列の幅は「均等に／自動調整／自動調整を維持／最終列のみ調整／指定」
- 常時プレビューと、実行前の選択状態の復元
- 指定値は定規の単位で入力でき、↑↓ キーで増減（Shift: ±10、Option: ±0.1）

## 使い方

1. 表内にカーソルを置く
2. スクリプトを実行する
3. 表全体の幅と列の幅を指定して［OK］

## 制限事項・メモ

- 列の幅で「指定」を選んだ場合、表全体の幅より列の幅が優先され、表全体の幅は「列の幅 × 列数」で自動更新されます。
- 「自動調整を維持」は、自動調整で得た各列幅を基準に、表全体の幅の増減差分を列数で均等配分します。
- 「最終列のみ調整」は、最終列以外の現在の幅を維持したまま最終列だけを調整します。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/table/IdTableWidthColumnWidthManager.jsx` |
| バージョン | v1.1.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-18 |
| 最終更新 | 2026-05-05 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
