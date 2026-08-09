# SetLastColumnToFrameWidth

カーソルのある表の最終列幅を調整し、表全体の幅を親テキストフレームの幅に合わせます。

## 主な機能

- 最終列以外の列幅を保ったまま、最終列だけを伸縮
- 表を含むテキストフレームの幅を基準にする

## 使い方

1. 表内にカーソルを置く
2. スクリプトを実行する

## 制限事項・メモ

- 表がテキストフレーム内にある必要があります。
- 最終列以外の合計が親フレーム幅を超える場合、最終列の幅は負の値になり得ます。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/table/SetLastColumnToFrameWidth.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-17 |
| 最終更新 | 2026-04-17 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
