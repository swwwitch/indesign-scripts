# 近接するオブジェクトを行・列単位でグループ化

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSmartGroup.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/group/IdSmartGroup.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSmartGroup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択したオブジェクトを水平方向（行）または垂直方向（列）の近さでまとめてグループ化します。ダイアログ表示中はグループ範囲を赤いプレビューで確認できます。

## 主な機能

- 方向（水平／垂直）をラジオボタンで切り替え
- 許容値をスライダーで 0〜50 の範囲で調整
- グループ化される範囲を非印刷レイヤー上に赤い矩形でリアルタイム表示
- ダイアログを閉じるとプレビューは自動的に削除

## 使い方

1. グループ化したいオブジェクトを選択する
2. スクリプトを実行する
3. 方向と許容値を調整し、プレビューを確認して［OK］

## 制限事項・メモ

- プレビューは "IdSmartGroup Preview" という非印刷レイヤーに描かれ、終了時に削除されます。
- グループ化には 1 つのまとまりに 2 つ以上のオブジェクトが必要です。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/group/IdSmartGroup.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-11 |
| 最終更新 | 2026-04-17 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
