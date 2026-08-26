# 段落背景色の高さに合わせた境界線スペーサーを設定

[![Direct](https://img.shields.io/badge/Direct%20Link-IdParagraphShadingMatchRule.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdParagraphShadingMatchRule.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdParagraphShadingMatchRule.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択テキストの各段落に、段落背景色の高さに合わせた不可視の段落境界線（前境界線）をスペーサーとして設定します。

## 主な機能

- 線幅は「段落背景色の上オフセット＋文字サイズ」で算出
- 罫線カラーは「なし」に設定するため画面・出力には現れない
- テキストフレーム内に罫線を保持する設定を自動で有効化

## 使い方

1. 対象のテキストを選択する
2. スクリプトを実行する

## 制限事項・メモ

- 見た目の線を描くためのスクリプトではありません。太い透明の段落境界線をスペーサーとして使い、背景領域の高さを疑似的に再現する設計です。
- 段落背景色の上マージンが未設定の場合は 0 として扱います。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdParagraphShadingMatchRule.jsx` |
| バージョン | v1.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-12 |
| 最終更新 | 2026-04-12 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
