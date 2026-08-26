# 同一スタイル間の段落間隔をスタイル定義に設定

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSetSameParaStyleSpacing.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdSetSameParaStyleSpacing.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSetSameParaStyleSpacing.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

段落スタイルの「同一スタイル間の段落間隔」を、スタイル定義そのものに対して設定します。

## 主な機能

- グループ内の段落スタイルも再帰的に収集してドロップダウンに表示
- 間隔は「無視」「0」「数値指定」から選択
- 数値は環境設定の表示単位に連動し、↑↓ キーで増減（Shift＝10 の倍数、Option＝0.1 刻み）
- テキスト選択中なら、適用中の段落スタイルを初期選択にして現在値も反映

## 使い方

1. 対象のドキュメントを開く（設定したい段落にカーソルを置いておくと初期選択が便利）
2. スクリプトを実行する
3. 段落スタイルと間隔を指定して［OK］

## 制限事項・メモ

- 変更はスタイル定義に対して行われるため、そのスタイルを使うすべての段落に反映されます。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdSetSameParaStyleSpacing.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-06-30 |
| 最終更新 | 2026-06-30 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
