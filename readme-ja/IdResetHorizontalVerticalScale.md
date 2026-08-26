# 文字の水平・垂直比率を 100% に戻す

[![Direct](https://img.shields.io/badge/Direct%20Link-IdResetHorizontalVerticalScale.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdResetHorizontalVerticalScale.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdResetHorizontalVerticalScale.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

ドキュメント内すべてのストーリー（表セル・入れ子の表を含む）を走査し、文字の水平・垂直比率を 100% に戻します。

## 主な機能

- 水平比率・垂直比率のいずれかが 100% でない文字範囲をリセット
- 表セル内の文字も、入れ子の表を含めて再帰的に処理
- 変更した文字範囲の件数を最後に表示

## 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する

## 制限事項・メモ

- 処理全体が 1 回の取り消しにまとまります。
- 対象はアクティブドキュメント内のすべてのストーリーです。範囲は指定できません。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/text/IdResetHorizontalVerticalScale.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-07-19 |
| 最終更新 | 2026-07-19 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
