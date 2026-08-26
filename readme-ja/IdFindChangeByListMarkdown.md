# IdFindChangeByListMarkdown

[![Direct](https://img.shields.io/badge/Direct%20Link-IdFindChangeByListMarkdown.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/IdFindChangeByListMarkdown.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdFindChangeByListMarkdown.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

Markdown 記法をまとめて検索・置換し、対応する段落スタイルと文字スタイルを適用します。InDesign 付属の FindChangeByList.jsx をベースにしています。

## 主な機能

- 見出し・箇条書き・番号リスト・引用・コードブロック・表組み・画像などに対応
- HTML / MS Word のスタイル名プリセットを切り替え可能
- スコープをドキュメント／ストーリー／選択範囲から選択
- 行頭・行末スペース、連続する空行、HTML コメント、水平線のクリーンアップ

## 使い方

1. 変換したい Markdown テキストを含むドキュメントを開く
2. スクリプトを実行する
3. スコープと形式、各項目のスタイルを指定して［実行］

## 制限事項・メモ

- 変換順序は依存関係を考慮して固定されています（コードブロック→番号リスト→見出し…）。
- すべての置換は 1 回の取り消しで元に戻せます。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/IdFindChangeByListMarkdown.jsx` |
| バージョン | v1.0.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-03-17 |
| 最終更新 | 2026-03-17 |
| 紹介記事 | https://note.com/dtp_tranist/n/n8c0211d92c96 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
