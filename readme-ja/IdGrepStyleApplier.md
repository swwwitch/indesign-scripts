# IdGrepStyleApplier

[![Direct](https://img.shields.io/badge/Direct%20Link-IdGrepStyleApplier.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdGrepStyleApplier.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdGrepStyleApplier.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

段落スタイルに正規表現スタイル（GREP スタイル）を適用・管理します。ルールと文字スタイルを選び、複数の段落スタイルへまとめて反映できます。

## 主な機能

- 登録済みルール（箇条書きラベル／言語設定／スマル／目次の数字／インライングラフィック）から選択
- 専用ダイアログでカスタムルールを追加
- 文字スタイルの選択と新規作成（ルールに応じた初期設定を自動適用）
- 適用先の段落スタイルを複数選択（Option/Alt クリックで全選択／全解除）
- 同じ段落スタイル内に同一の正規表現がある場合は上書き

## 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する
3. ルール・文字スタイル・適用先の段落スタイルを選んで［OK］

## 制限事項・メモ

- 必須項目（段落スタイル／文字スタイル）が未選択のあいだ OK ボタンは無効です。
- IdNestedStyleSetup.jsx は同一内容です。どちらか一方の利用を推奨します。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdGrepStyleApplier.jsx` |
| バージョン | v1.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-05-03 |
| 最終更新 | 2026-05-03 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
