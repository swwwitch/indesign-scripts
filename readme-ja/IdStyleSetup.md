# IdStyleSetup

[![Direct](https://img.shields.io/badge/Direct%20Link-IdStyleSetup.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdStyleSetup.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdStyleSetup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

段落スタイル・文字スタイルとそのグループ、継承関係、正規表現スタイルまでを一括で登録します。

## 主な機能

- スタイルとスタイルグループの作成、属性の適用、正規表現スタイルの設定、並び替えを 4 段階で実行
- basedOn による継承関係を常に設定し、共通の組版設定を基準スタイルへ集約
- 共通の GREP（lang-US / no-break / inline-graphic）を base-regex に集約し、子スタイルへ継承
- ul-li は自身に li-label を持つため、共通 3 つも直接設定して継承切れを回避
- 処理中はプログレスバーを表示

## 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する

## 制限事項・メモ

- 既定では既存の同名スタイルには触れません。`OVERWRITE_EXISTING_STYLES` を true にすると全属性を再適用し、GREP を付け直します。
- フォント・太さ・色などスクリプトが扱わない属性は変更しません。
- 全処理が 1 回の取り消しにまとまります。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdStyleSetup.jsx` |
| バージョン | v1.3.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-05-03 |
| 最終更新 | 2026-06-30 |
| 紹介記事 | https://note.com/dtp_tranist/n/nfe87ec253780 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
