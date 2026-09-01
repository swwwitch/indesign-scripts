# 段落スタイル・文字スタイルを一括登録

[![Direct](https://img.shields.io/badge/Direct%20Link-IdStyleSetup.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdStyleSetup.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdStyleSetup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

段落スタイル・文字スタイルとそのグループ、継承関係、正規表現スタイルまでを一括で登録します。

## 主な機能

- スタイルとスタイルグループの作成、属性の適用、正規表現スタイルの設定、並び替えを 4 段階で実行
- basedOn による継承関係を先に確定させてから属性を適用し、共通の組版設定は基準スタイルへ集約
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
- カーニング方式は UI 言語で表示名が変わるため、候補を順に試します。どれも該当しない環境では設定を据え置きます。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdStyleSetup.jsx` |
| バージョン | v1.3.2 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-05-03 |
| 最終更新 | 2026-09-01 |
| 紹介記事 | https://note.com/dtp_tranist/n/nfe87ec253780 |

## 更新履歴

### v1.3.2（2026-09-01）

- `p` の分離禁止オプションを OFF にする設定などが効かないことがある問題を修正。属性を設定してから basedOn を張っていたため、親（body-text）側の設定を継承し直して打ち消される可能性がありました。継承関係を先に確定させてから属性を適用するように変更
- `basestyle` グループに `body-text` がないと、`heading` と h1〜h6 の継承関係まで設定されなくなる問題を修正
- 日本語版以外の InDesign でカーニング方式の設定に失敗し、スクリプト全体が取り消される問題を修正。表示名の候補（`和文等幅` / `Japanese Mojikumi` など）を順に試すように変更
- 未使用の UI ヘルパーを削除し、JSDoc の戻り値型とコメントの記述を実装に合わせて整理

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
