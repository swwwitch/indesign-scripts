# IdSwitchToMasterOrDocument

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSwitchToMasterOrDocument.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/IdSwitchToMasterOrDocument.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSwitchToMasterOrDocument.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

アクティブページが親（マスター）ページかドキュメントページかを判定し、対応するもう一方へ切り替えます。

## 主な機能

- ドキュメントページ表示中なら適用中の親ページへジャンプ
- 親ページ表示中なら元のドキュメントページへ戻る
- 見開きでは、ページの左右に応じた親ページ側のページを選択

## 使い方

1. 切り替えたいページをアクティブにする
2. スクリプトを実行する

## 制限事項・メモ

- 戻り先のページ名はドキュメントの `label` プロパティに一時的に保存されます。
- 親ページが適用されていないページでは切り替えできません。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/IdSwitchToMasterOrDocument.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2025-07-02 |
| 最終更新 | 2025-07-02 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
