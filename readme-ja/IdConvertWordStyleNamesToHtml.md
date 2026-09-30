# Word の段落スタイル名を HTML 要素名に変換

[![Direct](https://img.shields.io/badge/Direct%20Link-IdConvertWordStyleNamesToHtml.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdConvertWordStyleNamesToHtml.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdConvertWordStyleNamesToHtml.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

MS Word から取り込んだ段落スタイル名（Heading 1 / Normal / Quote など）を、対応する HTML 要素名（h1 / p / blockquote など）へ一括でリネームします。ダイアログは出ず、実行するとすぐに処理して結果を表示します。

### 主な機能

- `Heading 1`〜`Heading 6` を `h1`〜`h6` へ（`Heading1` のようにスペースが無くても可）
- そのほかの対応は下表のとおり
- 大文字小文字・全角スペース・連続する空白の違いは無視して照合
- スタイルグループの中も再帰的にたどって処理
- リネーム先の名前の段落スタイルがすでにある場合は、そのスタイルに置き換えて統合（元のスタイルを使っていた段落は既存のスタイルに切り替わり、元のスタイルは削除）
- 実行後に、変換・統合・スキップ・エラーの件数と明細を表示
- 一連の処理は 1 回の取り消し（Cmd+Z）で元に戻せます

| Word のスタイル名 | 変換後 |
| --- | --- |
| Normal / Default Paragraph Style / Body Text / Body Text 2 / No Spacing | p |
| Heading 1〜6 | h1〜h6 |
| Title / TOC Heading | h1 |
| Subtitle | h2 |
| Quote / Intense Quote | blockquote |
| List Paragraph / List Bullet / List Number | li |
| Caption | figcaption |
| Header | header |
| Footer | footer |
| Hyperlink | a |
| Emphasis | em |
| Strong | strong |

### 使い方

1. Word ファイルを配置・取り込んだドキュメントを開く
2. スクリプトを実行する
3. 結果のレポートを確認する

### 制限事項・メモ

- 対象は段落スタイルだけです。文字スタイルは変換しません。
- `[基本段落]` など角かっこで始まるスタイルは対象外です。
- 対応表に無い名前は「対応なし」、すでに変換後の名前になっているものは「変更不要」としてスキップします。
- 統合先は、同じグループ → ドキュメント直下 → ほかのグループの順に探し、最初に見つかったスタイルを使います。
- 対応表はスクリプト冒頭の `STYLE_NAME_MAP` で変更できます。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdConvertWordStyleNamesToHtml.jsx` |
| バージョン | v1.0.2 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-09-20 |
| 最終更新 | 2026-09-30 |

### 更新履歴

#### v1.0.2（2026-09-30）

- ローカライズの処理を共通部品に統一（動作は同じ）

- v1.0.1（2026-09-20）初期バージョン

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
