# ExtendScript ファイルを選んで実行

[![Direct](https://img.shields.io/badge/Direct%20Link-IdScriptRunner.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/runner/IdScriptRunner.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdScriptRunner.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

任意の ExtendScript ファイル（.jsx / .jsxbin / .js）をダイアログで選んで実行するランチャーです。

### 主な機能

- ファイル選択ダイアログから実行するスクリプトを指定
- 実行前にファイルの存在と拡張子を確認
- 実行エラー時はファイル名・行番号・エラー番号とメッセージを表示

### 使い方

1. スクリプトを実行する
2. 実行したい ExtendScript ファイルを選ぶ

### 制限事項・メモ

- 取り消し単位は実行されるスクリプト側で管理されます。
- 最小構成の IdScriptRunnerSimple.jsx もあります。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/runner/IdScriptRunner.jsx` |
| バージョン | v1.0.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-17 |
| 最終更新 | 2026-09-30 |

### 更新履歴
- v1.0.2（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）

#### v1.0.1（2026-09-30）

- ローカライズとファイル選択の処理を共通の書き方に統一（動作は同じ）

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
