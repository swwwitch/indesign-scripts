# ScriptRunner

任意の ExtendScript ファイル（.jsx / .jsxbin / .js）をダイアログで選んで実行するランチャーです。

## 主な機能

- ファイル選択ダイアログから実行するスクリプトを指定
- 実行前にファイルの存在と拡張子を確認
- 実行エラー時はファイル名・行番号・エラー番号とメッセージを表示

## 使い方

1. スクリプトを実行する
2. 実行したい ExtendScript ファイルを選ぶ

## 制限事項・メモ

- 取り消し単位は実行されるスクリプト側で管理されます。
- 最小構成の ScriptRunner-simple.jsx もあります。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/runner/ScriptRunner.jsx` |
| バージョン | v1.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-17 |
| 最終更新 | 2026-04-17 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
