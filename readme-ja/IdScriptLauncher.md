# スクリプトをキーワードで絞り込んで実行

[![Direct](https://img.shields.io/badge/Direct%20Link-IdScriptLauncher.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/misc/IdScriptLauncher.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdScriptLauncher.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

指定したフォルダー内の .jsx / .js / .jsxbin をキーワードで絞り込み、選んだスクリプトをその場で実行するランチャーです。フォルダーとファイル名を左右のリストに分けて表示します。

## 主な機能

- キーワード欄に入力すると、フォルダーリストとファイル名リストが即時に絞り込まれる（スペース区切りで AND 検索）
- 大文字小文字・全角半角・ひらがなカタカナ・濁点や小書きの違いを無視して一致（「ﾃｰﾌﾞﾙ」「ﾃﾞｰﾀ」なども拾う）
- 絞り込んだ結果によく出てくる語を、ワンクリックのキーワードボタンとして自動で表示
- 「サブディレクトリを含む」で、対象フォルダー直下までに限るか、下位すべてをたどるかを切り替え
- 「フルパス」で、対象フォルダーのパス表示をフルパスとホームフォルダーを `~` に略した表記で切り替え
- 環境設定で、対象フォルダーとキーワードボタンの出現条件（出現数・キーワード数）を変更
- 対象フォルダーとキーワードボタンの設定は次回起動時に引き継がれる

## 使い方

1. スクリプトを実行する
2. 初回は検索対象のスクリプトフォルダーを選ぶ（次回からは前回のフォルダーが開く）
3. キーワードを入力するか、キーワードボタンをクリックして絞り込む
4. ファイル名を選んで［実行］、またはダブルクリック

### キーボード操作

| 操作 | 動作 |
| --- | --- |
| `↓`（キーワード欄） | ファイル名リストへ移動 |
| `Enter` | 選択中のスクリプトを実行（ダイアログ内のどこからでも） |
| 環境設定の数値欄で `↑` `↓` | ±1（`shift` で ±10、`option` で ±0.1） |

### マウス操作

| 操作 | 動作 |
| --- | --- |
| ファイル名をダブルクリック | 実行 |
| `option` + ファイル名をダブルクリック | 実行せず Finder で表示 |
| フォルダーをダブルクリック | そのフォルダーを Finder で開く |
| `option` + キーワードボタン | 入力中のキーワードを置き換えず、空白を挟んで追加（AND 検索） |

## 制限事項・メモ

- 取り消し単位は実行されるスクリプト側で管理されます（`doScript` でラップしません）。
- キーワードボタンになるのは英字のみ・3文字以上の語です。日本語のファイル名は絞り込みでは拾えますが、ボタンにはなりません。
- 対象フォルダーにこのランチャー自身がある場合、一覧には表示されません。
- エイリアスのフォルダーは循環を避けるためたどりません。
- 「Finder で表示」には `/Applications/RevealInFinder.app` を使います。無い場合は囲みフォルダーを開きます（macOS 以外も同様）。
- 設定は `IdScriptLauncher-prefs.txt`（ユーザーデータフォルダー）に保存されます。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/misc/IdScriptLauncher.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-08-26 |
| 最終更新 | 2026-08-26 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
