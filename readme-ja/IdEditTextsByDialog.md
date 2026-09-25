# ダイアログでテキストを編集して置換・挿入

[![Direct](https://img.shields.io/badge/Direct%20Link-IdEditTextsByDialog.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdEditTextsByDialog.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdEditTextsByDialog.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

複数行入力ダイアログでテキストを編集し、選択範囲の置換・カーソル位置への挿入・新規テキストフレーム作成を行います。

## 主な機能

- テキスト選択中は置換、カーソル位置のみなら挿入、テキストフレーム選択時は末尾へ挿入
- 何も選択していない場合はページ中央に新規テキストフレームを作成
- `@#` を強制改行（\n）の可視マーカーとして扱う
- ［改行を削除］ボタンで改行とマーカーを一括削除

## 使い方

1. 編集したいテキストを選択する（未選択でも実行可）
2. スクリプトを実行する
3. 入力欄で編集して［OK］

## 制限事項・メモ

- 入力欄の Enter は段落（\r）になります。
- 強制改行を入れたいときは［@# を追加］ボタンを使います（入力欄の末尾に追加されます）。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/text/IdEditTextsByDialog.jsx` |
| バージョン | v0.1.4 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2025-05-28 |
| 最終更新 | 2026-09-25 |

## 更新履歴

### v0.1.4（2026-09-25）

- ダイアログのタイトルを「テキストを編集」に変更
- ボタン名を［改行を削除］［@# を追加］に変更
- 入力欄と左側のボタンにツールチップを追加
- ドキュメントが開いていないときはメッセージを出して終了
- 1 文字だけの特殊文字を選択して実行するとエラーになる問題を修正

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
