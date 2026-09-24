# 表の行の高さをプレビュー付きで設定

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTableRowHeightManager.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdTableRowHeightManager.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableRowHeightManager.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択した表の行の高さを、範囲（選択範囲／ストーリー／ドキュメント）と対象行を指定しながらプレビュー付きで設定します。

![ダイアログ。左に範囲と対象、右に行の高さとオプションのパネル](../png/ss-952-692-144-20260925-063059.png)

## 主な機能

- 範囲を「選択範囲／ストーリー／ドキュメント」から選択し、複数の表へまとめて適用
- 対象を「表全体」「表全体（ヘッダー行を除く）」「選択した行のみ」から選択
- 行の高さは「最小限度」または「指定値」から選択
- 初期値には現在の行高を使用（すべて同じならその値、異なればヘッダー行を除く平均）
- ［フレームを内容に合わせる］（初期値 ON）で、表を含むテキストフレームの高さを内容に合わせる（プレビューにも反映）
- ［プレビューモード］ボタンで画面モードを切り替え、ガイドや枠線を隠して確認できる

## 使い方

1. 表、セル、または表を含むテキストフレームを選択する
2. スクリプトを実行する
3. 範囲・対象・行の高さを指定して［OK］

## 制限事項・メモ

- 「選択した行のみ」は範囲が「選択範囲」のときだけ有効です。
- 複数の表が混在する選択はエラーになります。
- 入力値はドキュメントの縦方向単位で表示・入力し、内部では pt に変換して適用します。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/table/IdTableRowHeightManager.jsx` |
| バージョン | v1.3.2 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-20 |
| 最終更新 | 2026-09-25 |
| 紹介記事 | https://note.com/dtp_tranist/n/n9f95f8e98db6 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
