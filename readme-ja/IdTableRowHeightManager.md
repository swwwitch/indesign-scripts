# 表の行の高さをプレビュー付きで設定

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTableRowHeightManager.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdTableRowHeightManager.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableRowHeightManager.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択した表の行の高さを、範囲（選択範囲／ストーリー／ドキュメント）と対象行を指定しながらプレビュー付きで設定します。

## 主な機能

- 範囲を「選択範囲／ストーリー／ドキュメント」から選択し、複数の表へまとめて適用
- 対象を「表全体」「表全体（見出し行を除く）」「選択した行のみ」から選択
- 行の高さは「最小限度」または「指定値」から選択
- 初期値には現在の行高を使用（すべて同じならその値、異なれば見出し行を除く平均）
- ［親フレームの調整］ボタンで親テキストフレームを内容に合わせる

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
| バージョン | v1.3.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-20 |
| 最終更新 | 2026-06-09 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
