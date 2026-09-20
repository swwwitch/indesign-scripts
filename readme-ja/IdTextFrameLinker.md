# 2つのテキストフレームを連結

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTextFrameLinker.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdTextFrameLinker.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTextFrameLinker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択した2つのテキストフレームを連結し、1つのストーリーにします。

## 主な機能

- 選択した2つのテキストフレームを `nextTextFrame` で連結
- フレームのサイズや位置は変更しない
- ダイアログボックスなしで実行
- 選択が2つのテキストフレームでない場合はメッセージを表示
- 1つ目のフレームにオーバーフローテキストがある場合は連結せずにメッセージを表示

## 使い方

1. 連結元・連結先の順に、2つのテキストフレームを選択する
2. スクリプトを実行する

## 制限事項・メモ

- 連結の向きは選択した順序で決まります。1つ目に選択したフレームが連結元になります。
- 1つ目のフレームにあふれたテキストがあると、連結時に2つ目のフレームの内容が押し出されるため、処理を中止します。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/text/IdTextFrameLinker.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-09-20 |
| 最終更新 | 2026-09-20 |
| 紹介記事 | https://note.com/dtp_tranist/n/n04ceaf4955a0 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
