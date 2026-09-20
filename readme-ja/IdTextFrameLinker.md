# 複数のテキストフレームを連結

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTextFrameLinker.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdTextFrameLinker.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTextFrameLinker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択した複数のテキストフレームを、選択順に連結して1つのストーリーにします。連結後に、1つ目のフレームの高さを調整したり、空になったフレームを削除したりできます。

## 主な機能

- 選択した2つ以上のテキストフレームを、選択順に `nextTextFrame` で連結
- ［1つ目のテキストフレームの高さを調整］：幅はそのままに、テキストが収まる高さまで広げる
- ［空になったテキストフレームを削除］：連結後に空になったフレームだけを削除する（1つ目は残す）
- どちらもオフなら、フレームのサイズや位置は変更しない
- 選択が2つ未満、またはテキストフレーム以外を含む場合はメッセージを表示
- 連結すると循環する場合は実行せずにメッセージを表示

## 使い方

1. 連結する順に、2つ以上のテキストフレームを選択する
2. スクリプトを実行する
3. 連結後の処理をチェックボックスで選んで［OK］

## 制限事項・メモ

- 連結の順序は選択した順序で決まります。1つ目に選択したフレームが先頭になります。ドラッグで囲んで選択した場合は重ね順になるため、クリックで順に選択してください。
- 2つ目以降のフレームに文字が入っている場合も内容は失われず、選択順に1つのストーリーへ結合されます。
- 選択したフレームがすでに連結されている場合、既存の連結は分断されず、その間に差し込む形でつなぎ直されます。
- 削除は2つ目以降の空のフレームだけが対象です。1つ目のフレームと、文字が残っているフレームは削除しません。
- チェックボックスの初期値はスクリプト冒頭の「初期設定」ブロック（`DEFAULT_FIT_FIRST_HEIGHT` / `DEFAULT_DELETE_EMPTY_FRAME`）で変更できます。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/text/IdTextFrameLinker.jsx` |
| バージョン | v1.0.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2023-12-26 |
| 最終更新 | 2026-09-20 |
| 紹介記事 | https://note.com/dtp_tranist/n/n04ceaf4955a0 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
