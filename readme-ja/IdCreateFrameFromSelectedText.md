# 選択テキストのサイズでグラフィックフレームを作成

[![Direct](https://img.shields.io/badge/Direct%20Link-IdCreateFrameFromSelectedText.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/frame/IdCreateFrameFromSelectedText.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCreateFrameFromSelectedText.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択したテキストのサイズを基準に、インライン（アンカー付き）またはページ上へグラフィックフレームを作成します。

## 主な機能

- 配置方法をインライン（アンカー付き）／ページ上のグラフィックフレームから選択
- フレーム幅を「選択したテキスト／カラム幅／親フレーム／ページのマージン」から選択
- テキスト未選択時は高さを行数または mm で指定
- 挿入行の段落スタイル、オブジェクトスタイル、テキストの回り込みを指定

## 使い方

1. フレームの基準にしたいテキストを選択する（未選択でも実行可）
2. スクリプトを実行する
3. 配置方法とサイズを指定して［OK］

## 制限事項・メモ

- 回り込みはページ配置かつオブジェクトスタイルが「なし」のときだけ有効です。
- 行数モードの高さは概算で、1 行目の文字サイズ＋残り行数×行送りで計算します。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/frame/IdCreateFrameFromSelectedText.jsx` |
| バージョン | v2.6 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-03-17 |
| 最終更新 | 2026-03-17 |
| 紹介記事 | https://note.com/dtp_tranist/n/ndd1b7c5246a3 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
