# 同じ段落スタイルで繰り返すテキストに連番を付ける

[![Direct](https://img.shields.io/badge/Direct%20Link-IdAppendParagraphNumbering.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/page/IdAppendParagraphNumbering.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAppendParagraphNumbering.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

同じテキストが同じ段落スタイルで繰り返すとき、末尾に連番を付けます。直近の親見出しごとに重複を判定します。

![ダイアログ。左に段落スタイルと対象範囲、右に繰り返しているテキストの一覧](../png/ss-1154-1016-144-20260813-223137.png)

## 主な機能

- 繰り返しているテキストを出現回数の多い順に一覧表示し、複数選択で対象を指定
- 段落スタイルのチェックボックスで一覧を絞り込み
- 選択範囲／ストーリー／ドキュメントの範囲切り替え
- 全角／半角括弧の切り替え（日本語 UI のみ）
- 既存のナンバリングを削除する［削除］ボタン

## 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する（解析後にダイアログが開きます）
3. 右のリストから番号を付けたい項目を選ぶ（複数選択可）
4. 対象の範囲を選んで［追加］

既存の番号を消したいときは、同じように項目を選んで［削除］を押します。

## 重複の判定について

「段落スタイル」「テキスト」「直近の親見出し」の 3 つが一致する段落を同じグループとして数え、2 回以上出現するものだけが一覧に並びます。

親見出しとして扱われるのは `HEADING_LEVEL_MAP` に定義された段落スタイル名（`h1`〜`h6` / `Heading 1`〜`Heading 6`）だけです。それ以外の名前を使っている文書では親なしとして扱われ、段落スタイルとテキストだけで判定されます。別の見出し名を使っている場合は、スクリプト内の `HEADING_LEVEL_MAP` に追加してください。

## 制限事項・メモ

- 対象は開いているアクティブドキュメントです。
- 親ページ（マスターページ）上のテキストは対象外です。
- `p.img` `p.table` の段落は対象外です（`IGNORE_STYLE_NAMES` で変更できます）。
- 1 文字以下のテキストは対象外です。
- ［選択範囲］を選んでもテキストが選択されていない場合は、確認メッセージを出したうえで広い範囲に切り替わります。
- 番号の付与も削除も 1 回の取り消しで元に戻せます。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/page/IdAppendParagraphNumbering.jsx` |
| バージョン | v1.2.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2025-06-30 |
| 最終更新 | 2026-08-13 |
| 紹介記事 | https://note.com/dtp_tranist/n/nc96549bb60f9 |

## 更新履歴

### v1.2.0（2026-08-13）

- 対象に［選択範囲］を追加。選択したテキストの範囲だけに番号を付けられます
- 親見出しが見つからない段落が対象から漏れていた問題を修正。見出しとして扱われるのは `HEADING_LEVEL_MAP` にある名前だけなので、それ以外の見出しスタイルを使った文書では対象が 1 件も出ませんでした
- 空段落や 1 文字の見出しがあると、選んだ対象が処理されないことがある問題を修正
- 日本語 UI で［キャンセル］がダイアログを閉じない問題を修正
- グループ内に配置された親ページのテキストフレームを、対象から正しく除外するように修正
- ［削除］の実行時にダイアログを閉じるように変更（実行後もリストが古いまま残っていました）
- ドキュメント未オープン時と、範囲指定に対して選択がないときのメッセージを追加
- ［OK］を［追加］に変更し、ボタンを左（削除）／右（キャンセル・追加）の配置に整理
- 解析処理を高速化

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
