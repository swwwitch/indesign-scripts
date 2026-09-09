# 画像入りフレームを順送りに入れ替え

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSwapImageFrames.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/frame/IdSwapImageFrames.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSwapImageFrames.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択した画像入りフレームを、リンク画像だけ／フレームの位置ごと、いずれかの方法で順送りに入れ替えます。

## 主な機能

- 「画像だけ入れ替え（フレームは固定）」は既存画像を削除し、別フレームのリンク画像を順送りに再配置
- 「フレームごと（位置を移動）」は中身を保ったままフレームの位置だけを入れ替え
- 配置後のフィット方法（フレームに均等に流し込む／内容を縦横比率に応じて合わせる）と、位置合わせの基準（左上／中央）を選択
- 入れ替え順は見た目の位置で「上から下、左から右」。1つ後ろのフレームへ送り、最後のフレームは先頭へ回します
- 各項目にツールチップを表示（マウスを重ねると動作の説明が出ます）
- 一連の処理は1回の取り消し（Cmd+Z）で元に戻せます

## 使い方

1. 画像入りフレームを 2 つ以上選択する
2. スクリプトを実行する
3. 入れ替えモードとオプションを指定して［OK］

## 制限事項・メモ

- 同じフレームを重複して選んだり、フレームと中の画像を同時に選んだ場合も 1 回だけ処理します。
- 「画像だけ入れ替え（フレームは固定）」は各フレームに主画像が 1 点だけある前提です。リンクのない画像（埋め込みなど）が含まれる場合は実行できません。
- 「フレームごと（位置を移動）」ではフレームのサイズは変わりません。サイズが異なるフレーム同士では、基準（左上／中央）に応じて見え方が変わります。
- 削除・再配置・フィットに失敗した場合はその時点で中止してエラーを表示します。すでに入れ替わった分は取り消しで元に戻せます。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/frame/IdSwapImageFrames.jsx` |
| バージョン | v1.0.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-03-28 |
| 最終更新 | 2026-09-09 |
| 紹介記事 | https://note.com/dtp_tranist/n/n6dee03ae96e2 |

## 更新履歴

### v1.0.1（2026-09-09）

- ダイアログの文言を見直し。動作が「2つの交換」ではなく「順送り」であることが分かる表現にし、フィットの選択肢は InDesign 標準の用語に合わせました
- エラーメッセージに、途中で中止したことと取り消しで戻せることを明記
- 各ラジオボタンとパネルにツールチップを追加
- 配置画像そのものを選択したときのフレーム判定を簡素化。画像の種別によらず、親が画像フレームであれば対象にするようにしました
- 紹介記事へのリンクを追加
- 内部整理（命名の統一、順送り処理の共通化、エラー処理の整理）

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
