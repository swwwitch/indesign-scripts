# タイプスケールから段落スタイルを一括適用

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTypescale.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdTypescale.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTypescale.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

基準サイズとスケール倍率からタイプスケールを組み立て、本文・見出し・キャプションの段落スタイルへ一括適用します。

## 主な機能

- 基準サイズ（本文）、スケール倍率、見出しレベル数、サイズの丸め方を指定
- フォントは「本文と見出しで共通／別々に指定／変更しない」から選択
- 本文・見出しの行送り（%）、段落後のアキ（%）、カーニング方式を指定
- h1〜h6・本文・キャプションのサイズと行送りをプレビューで確認
- ライブプレビューで調整しながら確認

## 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する
3. 基準サイズと倍率を指定し、プレビューを見ながら調整して［OK］

## 制限事項・メモ

- 見出しは左揃え、本文とキャプションは左均等に設定されます。
- ダイアログを閉じるとオーバーライドを消去して最終設定を適用します。
- より多機能な IdTypeScaleStyleApplier.jsx もあります。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdTypescale.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-05-05 |
| 最終更新 | 2026-05-05 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
