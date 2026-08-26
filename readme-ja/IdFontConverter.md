# フォントの種別・ウエイトをまとめて切り替え

[![Direct](https://img.shields.io/badge/Direct%20Link-IdFontConverter.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/font/IdFontConverter.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdFontConverter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

フォントの種別（文字セット・P・UD・N・NT・ウエイト）をまとめて切り替え、選択範囲やドキュメント全体へ適用します。

## 主な機能

- 対象を「選択範囲／ストーリー／ドキュメント全体／アクティブスプレッド」から選択
- 文字セット（Std / Pro / Pr5 / Pr6）、N・NT、UD・P を「現状維持 / なし / あり」で切り替え
- 収録の最も多い文字セットへ寄せる Max / MaxN プリセット
- 段落スタイル・文字スタイル・合成フォントの各エントリ、ロック／非表示オブジェクトも対象に含められる
- 実行前に変更内容（旧 → 新）を和文フォント名でプレビュー確認。未インストールフォントは事前に警告

## 使い方

1. 対象のドキュメントを開く（選択範囲を対象にする場合はテキストやフレームを選択）
2. スクリプトを実行する
3. 対象と変換設定を指定して［実行］
4. プレビューで変更内容を確認し、必要な項目だけを残して確定

## 制限事項・メモ

- AXIS（Type Project）は文字セット体系が異なるため専用処理になります（幅と Joyo は保持し、N と Std⇄Pro のみ切り替え）。
- 同名ウエイトが見つからない場合は近いウエイトへ置換します。
- 適用は textStyleRange 単位でまとめて処理し、全体を 1 回の取り消しにまとめます。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/font/IdFontConverter.jsx` |
| バージョン | v1.1.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-06-17 |
| 最終更新 | 2026-06-30 |
| 紹介記事 | https://note.com/dtp_tranist/n/n261c771b4b41 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
