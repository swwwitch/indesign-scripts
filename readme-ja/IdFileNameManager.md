# ファイル名をセグメント単位で編集してリネーム・保存

[![Direct](https://img.shields.io/badge/Direct%20Link-IdFileNameManager.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdFileNameManager.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdFileNameManager.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

アクティブなドキュメントのファイル名を、ベース・サブテキスト・ステータス・日付・連番・バージョンのセグメント単位で編集し、リネーム／別名保存／コピー保存を行います。

## 主な機能

- セグメントは base / title / status / timestamp / page / version。並び順もカスタマイズ可能
- ステータス（wip / draft / review / approved / flattened など）をドロップダウンから挿入
- 連番（page01 / page001）は保存時にフォルダ内の最大値 +1 へ自動繰り上げ
- バージョンは v1 / v01 / v001 の書式を選べ、フォルダ内の最大値 +1 まで自動繰り上げ
- 区切り記号の統一、NFC 正規化、丸数字や法人略記の ASCII 化などのファイル名整形

## 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する
3. モードと各セグメント、並び順を指定して［OK］

## 制限事項・メモ

- 保存先は常に現在のファイルと同じフォルダです。未保存ドキュメントは保存先フォルダを選択します。
- 元ファイルに v 番号が無い場合、バージョンは v01 形式で付与されます。
- 設定はスクリプト実行のたびに保存され、次回の初期値になります。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/document/IdFileNameManager.jsx` |
| バージョン | v1.3.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-05-27 |
| 最終更新 | 2026-05-29 |
| 紹介記事 | https://note.com/dtp_tranist/n/nc88dd887eb1c |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
