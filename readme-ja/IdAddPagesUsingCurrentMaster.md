# 現在の親ページを引き継いでページを挿入

[![Direct](https://img.shields.io/badge/Direct%20Link-IdAddPagesUsingCurrentMaster.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/page/IdAddPagesUsingCurrentMaster.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAddPagesUsingCurrentMaster.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

現在のページに適用されている親（マスター）ページを引き継いだまま、指定した枚数のページを直後に挿入します。

### 主な機能

- 現在のページ名と適用中の親ページをダイアログに表示
- 挿入するページ数を入力（初期値 2）
- 追加したページに現在の親ページを自動適用
- ドキュメント基準で追加するため、見開きのスプレッドへ正しく流し込まれる

### 使い方

1. ページを挿入したい位置のページをアクティブにする
2. スクリプトを実行する
3. 挿入するページ数を入力して［OK］

### 制限事項・メモ

- アクティブな InDesign ドキュメントが必要です。
- 標準の［ページ挿入］ダイアログと違い、選択中ページの親ページを引き継ぎます。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/page/IdAddPagesUsingCurrentMaster.jsx` |
| バージョン | v1.3.2 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2025-06-26 |
| 最終更新 | 2026-10-01 |
| 紹介記事 | https://note.com/dtp_tranist/n/n2d1a76097a39 |

### 更新履歴
- v1.3.3（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）

#### v1.3.2（2026-10-01）

- 右側のボタンだけの行は、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更

#### v1.3.1（2026-09-30）

- ボタン行・余白・ローカライズの処理を共通部品に統一（動作は同じ）

- v1.3.0（2026-09-27）：ページ数の欄にステップボタン（∧∨）を追加。↑↓キーでも増減できる（次の整数へ、shift＋で次の10の倍数へ。1未満にはならない）

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
