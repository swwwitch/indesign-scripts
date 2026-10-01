# 表の列幅をまとめて調整

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTableColumnWidthAdjuster.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/table/IdTableColumnWidthAdjuster.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTableColumnWidthAdjuster.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択したセルや表の列幅を、列ごとの個別指定または一括入力でまとめて調整します。

### 主な機能

- 入力方法を「個別に設定」と「一括入力」から選択
- 指定方法は「幅で指定」または「文字数で指定」
- 列ごとに幅／文字数換算／左右の余白／自動調整を設定
- 自動調整は、2 行目がある列では 2 行目が消えるまで広げ、無い列では文字数ベースで推定
- 「全列に適用」で編集中の値を他列へ連動反映

### 使い方

1. 表内のセルにカーソルを置くか、表を選択する
2. スクリプトを実行する
3. 入力方法と各列の値を指定して［OK］

### 制限事項・メモ

- 列幅と左右の余白はドキュメントの単位設定に従います。
- 文字数換算は、表内で最も支配的なフォントサイズを基準に計算します。
- 変更は常に即時反映され、キャンセル時は元に戻ります。
- 一括入力はスペース区切りまたはカンマ区切りです（例: 30 50 70 70 / 30, 50, 70, 70）。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/table/IdTableColumnWidthAdjuster.jsx` |
| バージョン | v1.3.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-04-19 |
| 最終更新 | 2026-10-01 |
| 紹介記事 | https://note.com/dtp_tranist/n/n20d58c1bc003 |

### 更新履歴

#### v1.3.1（2026-10-01）

- 折り返しのある列の［自動調整］で、左右の余白が二重に足されて幅が広くなりすぎていた問題を修正
- 定規の単位が Q・H・アゲート・シセロ・ピクセルのとき、文字数の換算と単位の表示がずれていた問題を修正
- ダークUIで見出しや列番号が黒く表示されていた問題を修正
- ［文字数で指定］［全列に適用］［左右の余白］［自動調整］と一括入力欄にツールチップを追加
- ［指定方法］のラジオボタンを横並びに変更
- 英語の見出しと一括入力の説明を短く整理
- コードを整理（関数の分割・重複の統合・不要な処理の削除）

#### v1.3.0（2026-09-30）

- ［全列に適用］で値を反映するときにエラーになっていた問題を修正
- 画面モードの切り替えボタンの表記を、他のスクリプトと同じ「押すと切り替わる先」に統一
- ボタン行・表の選択・ローカライズの処理を共通部品に統一

#### v1.2.0（2026-09-27）

- 各列の幅・文字数・左右の余白の欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ、option＋で0.1ずつ）

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
