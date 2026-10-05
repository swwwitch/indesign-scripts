# 同一スタイル間の段落間隔をスタイル定義に設定

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSetSameParaStyleSpacing.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdSetSameParaStyleSpacing.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSetSameParaStyleSpacing.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

段落スタイルの「同一スタイル間の段落間隔」を、スタイル定義そのものに対して設定します。

### 主な機能

- グループ内の段落スタイルも再帰的に収集してドロップダウンに表示
- 間隔は「無視」「0」「数値指定」から選択
- 数値は環境設定の表示単位に連動し、∧∨・↑↓ キーで次の整数へ増減（1.5→2、Shift＝次の 10 の倍数、Option＝0.1 刻み）
- テキスト選択中なら、適用中の段落スタイルを初期選択にして現在値も反映

### 使い方

1. 対象のドキュメントを開く（設定したい段落にカーソルを置いておくと初期選択が便利）
2. スクリプトを実行する
3. 段落スタイルと間隔を指定して［OK］

### 制限事項・メモ

- 変更はスタイル定義に対して行われるため、そのスタイルを使うすべての段落に反映されます。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdSetSameParaStyleSpacing.jsx` |
| バージョン | v1.1.2 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-06-30 |
| 最終更新 | 2026-10-01 |

### 更新履歴
- v1.1.3（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）
- v1.2.0（2026-10-06）長さの入力欄に単位を入れた（「10 mm」の形）。別の単位で入れた値（「1in」など）や計算式も欄の単位へ換算する

#### v1.1.2（2026-10-01）

- 右側のボタンだけの行は、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更

#### v1.1.1（2026-09-30）

- ボタン行・余白・ローカライズの処理を共通部品に統一（動作は同じ）

#### v1.1.0（2026-09-27）

- 数値指定の欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ、option＋で0.1ずつ）

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
