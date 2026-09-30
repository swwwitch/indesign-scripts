# 日本語組版設定をマトリックスで一括適用

[![Direct](https://img.shields.io/badge/Direct%20Link-IdJapaneseParagraphTypesettingManager.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdJapaneseParagraphTypesettingManager.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdJapaneseParagraphTypesettingManager.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

段落スタイルの日本語組版設定（禁則処理セット・禁則調整方式・文字組みアキ量・コンポーザー）をマトリックス UI で確認・一括適用します。

### 主な機能

- 段落スタイルごとに現在の組版設定を 1 行で表示
- 先頭の「すべて」行をコピー元として、列ごとの［↓ 反映］ボタンで全段落スタイルへ反映
- 既存の段落スタイル設定を読み取り、ダイアログの初期値に反映
- カスタム文字組みアキ量設定と組み込みプリセット名の両方に対応
- 段落スタイルグループを再帰的に走査し、対象外グループを除外

### 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する
3. マトリックスで設定を調整して［OK］

### 制限事項・メモ

- ［段落スタイルなし］［基本段落］と、名前が「_」で始まるグループ配下のスタイルは対象外です。
- 既定値はスクリプト冒頭の DEFAULT_* 変数で変更できます。
- 全処理が 1 回の取り消しにまとまります。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdJapaneseParagraphTypesettingManager.jsx` |
| バージョン | v1.3.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-05-05 |
| 最終更新 | 2026-10-01 |

### 更新履歴

#### v1.3.1（2026-10-01）

- 右側のボタンだけの行は、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更

#### v1.3.0（2026-09-30）

- 文字組みが「なし」の段落スタイルに、そのまま OK すると「行末約物半角」が入ってしまう問題を修正（「なし」と表示して保持）
- 組み込みの文字組みアキ量プリセットを読み取れていなかった問題を修正
- 禁則・文字組み・コンポーザーの一覧づくりを IdTypesettingStyleManager と共通化。コンポーザーの表示名を「多言語対応段落コンポーザー」などに
- ボタン行・余白・ローカライズの処理を共通部品に統一

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
