# 段落スタイルの文字組版設定をまとめて設定

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTypesettingStyleManager.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdTypesettingStyleManager.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTypesettingStyleManager.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

段落スタイルの文字組版設定（禁則・文字組み・グリッド揃え・ハイフネーションなど）をダイアログでまとめて設定します。

### 主な機能

- 対象を「選択中」「すべて」「指定」から選択
- 選択中の段落から現在の組版設定・言語・ハイフネーション設定を読み込んで初期値に利用
- プリセット（欧文組版／グリッド優先／グリッド無視／ソースコード／InDesign のデフォルト）の適用
- 現在の設定をプリセットコードとしてデスクトップへ書き出し
- ハイフネーションの ON/OFF に応じて関連項目を有効／無効化
- 数値欄の∧∨・↑↓キーで次の整数へ増減（1.5→2）、shift併用で次の10の倍数へ、option併用で0.1ずつ

### 使い方

1. 対象のドキュメントを開く（選択中の段落を初期値にしたい場合はカーソルを置く）
2. スクリプトを実行する
3. 各設定を選んで［OK］

### 制限事項・メモ

- ［段落スタイルなし］［基本段落］と、名前が「_」で始まるグループ配下のスタイルは対象外です。
- 適用後は選択範囲のオーバーライドを常に消去します。
- 引用符・言語・単位は環境設定側に適用されます。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdTypesettingStyleManager.jsx` |
| バージョン | v1.3.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-05-06 |
| 最終更新 | 2026-10-01 |
| 紹介記事 | https://note.com/dtp_tranist/n/n7f67e8da571f |

### 更新履歴
- v1.3.2（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）

#### v1.3.1（2026-10-01）

- 右側のボタンだけの行は、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更

#### v1.3.0（2026-09-30）

- 日本語 UI の項目名が「field.autoKerning」のようなキー文字列で表示されていた問題を修正
- 禁則・文字組み・コンポーザーの一覧づくりを IdJapaneseParagraphTypesettingManager と共通化。残っていた日本語の直書きを英語 UI にも対応
- ボタン行・余白・ローカライズの処理を共通部品に統一

#### v1.2.0（2026-09-27）

- 数値欄（自動行送り・ハイフネーション）にステップボタン（∧∨）を追加。↑↓キーでも同じ処理で増減できるように（次の整数へ、shift＋で次の10の倍数へ）
- 読み込み時にエラーになり起動しなかった不具合を修正

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
