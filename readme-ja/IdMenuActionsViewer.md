# メニューアクションを調べて実行コードをコピー

[![Direct](https://img.shields.io/badge/Direct%20Link-IdMenuActionsViewer.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/misc/IdMenuActionsViewer.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdMenuActionsViewer.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

InDesign のメニューアクション（メニューコマンド）を一覧し、エリア・メニュー・名前・ID で絞り込んで調べます。選択したアクションのキーストリング（`$ID/…`）と、スクリプトから実行するためのコードをコピーできます。

Peter Kahrel 氏の `menu_actions.jsx` をもとに改変したものです。

## 主な機能

- メニューアクションを「エリア｜名前」の2列のリストで表示
- 絞り込み
  - **エリア**：エリアを1つ選んで絞り込み
  - **メニュー**：ファイル／編集／表示／フォーマット／検索／ツール／ウィンドウ／パネルメニューから選ぶと、エリア名がそのメニュー名で始まるアクションに絞り込み。エリアの候補も絞られる
  - **名前・ID**：名前に含まれる文字で絞り込み（大文字小文字を区別しない。正規表現も可）。数字だけを入れると、その ID のアクションも対象にする
- 並び順をラジオボタンで切り替え（エリア／名前／ID）
- リストで選択したアクションの情報を、リストの下に自動で表示
  - **キーストリング**：`$ID/Find/Change...` など（複数ある場合は「 | 」区切り）
  - **実行コード**：`app.menuActions.itemByID(18694).invoke();`
- それぞれ［コピー］ボタンでクリップボードにコピー
- 絞り込んだ結果が1件なら、自動で選択して情報を表示
- 読み込み中は進行状況バーを表示
- ダイアログのタイトルに表示中の件数を表示

## 使い方

1. スクリプトを実行する
2. エリア・メニュー・名前・ID で目的のアクションを絞り込む（名前・ID は Enter で確定）
3. リストでアクションを選択する
4. 実行コードまたはキーストリングの［コピー］をクリックし、自分のスクリプトに貼り付ける

コピーした実行コードは、次のように使えます。

```javascript
app.menuActions.itemByID(18694).invoke();
app.menuActions.item("$ID/Find/Change...").invoke();
```

## 制限事項・メモ

- フォント名・スタイル名（ID 57603〜61066）、最近使用したファイル・スクリプト名（`.indd` `.jsx` `.jsxbin`）、英語版の「Text Selection」「Menu:Insert」エリアは一覧に出しません。
- 「メニュー」の絞り込みは、エリア名の先頭との一致で判定します。日本語版と英語版の名称を両方照合しますが、InDesign のバージョンや言語によっては該当しないメニューがあります。
- 実行コードの ID は InDesign のバージョンや環境によって変わることがあります。配布するスクリプトでは、キーストリング（`$ID/…`）を使う方法も検討してください。
- クリップボードへのコピーは、Mac では AppleScript、Windows では VBScript を経由します。
- 名前・ID の絞り込みは Enter で確定します（数千件の作り直しを入力のたびに行わないため）。

## オリジナル、謝辞

Peter Kahrel 氏の `menu_actions.jsx` をもとにしています。

- https://creativepro.com/menu_actions/
- http://kasyan.ho.com.ua/open_menu_item.html

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/misc/IdMenuActionsViewer.jsx` |
| バージョン | v1.0.14 |
| 原作 | Peter Kahrel |
| 改変 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-09-25 |
| 最終更新 | 2026-09-25 |

## 更新履歴

- v1.0.14（2026-09-25）：初版

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
