# 段落のページ番号を「p.番号」の相互参照で挿入

[![Direct](https://img.shields.io/badge/Direct%20Link-IdInsertPageCrossReference.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdInsertPageCrossReference.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdInsertPageCrossReference.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

1つ目のストーリーの3段落目が載っているページの番号を、2つ目のストーリーの末尾に「p.12」の形の相互参照で挿入します。

ページ番号を文字で書き込むと、ページが増減したときに手で直す必要があります。相互参照にしておけば、番号はページの移動に追従し、クリックで参照先へ移動するリンクにもなります。

### 主な機能

- 参照先の段落に段落の参照先（相互参照の行き先）を作り、書き込み先に相互参照を挿入
- 相互参照の書式「p.ページ番号」（文字列「p.」＋ページ番号）を作成。同じ名前の書式があればそれを使う
- 取り消しは1回で元に戻る

### 使い方

1. ドキュメントを開いてスクリプトを実行する
2. 2つ目のストーリーの末尾に「p.番号」が入り、［相互参照］パネルに追加される

### 制限事項・メモ

- 参照先（1つ目のストーリーの3段落目）と書き込み先（2つ目のストーリーの末尾）は決め打ちです。ストーリーの順番はページ上の位置や作った順とは一致しないため、ドキュメントによっては意図しない段落が対象になります。
- ストーリーが2つ未満、または1つ目のストーリーの段落が3つ未満のときは、メッセージを出して終了します。
- 参照先がオーバーセットやペーストボード上にあるときの表示は InDesign の相互参照の扱いに従います。
- 実機での動作は未確認です。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/document/IdInsertPageCrossReference.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-11-06 |
| 最終更新 | 2026-11-06 |

### 更新履歴

- v1.0.0（2026-11-06）：初版

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
