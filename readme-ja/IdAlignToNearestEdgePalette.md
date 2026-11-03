# 移動方向でいちばん近いガイド・マージン・ページ・裁ち落としに揃えるパレット

[![Direct](https://img.shields.io/badge/Direct%20Link-IdAlignToNearestEdgePalette.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/frame/IdAlignToNearestEdgePalette.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAlignToNearestEdgePalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択したオブジェクトを、パレットのボタンで選んだ方向へ動かし、その方向でいちばん近いガイド・マージンガイド・ページのエッジ・裁ち落としに辺を揃えます。

### 主な機能

- ［←］［→］［↑］［↓］の4方向。左なら左辺、上なら上辺のように、移動方向の辺を揃える
- 揃える先は、ガイド（スプレッド・ページ・親ページ）・マージンガイド・ページのエッジ・裁ち落としのうち、移動方向でいちばん近い線
- 続けて押すと、次の線へ進む（例：マージン → ページのエッジ → 裁ち落とし）
- オブジェクトごとに、それぞれいちばん近い線へ動かす
- ［選択をまとめて移動］で、移動方向の先頭にあるオブジェクトを基準に、相対位置を保ったまま全体を動かす
- パレットを選んでいるときは、矢印キーでも動かせる。Esc で閉じる
- 1回の移動が1回の取り消しにまとまる
- ［選択をまとめて移動］とパレットの位置は、InDesign の終了まで覚える。日本語／英語の自動切り替え

### 使い方

1. スクリプトを実行してパレットを開く
2. 動かすオブジェクトを選択する（文字ツールで文字を選んでいるときは、そのテキストフレームが対象）
3. 動かしたい方向のボタンをクリックする

### 制限事項・メモ

- 辺は線の太さを含む境界（表示上の境界）で測ります。
- 移動方向に線が無いオブジェクトはその場に残します。どのオブジェクトも動かなかったときだけ、警告を表示します。
- ロックしたオブジェクト、ロックしたレイヤーのオブジェクト、アンカー付きオブジェクトは対象にしません。
- 非表示のレイヤーにあるガイドは使いません。親ページのガイドは、ページに直接当たっている親ページのものだけを使います（親ページの元になった親ページのガイドは使いません）。
- 段組みのガイドは使いません。
- パレットを開いたままスクリプトを再実行すると、前のパレットを閉じてから開き直します。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/frame/IdAlignToNearestEdgePalette.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-11-03 |
| 最終更新 | 2026-11-03 |

### 更新履歴

- v1.0.0（2026-11-03）初版

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
