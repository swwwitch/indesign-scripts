# カーソル位置にアンカー付き長方形とテキストアンカーを作成

[![Direct](https://img.shields.io/badge/Direct%20Link-IdInsertAnchoredRectangle.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdInsertAnchoredRectangle.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdInsertAnchoredRectangle.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

カーソルを置いた位置に、アンカー付き（カスタム）でプリントしない小さな長方形と、その直前のテキストアンカーを作成します。

長方形とテキストアンカーは同じ名前で対応づけられるので、ハイパーリンクや相互参照の行き先として使いながら、紙面上では長方形で位置を確かめられます。

### 主な機能

- カーソル位置（テキストを選択しているときはその先頭）に、12×12 pt・塗り M100Y100・線なしの長方形を、アンカー付きオブジェクト（［アンカー付きオブジェクトオプション］の位置は「カスタム」）として挿入
- 塗りはスウォッチ「C=0 M=100 Y=100 K=0」。ドキュメントに無ければ作成する
- 長方形は［プリントしない］に設定（［属性］パネル）。プリントや PDF 書き出しには出ない
- 長方形のマーカー文字の直前にテキストアンカーを作成
- テキストアンカー名は「anchored_rectangle_<時刻>」。同じ名前を長方形のスクリプトラベルにも入れる
- 作成後は長方形を選択し、テキストアンカー名を表示
- 取り消しは1回で元に戻る

### 使い方

1. 行き先にしたいテキストの位置に文字ツールでカーソルを置く（またはテキストを選択する）
2. スクリプトを実行する
3. 長方形が入り、作成したテキストアンカー名が表示される
4. ［ハイパーリンク］パネルや［相互参照を挿入］で、このテキストアンカーを行き先に選ぶ

### 制限事項・メモ

- テキストフレームを選択ツールで選んだ状態や、オブジェクトを選んだ状態では実行できません。
- ペーストボード上やオーバーセットのテキストでは使えません。
- 大きさ・塗りの CMYK・名前の接頭辞は、スクリプト冒頭の基本設定で変えられます。
- 実機での動作は未確認です。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/document/IdInsertAnchoredRectangle.jsx` |
| バージョン | v2.2.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-11-06 |
| 最終更新 | 2026-11-06 |

### 更新履歴

- v2.2.0（2026-11-06）：長方形を、塗り M100Y100（スウォッチ「C=0 M=100 Y=100 K=0」、無ければ作成）・線なしに変更
- v2.1.0（2026-11-06）：アンカー付きオブジェクトの位置を「インライン」から「カスタム」に変更し、長方形を［プリントしない］に設定
- v2.0.0（2026-11-06）：IdInsertPageCrossReference から改名し、カーソル位置にアンカー付き長方形とテキストアンカーを作成する内容に差し替え
- v1.0.0（2026-11-06）：初版（段落のページ番号を「p.番号」の相互参照で挿入）

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
