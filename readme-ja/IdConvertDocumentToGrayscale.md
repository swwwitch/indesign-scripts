# ドキュメント全体を CMYK のモノクロにする（テキストは K100）

[![Direct](https://img.shields.io/badge/Direct%20Link-IdConvertDocumentToGrayscale.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdConvertDocumentToGrayscale.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdConvertDocumentToGrayscale.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

ドキュメントのカラーをすべて CMYK の K だけのグレーに変換し、テキストを K100（［黒］）にします。モノクロ印刷用のデータ作りに使います。Illustrator 版（[ConvertDocumentToGrayscale](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertDocumentToGrayscale.md)）と同じ目的のスクリプトです。

### 主な機能

次の順に処理します。

1. 透明の分割・統合に使うカラースペース（［編集］→［透明ブレンド領域］）を CMYK にする
2. スウォッチのカラーと名前の無いカラーを、K だけの CMYK プロセスカラーに変換する（RGB・Lab・特色も対象。特色はプロセスカラーにする）
3. 段落スタイル・文字スタイルと、テキストに直接付けた塗りと線を K100（［黒］）にする（表・脚注・親ページのテキストも対象）
4. 同じ濃さになったスウォッチを1つにまとめ、「C=0 M=0 Y=0 K=n」の名前にする（K100 は［黒］にまとめる）
5. 使われていないスウォッチを削除する

ロックしたレイヤーとオブジェクトも一時的に解除して処理し、終わったら元に戻します。

### 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する

### 制限事項・メモ

- ダイアログボックスはありません。選択の有無にかかわらず、ドキュメント全体が対象です。
- InDesign には［グレースケールに変換］が無いので、カラーを Lab に変換したときの明度（L*）が同じになる K の値をグレーの濃さにします（ドキュメントのカラー設定で換算）。もとから K だけのカラーは値を変えません。
- テキストで K100 にするのは、塗りと線の単色だけです。白（［紙色］・K0）は白のまま残します（スクリプト冒頭の KEEP_WHITE_TEXT を false にすると白も K100）。グラデーションの文字は、グラデーションのカラーがグレーになるだけです。
- 配置した画像・グラフィック（PSD・AI・PDF など）の中身は変換しません。
- 混合インキは変換しません。
- グラデーションの分岐点と濃淡の元のカラーは、未使用でも削除しません。
- ［検索と置換］の検索条件は空になります（検索範囲の設定は元に戻します）。
- 1回の取り消しで戻せます。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/document/IdConvertDocumentToGrayscale.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-11-21 |
| 最終更新 | 2026-11-21 |

### 更新履歴

- v1.0.0（2026-11-21）：初版

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
