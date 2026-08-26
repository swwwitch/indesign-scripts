# IdTypeScaleStyleApplier

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTypeScaleStyleApplier.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdTypeScaleStyleApplier.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTypeScaleStyleApplier.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

基準サイズとスケール倍率からタイプスケールを組み立て、本文・見出し・リスト・表の段落スタイルへ一括適用します。

## 主な機能

- リストは本文比 100%、表は 94% を既定として算出
- 本文・見出しごとの行送り（%）、カーニング方式、サイズの丸め単位を指定
- 段落前後のアキは既定の％から導出し、プレビューで行ごとに上書き可能（サイズも同様）
- フォントは「共通／別々／変更しない」。見出しは本文フォントを参照してウエイトだけ変えることも可能
- フォント情報は起動時には読み込まず、［フォントも含める］を押したときだけ読み込む（一覧はディスクにキャッシュ）

## 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する
3. 基準サイズと倍率を指定し、プレビューで各行を調整して［OK］

## 制限事項・メモ

- 「同じスタイルの段落間隔」はスタイル名ごとにルールが変わります（ul-li は 0、p / ol-li は段落前のアキと同値、それ以外は変更しない）。
- 字揃えの強制適用は既定で OFF です（`ENABLE_JUSTIFICATION`）。
- 全処理が 1 回の取り消しにまとまります。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/style/IdTypeScaleStyleApplier.jsx` |
| バージョン | v1.6.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-05-05 |
| 最終更新 | 2026-06-30 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
