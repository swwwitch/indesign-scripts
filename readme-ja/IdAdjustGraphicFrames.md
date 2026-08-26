# アンカー付きグラフィックフレームの幅・サイズ・縮尺を一括調整

[![Direct](https://img.shields.io/badge/Direct%20Link-IdAdjustGraphicFrames.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/frame/IdAdjustGraphicFrames.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAdjustGraphicFrames.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

テキストにアンカーされたグラフィックフレームを集め、フレーム幅・フレームサイズ・画像の縮尺率をまとめて調整します。

## 主な機能

- 対象をドキュメント／ストーリー／選択範囲から選択
- フレーム幅を「変更しない／親フレームに合わせる／マージンに合わせる」から選択（「合わせ先より大きい場合のみ」の切り替え付き）
- フレームを内容に合わせる／インライン画像を文字サイズに合わせる
- 画像の縮尺率を 1%／5%／10% 単位で切り捨て（72/96/144 ppi 限定、切り捨て後の再フィットも指定可）
- 縦横比が崩れた画像は、横スケールを基準に常に補正

## 使い方

1. 対象のドキュメントを開く（ストーリー／選択範囲を対象にする場合はテキストやフレームを選択）
2. スクリプトを実行する
3. 対象・フレーム幅・フレームサイズ・縮尺率を指定して［OK］

## 制限事項・メモ

- 対象はテキストにアンカーされたグラフィックフレームだけです。独立配置のフレームとテキストフレームは処理しません。
- 前後に文字がある真のインライン画像は、文字サイズへの高さ合わせを優先し、縮尺率の切り捨てとフレーム幅の調整は行いません。
- 全処理が 1 回の取り消しにまとまります。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/frame/IdAdjustGraphicFrames.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-06-02 |
| 最終更新 | 2026-06-02 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
