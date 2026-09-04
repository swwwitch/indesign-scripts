# 段落ごとに独立したテキストフレームへ分割

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSplitParagraph.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdSplitParagraph.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSplitParagraph.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択したテキストフレーム内の各段落を、元の位置と幅（縦組みは高さ）を保ったまま独立したテキストフレームへ分割します。縦組みにも対応しています。

## 主な機能

- 元段落のベースライン位置に合わせて配置（横組みは上下、縦組みは左右）
- 元フレームの幅（縦組みは高さ）を維持
- テキストフレーム設定（textFramePreferences）を引き継ぎ
- 縦組みテキストフレームにも対応（オーバーセット解消の拡張は左方向）
- オーバーセットテキストがある場合は「拡張して解消」「そのまま実行」を選択可能

## 使い方

1. 分割したいテキストフレームを 1 つだけ選択する
2. スクリプトを実行する
3. オーバーセットがある場合は処理方法を選んで［OK］

## 制限事項・メモ

- 連結（スレッド）されたテキストフレーム、アンカー付きフレーム、ほかのオブジェクトの内側にあるフレームには対応していません（実行時にメッセージを表示して中止します）。
- 回転したフレームは想定していません（回転前の外接矩形で分割されます）。
- 空段落（空行）は出力対象から除外されます。分割できる段落が 1 つもない場合は、元のフレームをそのまま残します。
- 「そのまま実行」を選ぶと、あふれて隠れているテキストは出力されません。
- オーバーセットを解消できなかった場合は、フレームを元の大きさに戻して中止します。
- 一連の処理は 1 回の取り消しでまとめて元に戻せます。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/text/IdSplitParagraph.jsx` |
| バージョン | v1.1.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-03-16 |
| 最終更新 | 2026-09-04 |
| 紹介記事 | https://note.com/dtp_tranist/n/n8793ea71526b |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
