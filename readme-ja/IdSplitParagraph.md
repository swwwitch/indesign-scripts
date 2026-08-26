# 段落ごとに独立したテキストフレームへ分割

[![Direct](https://img.shields.io/badge/Direct%20Link-IdSplitParagraph.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdSplitParagraph.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSplitParagraph.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

選択したテキストフレーム内の各段落を、元の位置と幅を保ったまま独立したテキストフレームへ分割します。

## 主な機能

- 元段落のベースライン位置を基準に Y 位置を補正
- 元フレームの左右座標を維持し、幅を保持
- テキストフレーム設定（textFramePreferences）を引き継ぎ
- オーバーセットテキストがある場合は「拡張して解消」「そのまま実行」を選択可能

## 使い方

1. 分割したいテキストフレームを 1 つだけ選択する
2. スクリプトを実行する
3. オーバーセットがある場合は処理方法を選んで［OK］

## 制限事項・メモ

- 空段落（空行）は出力対象から除外されます。
- 「そのまま実行」を選ぶと、あふれて隠れているテキストは出力されません。
- 一連の処理は 1 回の取り消しでまとめて元に戻せます。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/text/IdSplitParagraph.jsx` |
| バージョン | v1.1.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-03-16 |
| 最終更新 | 2026-06-30 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
