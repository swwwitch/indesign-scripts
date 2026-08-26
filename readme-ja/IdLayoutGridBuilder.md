# IdLayoutGridBuilder

[![Direct](https://img.shields.io/badge/Direct%20Link-IdLayoutGridBuilder.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/IdLayoutGridBuilder.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdLayoutGridBuilder.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

アクティブページに版面・タイトルエリア・フレーム・列行グリッド・区切り線などを、プレビューを見ながら一括作成します。

## 主な機能

- 版面（罫線・角丸・伸縮・線端）、タイトルエリア、フッターのコラムエリアを設定
- ページのマージン、フレーム、実コンテンツ領域のオフセットを個別に指定
- 列数・行数・間隔からグリッドを作成し、区切り線（実線／破線／ドット点線）を描画
- 基本テキストのサイズ・行送りから 1 行の文字数を自動計算
- プレビューレイヤー上に描画し、確認しながら調整

## 使い方

1. 対象のページをアクティブにする
2. スクリプトを実行する
3. 各パネルで設定し、プレビューを確認して［OK］

## 制限事項・メモ

- 外側エリアと実コンテンツ領域は分離して扱い、タイトルエリアとカラムエリア（＋アキ）を実コンテンツ領域から差し引きます。
- 見開きに対応するため、自動調整系で使うページ境界はスプレッド座標に統一しています。
- 破線・点線は itemByName("破線 (3 & 2)" / "点線 (1 & 1)") で取得します。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/IdLayoutGridBuilder.jsx` |
| バージョン | v0.2.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-03-13 |
| 最終更新 | 2026-03-15 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
