# IdFitAnchoredImageHeight

[![Direct](https://img.shields.io/badge/Direct%20Link-IdFitAnchoredImageHeight.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/frame/IdFitAnchoredImageHeight.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdFitAnchoredImageHeight.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

インラインアンカーされた画像フレームの高さを、同じ段落の文字サイズに合わせて縦横比を保ったまま調整します。

## 主な機能

- GREP（`~a`）でアンカーオブジェクトを含む段落を検索
- 画像以外の最大文字サイズを基準に高さを決定
- `resize()` を使うため、フレームだけでなく中の画像も一緒に拡大縮小
- 処理範囲をドキュメント／ストーリー／選択範囲から選択（スクリプト冒頭の設定）

## 使い方

1. 対象のドキュメントを開く（ストーリー／選択範囲を対象にする場合はテキストを選択）
2. スクリプトを実行する

## 制限事項・メモ

- 画像のみの段落（前後に文字なし）は対象外です。
- 単位は処理中だけ mm に変更し、終了時に必ず元へ戻します。
- 処理範囲は `SEARCH_SCOPE` で変更できます（既定は `"document"`）。

## スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/frame/IdFitAnchoredImageHeight.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-06-01 |
| 最終更新 | 2026-06-01 |

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
