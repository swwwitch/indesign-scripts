# タイプスケールから段落スタイルを一括適用（本文・見出し・リスト・表）

[![Direct](https://img.shields.io/badge/Direct%20Link-IdTypeScaleStyleApplier.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdTypeScaleStyleApplier.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdTypeScaleStyleApplier.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

基準サイズとスケール倍率からタイプスケールを組み立て、本文・見出し・リスト・表の段落スタイルへ一括適用します。

## 主な機能

- リストは本文比 100%、表は 94% を既定として算出
- 本文・見出しごとの行送り（%）、カーニング方式、サイズの丸め単位を指定
- 段落前後のアキは既定の％から導出し、プレビューで行ごとに上書き可能（サイズも同様）
- ［フォント、スタイルを含める］がオフのあいだはフォントを変更しない。オンにすると「本文と見出しで共通」「別々に指定」を選べる
- フォントファミリーを指定しなくても、ウエイトだけの変更が可能。候補は各段落スタイルが現在使っているフォントから作られる
- ［サイズのみ］をオンにすると、文字サイズ以外（フォント・行送り・アキ・カーニング）は元の値のまま
- フォント一覧は起動時に読み込み、ディスクにキャッシュする

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
| バージョン | v1.6.1 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-05-05 |
| 最終更新 | 2026-09-20 |
| 紹介記事 | https://note.com/dtp_tranist/n/n4f9b0666db66 |

## 更新履歴

### v1.6.1（2026-09-20）

- ［サイズのみ］の既定をオフに変更。これまでは初期状態のままだと行送りと段落前後のアキが適用されませんでした
- ［フォント、スタイルを含める］をボタンからチェックボックスへ変更し、［フォント指定オプション］パネルへ移動。オフの状態が従来の［フォントを変更しない］にあたるため、同ラジオボタンは削除しました
- フォントファミリーを指定しないまま、ウエイトだけを変更できるように修正。候補は各段落スタイルが現在使っているフォントから作ります
- 段落前後のアキがドキュメントの定規単位で解釈され、mm のドキュメントで指定の約 2.83 倍になっていた問題を修正
- 文字サイズの単位が pt 以外（Q など）のとき、基準サイズの初期値がずれていた問題を修正。単位はアプリケーションの環境設定ではなくドキュメントの設定を参照します

## ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
