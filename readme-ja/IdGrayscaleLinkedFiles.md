# リンクした AI・EPS・PSD をモノクロにして差し替え

[![Direct](https://img.shields.io/badge/Direct%20Link-IdGrayscaleLinkedFiles.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdGrayscaleLinkedFiles.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdGrayscaleLinkedFiles.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

ドキュメントにリンクした AI・EPS は Illustrator で CMYK のモノクロ（テキストは K100）に、PSD は Photoshop でグレースケールにし、「_gray」を付けた別名で保存してリンクを差し替えます。[ドキュメント全体を CMYK のモノクロにする](IdConvertDocumentToGrayscale.md) と組み合わせると、配置したイラストや画像まで含めてモノクロにできます。

### 主な機能

1. ドキュメントのリンクから AI・EPS・PSD を集める（同じファイルを何か所に配置していても1回だけ処理する）
2. AI・EPS は BridgeTalk で Illustrator に送り、開いて次の順に処理する（Illustrator 版 [ConvertDocumentToGrayscale](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertDocumentToGrayscale.md) と同じ）
   - ［編集］→［カラーを編集］→［CMYK に変換］、ドキュメントのカラーモードを CMYK に
   - 使っていないシンボルを削除し、使っているシンボルの中身をグレースケールに
   - ［編集］→［カラーを編集］→［グレースケールに変換］
   - テキストの塗りと線を K100 に
   - スウォッチを削除
3. PSD は BridgeTalk で Photoshop に送り、［イメージ］→［モード］→［グレースケール］にして、テキストレイヤーを黒（グレー 100%）にする。レイヤーは残す
4. 同じフォルダーに「名前_gray.ai」「名前_gray.eps」「名前_gray.psd」として保存して閉じる
5. InDesign のリンクを、保存したファイルに差し替える

Illustrator ではロック・非表示のレイヤーとオブジェクトも一時的に解除して処理します。Photoshop ではロックしたテキストレイヤーも黒にします。

### 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する
3. 対象のファイル数を確かめて［はい］をクリックする

Illustrator・Photoshop は、使うときに起動していなければ起動します。実行中は、処理中のファイル名と進行状況バーを表示します。終わると、変換したファイル数と差し替えたリンク数、変換できなかったファイルを表示します。

### 制限事項・メモ

- 元のファイルは変更しません。同じ名前の「_gray」のファイルがすでにあるときは上書きします。
- 名前が「_gray」で終わるリンク（変換済み）、埋め込んだリンク、見つからないリンクは対象外です。
- Illustrator・Photoshop で開いたままのファイルは、編集中の内容を守るため処理しません。閉じてから実行し直してください。
- AI は PDF 互換ファイルを含めて保存します。EPS は Illustrator の EPS の既定の設定で保存します。PSD はレイヤー・アルファチャンネル・特色・カラープロファイルを含めて保存します。
- 白い文字（グレー 0%）は、AI・EPS・PSD とも白のまま残します（スクリプト冒頭の KEEP_WHITE_TEXT を false にすると白も K100）。
- シンボルの基準点・ダイナミックシンボルの種類・9スライスなどのシンボルオプションは引き継がれません。リンク画像、パターン・ブラシの中身、グラフィックスタイルは変換されません。
- PSD のテキストレイヤーは、レイヤー全体を黒にします（1つのレイヤー内で色を分けていても黒1色になります）。調整レイヤー・スマートオブジェクトはそのまま残し、グレースケールへの変換時の確認は表示しません。
- 複数の版を起動しているときは、最新版に送ります。版を固定するときは、スクリプト冒頭の ILLUSTRATOR_TARGET・PHOTOSHOP_TARGET を "illustrator-31.0" などにします。
- 1ファイルの処理は最大 300 秒待ちます（FILE_TIMEOUT_SECONDS）。
- リンクの差し替えは1回の取り消しで戻せます。保存したファイルは残ります。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/document/IdGrayscaleLinkedFiles.jsx` |
| バージョン | v1.0.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-11-21 |
| 最終更新 | 2026-11-21 |

### 更新履歴

- v1.0.0（2026-11-21）：初版

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
