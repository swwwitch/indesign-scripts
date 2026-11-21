# ドキュメント全体を CMYK のモノクロにする（テキストは K100）

[![Direct](https://img.shields.io/badge/Direct%20Link-IdConvertDocumentToGrayscale.jsx-ffcc00.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdConvertDocumentToGrayscale.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdConvertDocumentToGrayscale.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/indesign-scripts/blob/main/README.md)

---

ドキュメントのカラーをすべて CMYK の K だけのグレーに変換し、テキストを K100（［黒］）にします。リンクした AI・EPS・PSD も、Illustrator・Photoshop でモノクロにして差し替えられます。モノクロ印刷用のデータ作りに使います。Illustrator 版（[ConvertDocumentToGrayscale](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertDocumentToGrayscale.md)）と同じ目的のスクリプトです。

### 主な機能

次の順に処理します。

1. 透明の分割・統合に使うカラースペース（［編集］→［透明ブレンド領域］）を CMYK にする
2. スウォッチのカラーと名前の無いカラーを、K だけの CMYK プロセスカラーに変換する（RGB・Lab・特色も対象。特色はプロセスカラーにする）
3. 段落スタイル・文字スタイルと、テキストに直接付けた塗りと線を K100（［黒］）にする（表・脚注・親ページのテキストも対象）
4. 同じ濃さになったスウォッチを1つにまとめ、「C=0 M=0 Y=0 K=n」の名前にする（K100 は［黒］にまとめる）
5. 使われていないスウォッチを削除する

ロックしたレイヤーとオブジェクトも一時的に解除して処理し、終わったら元に戻します。

#### リンクしたファイル

リンクした AI・EPS・PSD があるときは、一緒に変換するか確認します。［はい］のときは、ドキュメントのカラーの前に次を行います（［いいえ］のときはドキュメントのカラーだけ）。

1. AI・EPS は BridgeTalk で Illustrator に送り、Illustrator 版 ConvertDocumentToGrayscale と同じ処理をする（［CMYK に変換］→ カラーモードを CMYK → 使っているシンボルの中身をグレースケールに・使っていないシンボルは削除 →［グレースケールに変換］→ テキストを K100 → スウォッチを削除）
2. PSD は BridgeTalk で Photoshop に送り、［イメージ］→［モード］→［グレースケール］にして、テキストレイヤーを黒（グレー 100%）にする。レイヤーは残す
3. 同じフォルダーに「名前_gray.ai」「名前_gray.eps」「名前_gray.psd」として保存して閉じる
4. リンクを保存したファイルに差し替える

実行中は、処理中のファイル名と進行状況バーを表示します。Illustrator・Photoshop は、使うときに起動していなければ起動します。終わると、変換したファイル数と差し替えたリンク数、変換できなかったファイルを表示します。

### 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する
3. リンクした AI・EPS・PSD があるときは、一緒に変換するなら［はい］、ドキュメントのカラーだけなら［いいえ］をクリックする

### 制限事項・メモ

- 選択の有無にかかわらず、ドキュメント全体が対象です。
- InDesign には［グレースケールに変換］が無いので、カラーを Lab に変換したときの明度（L*）が同じになる K の値をグレーの濃さにします（ドキュメントのカラー設定で換算）。もとから K だけのカラーは値を変えません。
- テキストで K100 にするのは、塗りと線の単色だけです。白（［紙色］・K0）は白のまま残します（スクリプト冒頭の KEEP_WHITE_TEXT を false にすると白も K100）。グラデーションの文字は、グラデーションのカラーがグレーになるだけです。
- リンクしたファイルで変換するのは AI・EPS・PSD だけです。PDF・TIFF・JPEG などの中身や、埋め込んだリンクは変換しません。
- リンクした元のファイルは変更しません。同じ名前の「_gray」のファイルがすでにあるときは上書きします。名前が「_gray」で終わるリンク（変換済み）、見つからないリンク、Illustrator・Photoshop で開いたままのファイルは対象外です。
- AI は PDF 互換ファイルを含めて保存します。EPS は Illustrator の EPS の既定の設定、PSD はレイヤー・アルファチャンネル・特色・カラープロファイルを含めて保存します。
- PSD のテキストレイヤーは、レイヤー全体を黒にします（1つのレイヤー内で色を分けていても黒1色になります）。白い文字は白のまま残します。
- AI・EPS のシンボルの基準点・ダイナミックシンボルの種類・9スライスなどのシンボルオプションは引き継がれません。
- 複数の版の Illustrator・Photoshop を起動しているときは最新版に送ります。版を固定するときは、スクリプト冒頭の ILLUSTRATOR_TARGET・PHOTOSHOP_TARGET を "illustrator-31.0" などにします。1ファイルの処理は最大 300 秒待ちます（FILE_TIMEOUT_SECONDS）。
- 混合インキは変換しません。
- グラデーションの分岐点と濃淡の元のカラーは、未使用でも削除しません。
- ［検索と置換］の検索条件は空になります（検索範囲の設定は元に戻します）。
- リンクの差し替えとドキュメントのカラーの変換は、1回の取り消しで戻せます。保存したファイルは残ります。

### スクリプト情報

| 項目 | 内容 |
| --- | --- |
| ファイル | `jsx/document/IdConvertDocumentToGrayscale.jsx` |
| バージョン | v1.1.0 |
| 作者 | Masahiro Takano (@swwwitch) |
| 初回リリース | 2026-11-21 |
| 最終更新 | 2026-11-21 |

### 更新履歴

- v1.1.0（2026-11-21）：リンクした AI・EPS・PSD を Illustrator・Photoshop でモノクロにして差し替える機能を追加（IdGrayscaleLinkedFiles を統合）
- v1.0.0（2026-11-21）：初版

### ライセンス

MIT License — <http://opensource.org/licenses/mit-license.php>
