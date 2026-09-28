# Change Log

## 20260929

### 調整

- README を追加：[目印の文字列を削除して段落スタイルを適用（シンプル版）](readme-ja/IdRemoveMarkerApplyStyleSimple.md)、[Word の段落スタイル名を HTML 要素名に変換](readme-ja/IdConvertWordStyleNamesToHtml.md)
- JSX の概要に README と note 記事の URL を追加（47本）
- README の見出しを `###` にそろえ、英語版の節名（Update History / Script info）を統一（全本）

## 20260927

### 調整

- 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ、option＋で0.1ずつ）（13本）
  - [現在の親ページを引き継いでページを挿入](readme-ja/IdAddPagesUsingCurrentMaster.md)（v1.3.0）
  - [選択テキストのサイズでグラフィックフレームを作成](readme-ja/IdCreateFrameFromSelectedText.md)（v2.7.0）
  - [画像の縮尺率を変更してフレームを合わせる](readme-ja/IdSetImageScale.md)（v1.4.0）
  - [版面・グリッド・区切り線をプレビュー付きで作成](readme-ja/IdLayoutGridBuilder.md)（v1.1.0）
  - [スクリプトをキーワードで絞り込んで実行](readme-ja/IdScriptLauncher.md)（v1.2.0）
  - [同一スタイル間の段落間隔をスタイル定義に設定](readme-ja/IdSetSameParaStyleSpacing.md)（v1.1.0）
  - [タイプスケールから段落スタイルを一括適用](readme-ja/IdTypeScaleStyleApplier.md)（v1.7.0）
  - [段落スタイルの文字組版設定をまとめて設定](readme-ja/IdTypesettingStyleManager.md)（v1.2.0）：読み込み時にエラーになり起動しなかった不具合も修正
  - [表の罫線をプレビュー付きで描画・消去](readme-ja/IdSmartBorderBuilder.md)（v1.7.0）
  - [表の列幅をまとめて調整](readme-ja/IdTableColumnWidthAdjuster.md)（v1.2.0）
  - [表の行の高さをプレビュー付きで設定](readme-ja/IdTableRowHeightManager.md)（v1.4.0）
  - [表全体の幅と列幅をまとめて調整](readme-ja/IdTableWidthColumnWidthManager.md)（v1.2.0）
  - [表の行に交互の塗り（縞模様）を適用](readme-ja/IdZebraRowFill.md)（v1.3.0）

## 20260926

### 新しいスクリプトを追加

- [画像の縮尺率を変更してフレームを合わせる](readme-ja/IdSetImageScale.md)

## 20260925

### 新しいスクリプトを追加

- [メニューアクションを調べて実行コードをコピー](readme-ja/IdMenuActionsViewer.md)
- [アンカー付きオブジェクトをまとめて解除](readme-ja/IdReleaseAnchoredObjects.md)

### 調整

- [同じ段落スタイルで繰り返すテキストに連番を付ける](readme-ja/IdAppendParagraphNumbering.md)（v1.2.1）：ダイアログのタイトルを「繰り返し段落に連番を追加」に、ボタン名を［番号を追加］［番号を削除］に、パネル名「対象」を「範囲」に変更。「括弧：」の項目名とツールチップを追加。メッセージの文言を実際の動作に合わせて修正
- [ダイアログでテキストを編集して置換・挿入](readme-ja/IdEditTextsByDialog.md)（v0.1.4）：ダイアログのタイトルを「テキストを編集」に、ボタン名を［改行を削除］［@# を追加］に変更。ツールチップを追加。ドキュメントが開いていないときはメッセージを出して終了。1 文字だけの特殊文字を選択して実行するとエラーになる問題を修正
- [版面・グリッド・区切り線をプレビュー付きで作成](readme-ja/IdLayoutGridBuilder.md)（v1.0.0）：ダイアログを［ページ］［額縁・エリア］［実コンテンツ領域］の3つのタブに変更し、タイトルを「版面とグリッドの作成」に。［ページのマージンにも反映］、額縁の［角丸］、マージンの［連動］を追加。［単位］の先頭にドキュメントの定規の単位を表示。線幅・文字サイズ・行送りの単位換算、サンプル文のフォント、項目の有効／無効の連動などの不具合を修正
- [近接するオブジェクトを行・列単位でグループ化](readme-ja/IdSmartGroup.md)（v1.1.0）：ダイアログ名を「スマートグループ」に変更。［プレビューを表示］を追加し、プレビューを最上位のレイヤーに赤い枠線で描くように変更。「許容値」に定規の単位を表示。グループ化できる並びがないときはアラートを表示。終了時にプレビュー用スウォッチを削除し、アクティブレイヤーを元に戻すように変更
- [表の行の高さをプレビュー付きで設定](readme-ja/IdTableRowHeightManager.md)（v1.3.2）：ダイアログを2列構成に変更。［親フレームの調整］ボタンを［フレームを内容に合わせる］チェックボックスに変更（プレビューにも反映）。↑↓キーで0.1ずつ増減したときに端数の誤差が出ないように修正。ツールチップを見直し
- IdGroupHorizontal.jsx を削除（横並びのグループ化は [近接するオブジェクトを行・列単位でグループ化](readme-ja/IdSmartGroup.md) でまかなえるため）
- IdScriptLauncher.jsx を jsx/runner/ へ移動

## 20260924

### 新しいスクリプトを追加

- [使われていないスタイルとスウォッチを削除](readme-ja/IdDeleteUnused.md)

## 20260920

### 新しいスクリプトを追加

- [テキスト・表・オブジェクトのスタイルオーバーライドを一括消去](readme-ja/IdClearStyleOverrides.md)
- [Word の段落スタイル名を HTML 要素名に変換](readme-ja/IdConvertWordStyleNamesToHtml.md)
- [ドキュメントで使用中のフォントを一括置換](readme-ja/IdReplaceDocumentFonts.md)
- [複数のテキストフレームを連結](readme-ja/IdTextFrameLinker.md)

### 調整

- [段落スタイル・文字スタイルを一括登録](readme-ja/IdStyleSetup.md)（v1.4.0）：実行時にダイアログを表示し、スタイル名の体系を［HTML］と［Word 対応］から選べるように変更。［Word 対応］では `Heading 1`〜`Heading 6`、`Normal` などの名前で登録
- [スクリプトをキーワードで絞り込んで実行](readme-ja/IdScriptLauncher.md)（v1.1.0）：件数に関わらず常に出すキーワードボタンを追加（font）。「検索条件を記憶」を追加（ONのあいだ、InDesign のセッション中は前回のキーワードやリストの選択を引き継ぐ）
- [タイプスケールから段落スタイルを一括適用](readme-ja/IdTypeScaleStyleApplier.md)（v1.6.1）：［サイズのみ］の既定をオフに変更。［フォント、スタイルを含める］をチェックボックスにし、［フォント指定オプション］パネルへ移動。フォントファミリーを指定しないままウエイトだけを変更できるように修正。段落前後のアキが mm のドキュメントで約 2.83 倍になっていた問題、文字サイズの単位が pt 以外のとき基準サイズの初期値がずれていた問題を修正
- IdTableBorderFill.jsx を削除

## 20260917

### 調整

- IdAutoParagraphStyleGeneratorV2.jsx を削除（[書式の組み合わせから段落スタイルを自動生成](readme-ja/IdAutoParagraphStyleGenerator.md) と処理内容が同じため）
- IdNestedStyleSetup.jsx を削除（[段落スタイルの正規表現スタイルを適用・管理](readme-ja/IdGrepStyleApplier.md) と同じ内容のため）

## 20260916

### 調整

- [表の結合セルを解除](readme-ja/IdCellUnmerge.md)（v1.0.2）：UI の文言を実際の動作に合わせて変更（「すべてのセルに複製」「元のセルにのみ残す」「解除後のテキスト」）。ダイアログのタイトルと取り消し名を「セルの結合を解除」に変更。表の一部を選択して実行したときは［対象］が「選択セルのみ」で開くように変更。ツールチップを追加
- [表の行に交互の塗り（縞模様）を適用](readme-ja/IdZebraRowFill.md)（v1.2.1）：ダイアログ名を「行の塗りを交互に設定」に、パネル名を「奇数番目の行」「偶数番目の行」に変更。［奇数行と偶数行を入れ替え］と行・列の［スキップ］をパネルにまとめ直した。ツールチップを追加。ボタン行を左右に分割

## 20260915

### 調整

- [表の罫線をプレビュー付きで描画・消去](readme-ja/IdSmartBorderBuilder.md)（v1.6.8）：項目名を変更（「外枠のみ」「内側のみ」「下端のみ」「右端のみ」「既存の罫線を消してから引く」、パネル「線の設定」など）。ツールチップに説明を追加。濃淡欄を↑↓キーで1ずつ、Shift＋↑↓キーで10ずつ増減するように変更

## 20260912

### 調整

- [ファイル名をセグメント単位で編集してリネーム・保存](readme-ja/IdFileNameManager.md)（v1.3.2）：ダイアログを開けなかった不具合を修正。「構成要素の順序」のラジオボタンが排他になっていなかった不具合、「現在のファイル名に準じる」が常に有効になっていた不具合を修正。大文字小文字だけ・濁点の合成違いだけのリネームで、保存したてのファイルをゴミ箱へ送ってしまう事故を修正。.indd 以外の書類ではリネーム・コピーを選べないように変更。v 番号・連番の認識を区切り単位に変更。保存先をダイアログ表示前に確定するよう変更

## 20260910

### 新しいスクリプトを追加

- [目印の文字列を削除して段落スタイルを適用（シンプル版）](readme-ja/IdRemoveMarkerApplyStyleSimple.md)

### 調整

- [Markdown の目印を削除してスタイルを適用](readme-ja/IdRemoveMarkerApplyStyle.md)（v1.1.1）：目印に続く半角／全角スペース・タブも一緒に削除。同じ文字が続く並びの一部（`##` と `### 見出し`）に一致しないように変更。「対象箇所」の件数表示、「正規表現（GREP）で検索」を追加。検索対象をラジオボタンのパネルに分けた。脚注とマスターページを検索対象から外した。検索に失敗したときは検索条件を残さずに中止

## 20260909

### 新しいスクリプトを追加

- [Markdown の目印を削除してスタイルを適用](readme-ja/IdRemoveMarkerApplyStyle.md)

### 調整

- [画像入りフレームを順送りに入れ替え](readme-ja/IdSwapImageFrames.md)（v1.0.1）：ダイアログの文言を見直し（「順送り」であることが分かる表現に、フィットの選択肢は InDesign 標準の用語に）。エラーメッセージに取り消しで戻せることを明記。ツールチップを追加。配置画像そのものを選択したときのフレーム判定を簡素化

## 20260907

### 調整

- [表の行と列を入れ替え](readme-ja/IdTransposeTableRowsCols.md)（v1.0.1）：セルを部分的に選択したときも表全体を対象にするよう修正。フッター行を持つ表で結果がずれる可能性があったのを修正。ダイアログの文言を実際の動作に合わせて変更。ツールチップを追加。セル結合がない表ではセル結合のパネルを無効表示に

## 20260904

### 調整

- [段落ごとに独立したテキストフレームへ分割](readme-ja/IdSplitParagraph.md)（v1.1.1）：縦組みのテキストフレームに対応。連結されたフレーム、アンカー付きフレームやほかのオブジェクトの内側のフレームは、アラートを出して処理しないように変更。分割できる段落がないときは元のフレームを残してアラートを表示。オーバーセットを解消できずに中断したときは、元のフレームの大きさを戻すように修正

## 20260901

### 調整

- [段落スタイル・文字スタイルを一括登録](readme-ja/IdStyleSetup.md)（v1.3.3）：表組み用の段落スタイル `p.table` を追加。`p` の分離禁止などの設定が効かないことがある問題、`body-text` がないと見出しの継承関係が設定されない問題、日本語版以外の InDesign でカーニング方式の設定に失敗して全体が取り消される問題を修正
- IdTypescale.jsx を削除（[タイプスケールから段落スタイルを一括適用](readme-ja/IdTypeScaleStyleApplier.md) に一本化）

## 20260827

### 新しいスクリプトを追加

- [正規表現スタイルの設定を一覧・書き出し](readme-ja/IdInspectGrepStyle.md)

### 調整

- スクリプトのファイル名を「Id」始まりにそろえ、カテゴリー別のフォルダー（document / font / frame / group / page / runner / style / table / text）に整理
- [内容が同じセルを自動で結合](readme-ja/IdCellMergeAuto.md)（v1.0.1）：［選択セルのみ］が使えなかった問題を修正。各パネルにツールチップを追加。ファイル名を `MergeCell-Auto.jsx` から変更
- [表の結合セルを解除](readme-ja/IdCellUnmerge.md)（v1.0.1）：［選択したセルのみ］で結合が解除されなかったり、別のセルのテキストが混ざって入る問題を修正。「テキストを分配」の対象を、その結合セルの解除で生じたセルだけに限定。［表全体］で結合セルを取りこぼすことがあった問題を修正
- [スクリプトをキーワードで絞り込んで実行](readme-ja/IdScriptLauncher.md)（v1.0.1）：キーワード欄にクリア（×）ボタンを追加

## 20260826

### 新しいスクリプトを追加

- [スクリプトをキーワードで絞り込んで実行](readme-ja/IdScriptLauncher.md)

## 20260814

### 調整

- [現在の親ページを引き継いでページを挿入](readme-ja/IdAddPagesUsingCurrentMaster.md)（v1.2.2）：挿入ページ数の欄で全角数字を使えるように変更。数字以外を含む入力はエラーに。レイアウトウィンドウがアクティブでないときや、親ページを表示しているときはアラートを出して中止。現在のページの親ページが［なし］のとき、挿入したページも［なし］にそろえるように修正

## 20260813

### 調整

- [同じ段落スタイルで繰り返すテキストに連番を付ける](readme-ja/IdAppendParagraphNumbering.md)（v1.2.0）：対象に［選択範囲］を追加。親見出しが見つからない段落が対象から漏れていた問題、空段落や1文字の見出しがあると処理されないことがある問題、日本語 UI で［キャンセル］がダイアログを閉じない問題を修正。［OK］を［追加］に変更し、ボタンを左（削除）／右（キャンセル・追加）に配置。解析処理を高速化

## 20260719

### 新しいスクリプトを追加

- [文字の水平・垂直比率を 100% に戻す](readme-ja/IdResetHorizontalVerticalScale.md)

## 20260709

### 新しいスクリプトを追加

- [行送りを自動行送り（％指定）に切り替える](readme-ja/IdAutoLeadingCalc.md)

## 20260705

### 調整

- [カーソル位置から段落末尾までを削除](readme-ja/IdDeleteFromCursorToEnd.md)（v1.2.1）：段落末尾の記号（。！？,.、，．）の直前にカーソルがあるときは、記号1文字だけを削除するように変更

## 20260630

### 新しいスクリプトを追加

- [同一スタイル間の段落間隔をスタイル定義に設定](readme-ja/IdSetSameParaStyleSpacing.md)
- [段落ごとに独立したテキストフレームへ分割](readme-ja/IdSplitParagraph.md)

### 調整

- [現在の親ページを引き継いでページを挿入](readme-ja/IdAddPagesUsingCurrentMaster.md)（v1.2.1）：数値が不正なときのアラートを表示言語に合わせて1つだけ出すように変更。ESC キーでダイアログを閉じられるように。jsx/page/ へ移動
- [同じ段落スタイルで繰り返すテキストに連番を付ける](readme-ja/IdAppendParagraphNumbering.md)（v1.1.2）：対象のラジオボタンが無視される不具合を修正。［削除］を1回の取り消しで戻せるように。jsx/page/ へ移動
- [段落スタイル・文字スタイルを一括登録](readme-ja/IdStyleSetup.md)（v1.3.1）：処理中にプログレスバーを表示。ul-li の「同じスタイルが連続する段落間のスペース」を0に設定。プログレスバー・アラート・取り消しの名前を英語表示に対応

## 20260627

### 新しいスクリプトを追加

- [カーソル位置から段落末尾までを削除](readme-ja/IdDeleteFromCursorToEnd.md)
- [段落の開始位置を「次の段」に設定](readme-ja/IdKeepOptionNextColumn.md)

## 20260617

### 新しいスクリプトを追加

- [フォントの種別・ウエイトをまとめて切り替え](readme-ja/IdFontConverter.md)

## 20260615

### 調整

- [段落スタイル・文字スタイルを一括登録](readme-ja/IdStyleSetup.md)（v1.3.0）：ul-li に正規表現スタイルが引き継がれていなかった問題を修正。既存の同名スタイルを「置き換える」設定を追加（オンにすると全属性と正規表現スタイルを付け直す）。基準スタイルは常に設定するように変更

## 20260614

### 調整

- [段落スタイル・文字スタイルを一括登録](readme-ja/IdStyleSetup.md)（v1.2.5）：basestyle グループ（body-text / heading / base-regex / base-toc）を追加し、本文系・見出し系の継承を再編。見出しに段落分離禁止と「次のスタイル」を設定。リスト・キャプション・コード・表セルの分離禁止などを設定。全体を1回の取り消しで戻せるように。「すべての行を分離禁止」が効かない問題を修正
- [タイプスケールから段落スタイルを一括適用](readme-ja/IdTypeScaleStyleApplier.md)（v1.6.0）：プレビューに「リスト」「テーブル」の行を追加。基準サイズの初期値を段落スタイル「p」→「Normal」の本文サイズから取るように変更。フォント情報は［フォント、スタイルを含める］で読み込むように変更。キャンセルしたときはプレビューで変わった段落スタイルを元に戻し、OK の適用は1回の取り消しで戻せるように変更

## 20260609

### 調整

- [表の行の高さをプレビュー付きで設定](readme-ja/IdTableRowHeightManager.md)（v1.3.1）：［範囲］（ドキュメント／ストーリー／選択範囲）で対象の表が実際に切り替わるように修正

## 20260602

### 新しいスクリプトを追加

- [アンカー付きグラフィックフレームの幅・サイズ・縮尺を一括調整](readme-ja/IdAdjustGraphicFrames.md)

## 20260601

### 新しいスクリプトを追加

- [アンカー画像の高さを文字サイズに合わせる](readme-ja/IdFitAnchoredImageHeight.md)

## 20260529

### 調整

- [ファイル名をセグメント単位で編集してリネーム・保存](readme-ja/IdFileNameManager.md)（v1.3.1）：「半角カナ → 全角」、丸数字や法人略記の「削除する」、連続する区切り記号の圧縮、ファイル名長の事前チェック、Windows 予約名の回避、「時刻も付与」、新セグメント「ページ番号」を追加。リネーム時の元ファイルはゴミ箱へ移動。「バージョンのみ」では整形をかけず v 番号だけ更新

## 20260528

### 調整

- [ファイル名をセグメント単位で編集してリネーム・保存](readme-ja/IdFileNameManager.md)（v1.2.5）：並び順「現在のファイル名に準じる」、サブテキストの「2 階層上のフォルダー」、「クリーンなファイル名」、「丸数字や法人略記など」の変換を追加。バージョン番号の表記を v1 / v01 / v001 の3段階に拡張し、フォルダー内の最大値 +1 を採用。濁点・半濁点の正規化を追加

## 20260527

### 新しいスクリプトを追加

- [ファイル名をセグメント単位で編集してリネーム・保存](readme-ja/IdFileNameManager.md)

## 20260513

### 調整

- [段落スタイル・文字スタイルを一括登録](readme-ja/IdStyleSetup.md)（v1.1.0）：文字スタイル「code-normal」「highlighter」「sumaru」と段落スタイル「toc-title」を追加。基準スタイルの継承を追加。p に正規表現スタイル（文末の2文字＋句読点→sumaru、アンカー付きオブジェクト→inline-graphic）を追加。パネル上の並び順を定義順にそろえるように変更

## 20260507

### 調整

- [段落スタイルの文字組版設定をまとめて設定](readme-ja/IdTypesettingStyleManager.md)（v1.1.1）：プリセットの初期値を見直し（「欧文組版」「グリッド優先」「グリッド無視」）。パネル名「単位」を「単位（環境設定）」に変更

## 20260506

### 新しいスクリプトを追加

- [段落スタイルの文字組版設定をまとめて設定](readme-ja/IdTypesettingStyleManager.md)

### 調整

- [日本語組版設定をマトリックスで一括適用](readme-ja/IdJapaneseParagraphTypesettingManager.md)（v1.2.0）：コンポーザーを日本語版・英語版のどちらの名前でも読み取り・適用できるように修正。欧文のコンポーザーの表示名を「欧文段落コンポーザー」「欧文単数行コンポーザー」に変更

## 20260505

### 新しいスクリプトを追加

- [タイプスケールから段落スタイルを一括適用](readme-ja/IdTypeScaleStyleApplier.md)
- [日本語組版設定をマトリックスで一括適用](readme-ja/IdJapaneseParagraphTypesettingManager.md)

## 20260503

### 新しいスクリプトを追加

- [段落スタイル・文字スタイルを一括登録](readme-ja/IdStyleSetup.md)
- [段落スタイルの正規表現スタイルを適用・管理](readme-ja/IdGrepStyleApplier.md)

## 20260420

### 新しいスクリプトを追加

- [表の行の高さをプレビュー付きで設定](readme-ja/IdTableRowHeightManager.md)

## 20260419

### 新しいスクリプトを追加

- [表の列幅をまとめて調整](readme-ja/IdTableColumnWidthAdjuster.md)

## 20260418

### 新しいスクリプトを追加

- [表全体の幅と列幅をまとめて調整](readme-ja/IdTableWidthColumnWidthManager.md)

## 20260417

### 新しいスクリプトを追加

- [内容が同じセルを自動で結合](readme-ja/IdCellMergeAuto.md)
- [表の結合セルを解除](readme-ja/IdCellUnmerge.md)
- [最終列を伸縮させて表幅をフレーム幅に合わせる](readme-ja/IdSetLastColumnToFrameWidth.md)
- [表幅を親テキストフレームの幅にそろえる](readme-ja/IdTableColumnEqualizer.md)
- [表の行に交互の塗り（縞模様）を適用](readme-ja/IdZebraRowFill.md)
- [ExtendScript ファイルを選んで実行](readme-ja/IdScriptRunner.md)
- [ExtendScript ファイルを選んで実行（最小構成）](readme-ja/IdScriptRunnerSimple.md)

### 調整

- [表の罫線をプレビュー付きで描画・消去](readme-ja/IdSmartBorderBuilder.md)（v1.6.7）：モード「最下辺のみ」「最右辺のみ」を追加。OK で確定した設定を覚えておき、次に開いたときに復元するように変更。［描画前に消去］の初期値を OFF に変更

## 20260413

### 調整

- [表の罫線をプレビュー付きで描画・消去](readme-ja/IdSmartBorderBuilder.md)（v1.6.5）：罫線専用のダイアログに戻し、［塗り］タブを廃止。カラーに濃淡を追加。モードを Option＋クリックすると［描画前に消去］が切り替わるように。「見出し行」「見出し列」は表全体を選んだときだけ使えるように変更。表の一部を選んだときは選択範囲の矩形からセルを組み直して処理。終了後は実行前の選択状態に戻す

## 20260412

### 新しいスクリプトを追加

- [段落背景色の高さに合わせた境界線スペーサーを設定](readme-ja/IdParagraphShadingMatchRule.md)

## 20260411

### 新しいスクリプトを追加

- [表の罫線をプレビュー付きで描画・消去](readme-ja/IdSmartBorderBuilder.md)
- [近接するオブジェクトを行・列単位でグループ化](readme-ja/IdSmartGroup.md)

## 20260328

### 新しいスクリプトを追加

- [画像入りフレームを順送りに入れ替え](readme-ja/IdSwapImageFrames.md)

## 20260317

### 新しいスクリプトを追加

- [選択テキストのサイズでグラフィックフレームを作成](readme-ja/IdCreateFrameFromSelectedText.md)
- [Markdown 記法を検索・置換してスタイルを適用](readme-ja/IdFindChangeByListMarkdown.md)

## 20260315

### 調整

- [版面・グリッド・区切り線をプレビュー付きで作成](readme-ja/IdLayoutGridBuilder.md)（v0.2.0）：区切り線の線種「点線」を「破線」に名称変更し、破線は「破線 (3 & 2)」、ドット点線は「点線 (1 & 1)」の線種で描くように変更。該当する線種がドキュメントにないときは警告を出して実線で描画

## 20260314

### 調整

- [版面・グリッド・区切り線をプレビュー付きで作成](readme-ja/IdLayoutGridBuilder.md)（v0.1.4）：見開きのページで、自動調整に使うページ境界がドキュメントの定規単位とずれる問題を修正。列幅の文字数の自動計算もドキュメントの単位で求めるよう修正
- [書式の組み合わせから段落スタイルを自動生成](readme-ja/IdAutoParagraphStyleGenerator.md)（v3.4）：スタイル名のサイズ表記を環境設定のテキストサイズ単位（pt／Q）に合わせるように変更。サイズと行送りを pt に換算してから比べ、単位換算の誤差で別のスタイルに分かれにくくした

## 20260313

### 新しいスクリプトを追加

- [版面・グリッド・区切り線をプレビュー付きで作成](readme-ja/IdLayoutGridBuilder.md)

## 20260213

### 新しいスクリプトを追加

- [書式の組み合わせから段落スタイルを自動生成](readme-ja/IdAutoParagraphStyleGenerator.md)

## 20251125

### 新しいスクリプトを追加

- [表の行と列を入れ替え](readme-ja/IdSwapTableRowColumn.md)
- [表の行と列を入れ替え](readme-ja/IdTransposeTableRowsCols.md)

## 20250702

### 新しいスクリプトを追加

- [親ページとドキュメントページを切り替え](readme-ja/IdSwitchToMasterOrDocument.md)

### 調整

- [同じ段落スタイルで繰り返すテキストに連番を付ける](readme-ja/IdAppendParagraphNumbering.md)（v1.1.0）：段落スタイルをリストアップし、チェックボックスを外したら対象リストから除外（ディム表示）。［番号を削除］ボタンを追加（クリックすると末尾の番号を削除）。p.img・p.table スタイルの段落を無視。階層判定ロジックを改良

## 20250630

### 新しいスクリプトを追加

- [同じ段落スタイルで繰り返すテキストに連番を付ける](readme-ja/IdAppendParagraphNumbering.md)

## 20250626

### 新しいスクリプトを追加

- [現在の親ページを引き継いでページを挿入](readme-ja/IdAddPagesUsingCurrentMaster.md)

## 20250528

### 新しいスクリプトを追加

- [ダイアログでテキストを編集して置換・挿入](readme-ja/IdEditTextsByDialog.md)
