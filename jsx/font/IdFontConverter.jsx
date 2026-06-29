/*

### 概要

フォントの種別（文字セット・P・UD・N・NT・ウエイト）をまとめて変更するスクリプト（InDesign 版）。

- 対象は「選択範囲 / ストーリー / ドキュメント全体 / アクティブスプレッド」から選択
- 変換設定は 3 カラム（左=文字セット／中央=N・NT／右=UD・P）。各項目は「現状維持 / なし / あり」で切り替え
- 文字セットは Std / Pro / Pr5 / Pr6（N は別軸でトグル）
- 新ゴ ⇄ 新ゴNT を NT 設定で切り替え
- プリセット Max（収録最多の N なし＋UD＋P）/ MaxN（収録最多の N あり込み＋UD＋P）
- G-OTF 学参書体（常改 / 学参 / K書体）を A-OTF に統合（チェックボックス）
- A1明朝など特殊シリーズは太さ等価で対応（A-OTF A1明朝 Std B ＝ A P-OTF A1明朝 StdN R）
- AXIS（Type Project）は文字セット体系が異なるため専用処理（幅 Basic/Cond/Comp と Joyo は保持し、N と Std⇄Pro のみ切り替え）
- CID フォント・実行前確認は先頭のスイッチ（UI 非表示）で制御
- 段落スタイル・文字スタイル・合成フォントの各エントリ、ロック / 非表示オブジェクトも対象に含められる
- 表セル（Cell）の選択にも対応
- 実行前に変更内容（旧 → 新）を和文フォント名でプレビュー確認、ページ上の位置順（上→下）に表示、未インストールフォントは事前に警告
- 同名ウエイトが見つからない場合は近いウエイトへ置換
- 適用は textStyleRange（同一書式の連続範囲）単位でまとめて処理（高速）。全体を 1 アンドゥにまとめる
- 変換対象ファミリーはフォントデータベースから生成（FONT_FAMILIES）
- フォントのインデックス化はダイアログ確定後まで遅延し、初回表示を高速化

### 参考

https://sttk3.com/blog/tips/illustrator/unify-character-set.html

### 紹介記事

https://note.com/dtp_tranist/n/n261c771b4b41

### 変更履歴

- v1.0.0 : 初版（Illustrator）
- v1.0.1 : 確認ダイアログを和文フォント名で表示・カンバス上から順（上→下）に並べ替え・→ の位置をそろえる（変更前列を固定幅）・タイトルのバージョン表記を削除・左右マージンを調整
- v1.0.1 (InDesign 移植) : 対象を選択 / ドキュメント / アクティブスプレッドに対応、フォント解決を app.fonts の PostScript 名インデックス化、適用を textStyleRange 単位 + 単一アンドゥに変更
- v1.1.0 : FontConverter.jsx を統合。対象範囲に「アクティブスプレッド」を追加（選択範囲 / ストーリー / ドキュメント / アクティブスプレッドの 4 モード）
- v1.1.1 : 初回ダイアログ表示を高速化（フォントインデックス化を確定後まで遅延）、ロック解除中のエラー時も確実に復元（try/finally）、Max/MaxN プリセット時に AXIS を ProN へ寄せる対応、概要・ツールチップを更新

*/

#target indesign

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.1.1";

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 実行前に変更内容を確認するか（UI 非表示）/ Whether to confirm changes before running (hidden) */
var CONFIRM_BEFORE_RUN = true;

/* CID フォント（文字セット表記なし）を OTF へ変換するか / Whether to convert CID fonts to OTF */
var CONVERT_CID_TO_OTF = true;

/* G-OTF 学参書体を A-OTF の通常書体に統合してから変換するか（UI 非表示）/ Merge G-OTF Gakusan fonts into A-OTF before converting (hidden) */
var INTEGRATE_GAKUSAN_TO_STANDARD = true;

/* G-OTF 学参書体の PostScript 名プレフィックス（GJ=常改 / G=学参 / K=K書体）。長い順に試す / Gakusan PostScript prefixes, tried longest first */
var GAKUSAN_PREFIXES = ["GJ", "G", "K"];

/* 特殊シリーズ（太さが等価なメンバーを完全名で結ぶ）/ Special series: members of equal thickness linked by full PostScript name
   A1明朝は Std が Bold 1 ウェイトのみで、A P-OTF StdN の Regular と同じ太さ。これを同一シリーズとして対応付ける。
   members のキーは文字セット（+N）。汎用ロジックではなくこの表だけで変換する。 */
var SPECIAL_SERIES = [
    {
        label: "A1明朝",
        members: {
            "Std": "A1MinchoStd-Bold",
            "StdN": "PA1MinchoStdN-Regular"
        }
    }
];

/* CID 変換時、文字セットが「現状維持」のときに補完する文字セット / Charset filled in when converting CID with "keep" */
var CID_FALLBACK_CHARSET = "Pr6";

/* 文字セットを収録文字数の多い順に並べたランキング / Charsets ranked by glyph count, richest first
   Max  は N なしの種類から、MaxN は N あり込みから、最初に選べる文字セットを採用する */
var CHARSET_RANK_NO_N = ["Pr6", "Pr5", "Pro", "Std", "Min2", "Min"];
var CHARSET_RANK_WITH_N = ["Pr6N", "Pr6", "Pr5N", "Pr5", "ProN", "Pro", "StdN", "Std", "Min2", "Min"];

/* 近いウエイトを探す並び順（軽い → 重い）/ Weight order for nearest-weight search (light to heavy) */
var WEIGHT_ORDER = [
    "Light", "Regular", "Medium",
    "DemiBold", "Demi", "DeBold",
    "Bold", "ExtraBold", "ExBold",
    "Heavy", "ExtraHeavy", "ExHeavy", "Ultra"
];

/* 変換対象のフォントファミリー定義（フォントデータベースから生成）/ Target font families (generated from the font database)
   regex の最後のグループをウエイト、それ以外を値（P / UD / Pro|Pr5|Pr6 / N）で自動判別する */
var FONT_FAMILIES = [
    { label: "見出ミンMA31", baseName: "MidashiMinMA31", regex: /^(P)?MidashiMinMA31(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "見出ゴMB31", baseName: "MidashiGoMB31", regex: /^(P)?MidashiGoMB31(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "太ミンA101", baseName: "FutoMinA101", regex: /^(P)?FutoMinA101(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "ゴシックMB101", baseName: "GothicMB101", regex: /^(P)?GothicMB101(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "UD新ゴ コンデンス90", baseName: "ShinGoCOniz", regex: /^(P)?(UD)?ShinGoCOniz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "UD新ゴ コンデンス80", baseName: "ShinGoCOeiz", regex: /^(P)?(UD)?ShinGoCOeiz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "UD新ゴ コンデンス70", baseName: "ShinGoCOsez", regex: /^(P)?(UD)?ShinGoCOsez(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "UD新ゴ コンデンス60", baseName: "ShinGoCOsiz", regex: /^(P)?(UD)?ShinGoCOsiz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "UD新ゴ コンデンス50", baseName: "ShinGoCOfiz", regex: /^(P)?(UD)?ShinGoCOfiz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "見出ミン", baseName: "MidashiMin", regex: /^MidashiMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "太ゴB101", baseName: "FutoGoB101", regex: /^(P)?FutoGoB101(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "新正楷書CBSK1", baseName: "ShinseiKai", regex: /^ShinseiKai(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "凸版文久明朝", baseName: "BunkyuMin", regex: /^(P)?BunkyuMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "中ゴシックBBB", baseName: "GothicBBB", regex: /^(P)?GothicBBB(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "見出ゴ", baseName: "MidashiGo", regex: /^MidashiGo(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "正楷書CB1", baseName: "SeiKaiCB1", regex: /^(P)?SeiKaiCB1(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "黎ミンY10", baseName: "ReimYonz", regex: /^(P)?ReimYonz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "黎ミンY20", baseName: "ReimYtwz", regex: /^(P)?ReimYtwz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "黎ミンY30", baseName: "ReimYthz", regex: /^(P)?ReimYthz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "黎ミンY40", baseName: "ReimYfoz", regex: /^(P)?ReimYfoz(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "秀英明朝", baseName: "ShueiMin", regex: /^(P)?ShueiMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "凸版文久ゴシック", baseName: "BunkyuGo", regex: /^(P)?BunkyuGo(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "新丸ゴ", baseName: "ShinMGo", regex: /^(P)?(UD)?ShinMGo(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "リュウミン", baseName: "Ryumin", regex: /^(P)?Ryumin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "新ゴ / 新ゴNT", baseName: "ShinGo", regex: /^(P)?(UD)?ShinGo(NT)?(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "UD黎ミン", baseName: "Reimin", regex: /^(P)?(UD)?Reimin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "游明朝体", baseName: "YuMin", regex: /^YuMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "黎ミン", baseName: "Reim", regex: /^(P)?Reim(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "游ゴシック体", baseName: "YuGo", regex: /^YuGo(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },

    // --- フォントワークス書体 ---
    { label: "セザンヌ", baseName: "Cezanne", regex: /^(P)?Cezanne(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "DNP 秀英明朝", baseName: "FOT_DNPShueiMin", regex: /^(P)?FOT_DNPShueiMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "ハミング", baseName: "Humming", regex: /^(P)?Humming(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "マティス", baseName: "Matisse", regex: /^(P)?Matisse(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "モード明朝Aラージ", baseName: "ModeMinALarge", regex: /^(P)?ModeMinALarge(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "モード明朝Bラージ", baseName: "ModeMinBLarge", regex: /^(P)?ModeMinBLarge(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "ニューセザンヌ", baseName: "NewCezanne", regex: /^(P)?NewCezanne(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "ニューグレコ", baseName: "NewGreco", regex: /^(P)?NewGreco(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "ニューロダン", baseName: "NewRodin", regex: /^(P)?NewRodin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "ロダン", baseName: "Rodin", regex: /^(P)?Rodin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "スーラ", baseName: "Seurat", regex: /^(P)?Seurat(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "スキップ", baseName: "Skip", regex: /^(P)?Skip(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "テロップ明朝", baseName: "TelopMin", regex: /^(P)?TelopMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫A見出ミン", baseName: "TsukuAMDMin", regex: /^(P)?TsukuAMDMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫Aオールド明朝", baseName: "TsukuAOldMin", regex: /^(P)?TsukuAOldMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫Aヴィンテージ明L", baseName: "TsukuAVintageMinL", regex: /^(P)?TsukuAVintageMinL(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫Aヴィンテージ明S", baseName: "TsukuAVintageMinS", regex: /^(P)?TsukuAVintageMinS(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫B見出ミン", baseName: "TsukuBMDMin", regex: /^(P)?TsukuBMDMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫B明朝", baseName: "TsukuBMin", regex: /^(P)?TsukuBMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫Bオールド明朝", baseName: "TsukuBOldMin", regex: /^(P)?TsukuBOldMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫Bヴィンテージ明L", baseName: "TsukuBVintageMinL", regex: /^(P)?TsukuBVintageMinL(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫Bヴィンテージ明S", baseName: "TsukuBVintageMinS", regex: /^(P)?TsukuBVintageMinS(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫C見出ミン", baseName: "TsukuCMDMin", regex: /^(P)?TsukuCMDMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫Cオールド明朝", baseName: "TsukuCOldMin", regex: /^(P)?TsukuCOldMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫Cヴィンテージ明L", baseName: "TsukuCVintageMinL", regex: /^(P)?TsukuCVintageMinL(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫Cヴィンテージ明S", baseName: "TsukuCVintageMinS", regex: /^(P)?TsukuCVintageMinS(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫ゴシック", baseName: "TsukuGo", regex: /^(P)?TsukuGo(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫明朝", baseName: "TsukuMin", regex: /^(P)?TsukuMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "筑紫新聞明朝", baseName: "TsukuNewsMin", regex: /^(P)?TsukuNewsMin(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "UD角ゴ_ラージ", baseName: "UDKakugo_Large", regex: /^(P)?UDKakugo_Large(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "UD角ゴ_スモール", baseName: "UDKakugo_Small", regex: /^(P)?UDKakugo_Small(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "UD丸ゴ_ラージ", baseName: "UDMarugo_Large", regex: /^(P)?UDMarugo_Large(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "UD丸ゴ_スモール", baseName: "UDMarugo_Small", regex: /^(P)?UDMarugo_Small(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ },
    { label: "UD明朝", baseName: "UDMincho", regex: /^(P)?UDMincho(?:(Pro|Pr5|Pr6|Std)(N)?)?-(.+)$/ }
];

/* AXIS（Type Project）は文字セット体系が Morisawa 系と異なるため専用処理する / AXIS uses a different charset scheme; handled separately
   - 幅（Basic / Cond / Comp）と Joyo は保持する
   - N と Std⇄Pro だけ切り替える（AXIS に無い Pr5 / Pr6 を選んでも現状維持）
   - P / UD / NT は AXIS に無いので無視する
   グループ: 1=幅(Basic|Cond|Comp) / 2=文字セット(Std|Pro|Joyo) / 3=N / 4=ウエイト */
var AXIS_FAMILY = {
    label: "AXIS",
    baseName: "Axis",
    regex: /^Axis(Basic|Cond|Comp)?(Std|Pro|Joyo)?(N)?-(.+)$/
};

// =========================================
// ローカライズ / Localization
// =========================================

/* 表示言語を判定（日本語環境なら ja）/ Detect UI language (ja for Japanese locale) */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "フォント種別変更", en: "Font Variant Switcher" }
    },
    panel: {
        target: { ja: "対象範囲", en: "Target Scope" },
        conversion: { ja: "変換設定", en: "Conversion Settings" },
        charset: { ja: "文字セット", en: "Character Set" },
        nSetting: { ja: "N（JIS2004対応）", en: "N (JIS2004)" },
        ntSetting: { ja: "新ゴ / 新ゴNT", en: "ShinGo / ShinGo NT" },
        udSetting: { ja: "UD", en: "UD" },
        pSetting: { ja: "AP版", en: "AP (Proportional)" },
        options: { ja: "処理オプション", en: "Processing Options" }
    },
    radio: {
        targetSelection: { ja: "選択範囲", en: "Selection" },
        targetStory: { ja: "ストーリー", en: "Story" },
        targetDocument: { ja: "ドキュメント", en: "Document" },
        targetSpread: { ja: "アクティブスプレッド", en: "Active spread" },
        keep: { ja: "現状維持", en: "Keep current" },
        nOff: { ja: "Nなし", en: "Without N" },
        nOn: { ja: "Nあり", en: "With N" },
        ntOff: { ja: "NTなし", en: "Without NT" },
        ntOn: { ja: "NTあり", en: "With NT" },
        udOff: { ja: "UDなし", en: "Without UD" },
        udOn: { ja: "UDあり", en: "With UD" },
        pOff: { ja: "Pなし", en: "Without P" },
        pOn: { ja: "Pあり", en: "With P" }
    },
    checkbox: {
        integrateGakusan: {
            ja: "G-OTF学参書体をA-OTFに統合",
            en: "Merge G-OTF Gakusan fonts into A-OTF"
        },
        includeStyles: {
            ja: "段落/文字スタイルも変更",
            en: "Also update paragraph/character styles"
        },
        includeComposite: {
            ja: "合成フォントを含む",
            en: "Include composite fonts"
        },
        includeLocked: {
            ja: "ロックされたオブジェクトも対象にする",
            en: "Include locked objects"
        },
        includeHidden: {
            ja: "非表示オブジェクトも対象にする",
            en: "Include hidden objects"
        }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" },
        run: { ja: "実行", en: "Run" }
    },
    confirm: {
        title: { ja: "変更内容の確認", en: "Confirm Changes" },
        willChange: {
            ja: "以下のフォントを変更します。",
            en: "The following fonts will be changed."
        },
        nearWeight: {
            ja: "同じウエイトのフォントが見つかりません。近いウエイトに置換します。",
            en: "Exact weight not found. The nearest weight will be substituted."
        },
        notInstalled: {
            ja: "次のフォントはインストールされていないため変更されません。",
            en: "The following fonts are not installed and will not be changed."
        },
        prompt: { ja: "実行しますか？", en: "Run now?" }
    },
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSelection: {
            ja: "テキストオブジェクトを選択してください。",
            en: "Please select text objects."
        },
        noTarget: {
            ja: "対象となるテキストが見つかりませんでした。",
            en: "No target text was found."
        },
        noLayoutWindow: {
            ja: "レイアウトウィンドウがアクティブではありません。ドキュメントのページを表示してから実行してください。",
            en: "No layout window is active. Switch to a document layout view and try again."
        },
        noChange: {
            ja: "変更対象のフォントが見つかりませんでした。",
            en: "No fonts to change were found."
        },
        done: { ja: "処理が完了しました。", en: "Done." },
        changedCount: { ja: "変更数", en: "Changed" }
    },
    tooltip: {
        target: {
            ja: "フォント種別を変更する範囲を選びます。選択範囲、選択が属するストーリー全体、ドキュメント全体、アクティブスプレッド内から選択できます。",
            en: "Choose the scope for changing font variants: the selection, the whole story containing the selection, the entire document, or the active spread."
        },
        charset: {
            ja: "文字セットを Std / Pro / Pr5 / Pr6 から選びます。N の有無は別の「N設定」で切り替えます。",
            en: "Choose the character set: Std, Pro, Pr5, or Pr6. Toggle the N variant separately under N Variant."
        },
        nSetting: {
            ja: "N あり/なしを切り替えます。例：Pr6 → Pr6N。現状維持では現在の N 状態を保ちます。",
            en: "Toggle the N variant on or off, for example Pr6 -> Pr6N. Keep current preserves the current N state."
        },
        udSetting: {
            ja: "UD 書体があるファミリーだけ、UD あり/なしを切り替えます。対応しないファミリーでは無視されます。",
            en: "Toggle the UD variant only for families that support it. Unsupported families are ignored."
        },
        pSetting: {
            ja: "AP版書体へ切り替えます。対応しないファミリーでは無視されます。",
            en: "Switch to the A P-OTF (proportional) version. Unsupported families are ignored."
        },
        includeLocked: {
            ja: "ロック中のオブジェクトも一時的に解除して変更し、処理後に元のロック状態へ戻します。",
            en: "Temporarily unlock locked objects, update them, then restore the original lock state."
        },
        includeHidden: {
            ja: "非表示のオブジェクト（レイヤー非表示を含む）も対象に含めて変更します。",
            en: "Include hidden objects (including hidden layers) in the scope."
        },
        nt: {
            ja: "新ゴと新ゴNTを切り替えます。対象は新ゴ系ファミリーのみです。",
            en: "Switch between ShinGo and ShinGo NT. This applies only to ShinGo families."
        },
        integrateGakusan: {
            ja: "G-OTF の学参・常改・K書体を、通常の A-OTF 書体として扱って変換します。",
            en: "Treat G-OTF Gakusan, revised, and K fonts as standard A-OTF fonts before conversion."
        },
        includeStyles: {
            ja: "段落スタイル・文字スタイルに設定されたフォントも、同じ条件で変更します。",
            en: "Also update fonts assigned in paragraph and character styles using the same conversion settings."
        },
        includeComposite: {
            ja: "合成フォントの各エントリ（漢字・かな・全角約物など）に設定されたフォントも、同じ条件で変更します。",
            en: "Also update fonts assigned in each composite font entry (kanji, kana, etc.) using the same conversion settings."
        },
        presetMax: {
            ja: "N なしで収録文字数が最も多い文字セットを選び、UD・P・NT をありにします。",
            en: "Select the richest available character set without N, and turn on UD, P, and NT."
        },
        presetMaxN: {
            ja: "N ありを含めて収録文字数が最も多い文字セットを選び、UD・P・NT をありにします。",
            en: "Select the richest available character set including N variants, and turn on UD, P, and NT."
        }
    }
};

/* ドット区切りのキーでラベルを取得 / Resolve a label by dot-separated key path */
function L(labelPath) {
    var keys = labelPath.split(".");
    var node = LABELS;
    for (var i = 0; i < keys.length; i++) {
        node = node[keys[i]];
        if (node == null) return labelPath;
    }
    var text = node[currentLanguage] || node.en || "";
    return text.replace(/\{slash\}/g, "/");
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(labelPath) {
    return L(labelPath) + (currentLanguage === "ja" ? "：" : ":");
}

// =========================================
// UI 共通 / UI helpers
// =========================================

/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout (alignChildren switches by orientation) */
function setupGroup(group, orientation, spacing) {
    var groupOrientation = orientation || "column";
    group.orientation = groupOrientation;
    /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
    group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
    group.alignment = "fill";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// =========================================
// メイン / Main
// =========================================

(function () {

    /* ドキュメントの有無を確認 / Ensure a document is open */
    if (app.documents.length === 0) {
        alert(L("alert.noDocument"));
        return;
    }

    var activeDocument = app.activeDocument;

    /* 選択範囲を判定するテキスト系オブジェクト型 / Text-bearing object types found in a selection */
    var TEXT_TYPES = {
        Text: 1, Story: 1, Paragraph: 1, Line: 1, Word: 1, Character: 1,
        TextStyleRange: 1, InsertionPoint: 1, TextColumn: 1, Cell: 1
    };

    // -----------------------------------------
    // フォントインデックス / Font index by PostScript name
    // -----------------------------------------

    /* PostScript 名 → Font オブジェクト。installed はインストール済みのみ、all は表示名解決用（未インストール含む）
       実体の構築はダイアログ確定後（実処理の直前）まで遅延する。フォント数が多い環境での初回表示を速くするため。
       Maps are populated lazily after the dialog is confirmed (right before processing), so the first dialog opens fast. */
    var installedFontByPs = {};
    var allFontByPs = {};
    function indexFonts() {
        var fonts = app.fonts;

        // まとめて取得する（1 書体ずつのプロパティ参照は DOM 往復が多く非常に遅い）
        // Fetch in bulk; per-font property access crosses the DOM many times and is very slow.
        var elements = null, psNames = null, statuses = null;
        try {
            elements = fonts.everyItem().getElements();
            psNames = fonts.everyItem().postscriptName;
            statuses = fonts.everyItem().status;
        } catch (bulkError) {
            elements = null;
        }

        if (elements && psNames && elements.length === psNames.length) {
            for (var i = 0; i < elements.length; i++) {
                var psName = psNames[i];
                if (!psName) continue;
                if (!allFontByPs.hasOwnProperty(psName)) allFontByPs[psName] = elements[i];
                var installed = statuses ? (statuses[i] === FontStatus.INSTALLED) : true;
                if (installed && !installedFontByPs.hasOwnProperty(psName)) installedFontByPs[psName] = elements[i];
            }
            return;
        }

        // フォールバック：1 件ずつ / Fallback: one by one
        for (var j = 0; j < fonts.length; j++) {
            var font = fonts[j];
            var ps;
            try { ps = font.postscriptName; } catch (e) { continue; }
            if (!ps) continue;
            if (!allFontByPs.hasOwnProperty(ps)) allFontByPs[ps] = font;
            var inst;
            try { inst = (font.status === FontStatus.INSTALLED); } catch (e2) { inst = true; }
            if (inst && !installedFontByPs.hasOwnProperty(ps)) installedFontByPs[ps] = font;
        }
    }

    // -----------------------------------------
    // ダイアログ / Dialog
    // -----------------------------------------

    var mainDialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
    mainDialog.orientation = "column";
    mainDialog.alignChildren = "left";

    /* 対象パネル（ラジオは縦並び）/ Target panel (radios in a column) */
    var targetPanel = addPanel(mainDialog, "panel.target");
    targetPanel.helpTip = L("tooltip.target");
    var rbTargetSelection = targetPanel.add("radiobutton", undefined, L("radio.targetSelection"));
    var rbTargetStory = targetPanel.add("radiobutton", undefined, L("radio.targetStory"));
    var rbTargetDocument = targetPanel.add("radiobutton", undefined, L("radio.targetDocument"));
    var rbTargetSpread = targetPanel.add("radiobutton", undefined, L("radio.targetSpread"));
    rbTargetSelection.value = true;

    /* 変換設定パネル（文字セット/N/UD/P と Max/MaxN をまとめる）/ Conversion panel (variant columns + presets) */
    var conversionPanel = addPanel(mainDialog, "panel.conversion");

    /* 変換設定: 3 カラム（左=文字セット / 中央=N・NT / 右=UD・P）/ Conversion: 3 columns */
    var variantColumns = conversionPanel.add("group");
    setupGroup(variantColumns, "row");
    variantColumns.alignChildren = ["fill", "top"];

    /* 左カラム: 文字セットパネル（Std/Pro/Pr5/Pr6 はローカライズ不要の固有名）/ Left: character set */
    var charsetPanel = addPanel(variantColumns, "panel.charset");
    charsetPanel.helpTip = L("tooltip.charset");
    var rbCharsetKeep = charsetPanel.add("radiobutton", undefined, L("radio.keep"));
    var rbCharsetStd = charsetPanel.add("radiobutton", undefined, "Std");
    var rbCharsetPro = charsetPanel.add("radiobutton", undefined, "Pro");
    var rbCharsetPr5 = charsetPanel.add("radiobutton", undefined, "Pr5");
    var rbCharsetPr6 = charsetPanel.add("radiobutton", undefined, "Pr6");
    rbCharsetKeep.value = true;

    /* 中央カラム: N 設定・NT 設定 / Center: N and NT */
    var nNtColumn = variantColumns.add("group");
    setupGroup(nNtColumn, "column");

    var nVariantPanel = addPanel(nNtColumn, "panel.nSetting");
    nVariantPanel.helpTip = L("tooltip.nSetting");
    var rbNKeep = nVariantPanel.add("radiobutton", undefined, L("radio.keep"));
    var rbNOff = nVariantPanel.add("radiobutton", undefined, L("radio.nOff"));
    var rbNOn = nVariantPanel.add("radiobutton", undefined, L("radio.nOn"));
    rbNKeep.value = true;

    /* NT 設定パネル（新ゴ ⇄ 新ゴNT）/ NT panel (ShinGo <-> ShinGo NT) */
    var ntPanel = addPanel(nNtColumn, "panel.ntSetting");
    ntPanel.helpTip = L("tooltip.nt");
    var rbNTKeep = ntPanel.add("radiobutton", undefined, L("radio.keep"));
    var rbNTOff = ntPanel.add("radiobutton", undefined, L("radio.ntOff"));
    var rbNTOn = ntPanel.add("radiobutton", undefined, L("radio.ntOn"));
    rbNTKeep.value = true;

    /* 右カラム: UD 設定・P 設定 / Right: UD and P */
    var udProportionalColumn = variantColumns.add("group");
    setupGroup(udProportionalColumn, "column");

    var udVariantPanel = addPanel(udProportionalColumn, "panel.udSetting");
    udVariantPanel.helpTip = L("tooltip.udSetting");
    var rbUDKeep = udVariantPanel.add("radiobutton", undefined, L("radio.keep"));
    var rbUDOff = udVariantPanel.add("radiobutton", undefined, L("radio.udOff"));
    var rbUDOn = udVariantPanel.add("radiobutton", undefined, L("radio.udOn"));
    rbUDKeep.value = true;

    var proportionalPanel = addPanel(udProportionalColumn, "panel.pSetting");
    proportionalPanel.helpTip = L("tooltip.pSetting");
    var rbPKeep = proportionalPanel.add("radiobutton", undefined, L("radio.keep"));
    var rbPOff = proportionalPanel.add("radiobutton", undefined, L("radio.pOff"));
    var rbPOn = proportionalPanel.add("radiobutton", undefined, L("radio.pOn"));
    rbPKeep.value = true;

    /* 文字セット名 → ラジオボタンの対応 / Map charset name to its radio button */
    var charsetRadioByName = {
        Pro: rbCharsetPro,
        Pr5: rbCharsetPr5,
        Pr6: rbCharsetPr6,
        Std: rbCharsetStd
    };

    /* ランキング順に、最初に選べる文字セット（＋N）を UI へ反映 / Apply the richest available charset (+N) to the UI */
    function applyRichestCharset(rankList) {
        for (var i = 0; i < rankList.length; i++) {
            var token = rankList[i];
            var withN = token.charAt(token.length - 1) === "N";
            var baseCharset = withN ? token.substring(0, token.length - 1) : token;
            var radio = charsetRadioByName[baseCharset];
            if (radio) {
                radio.value = true;
                if (withN) { rbNOn.value = true; } else { rbNOff.value = true; }
                return;
            }
        }
    }

    /* Max/MaxN プリセットが押されたか / Whether a Max/MaxN preset is active
       AXIS は Std と ProN しか無いため、プリセット時は設定に関わらず ProN へ寄せる。
       押下後に文字セット/N を手動変更したら解除する。/ AXIS only has Std & ProN, so a preset forces ProN; cleared if charset/N changes manually. */
    var maxPresetActive = false;
    var presetResetRadios = [rbCharsetKeep, rbCharsetStd, rbCharsetPro, rbCharsetPr5, rbCharsetPr6, rbNKeep, rbNOff, rbNOn];
    for (var presetResetIndex = 0; presetResetIndex < presetResetRadios.length; presetResetIndex++) {
        presetResetRadios[presetResetIndex].onClick = function () { maxPresetActive = false; };
    }

    /* プリセット（Max=収録最多のNなし＋UD＋P、MaxN=収録最多のNあり込み＋UD＋P）/ Presets (Max / MaxN) */
    var presetRow = conversionPanel.add("group");
    setupGroup(presetRow, "row");
    presetRow.alignment = ["center", "top"]; // 左右中央 / horizontally centered
    presetRow.margins = [0, 5, 0, 0]; // 上に 5px の余白 / 5px top margin
    var presetMaxButton = presetRow.add("button", undefined, "Max");
    presetMaxButton.helpTip = L("tooltip.presetMax");
    presetMaxButton.onClick = function () {
        applyRichestCharset(CHARSET_RANK_NO_N);
        rbUDOn.value = true;
        rbPOn.value = true;
        rbNTOn.value = true;
        maxPresetActive = true;
    };
    var presetMaxNButton = presetRow.add("button", undefined, "MaxN");
    presetMaxNButton.helpTip = L("tooltip.presetMaxN");
    presetMaxNButton.onClick = function () {
        applyRichestCharset(CHARSET_RANK_WITH_N);
        rbUDOn.value = true;
        rbPOn.value = true;
        rbNTOn.value = true;
        maxPresetActive = true;
    };

    /* オプションパネル / Options panel */
    var optionsPanel = addPanel(mainDialog, "panel.options");
    var cbIntegrateGakusan = optionsPanel.add("checkbox", undefined, L("checkbox.integrateGakusan"));
    cbIntegrateGakusan.helpTip = L("tooltip.integrateGakusan");
    cbIntegrateGakusan.value = INTEGRATE_GAKUSAN_TO_STANDARD;
    var cbIncludeStyles = optionsPanel.add("checkbox", undefined, L("checkbox.includeStyles"));
    cbIncludeStyles.helpTip = L("tooltip.includeStyles");
    cbIncludeStyles.value = true;
    var cbIncludeComposite = optionsPanel.add("checkbox", undefined, L("checkbox.includeComposite"));
    cbIncludeComposite.helpTip = L("tooltip.includeComposite");
    cbIncludeComposite.value = true;
    var cbIncludeLocked = optionsPanel.add("checkbox", undefined, L("checkbox.includeLocked"));
    cbIncludeLocked.helpTip = L("tooltip.includeLocked");
    cbIncludeLocked.value = false;
    var cbIncludeHidden = optionsPanel.add("checkbox", undefined, L("checkbox.includeHidden"));
    cbIncludeHidden.helpTip = L("tooltip.includeHidden");
    cbIncludeHidden.value = false;

    /* ボタン（Mac 規約: キャンセル → OK、OK は非ローカライズ）/ Buttons (Mac order: Cancel then OK, OK is not localized) */
    var buttonRow = mainDialog.add("group");
    buttonRow.orientation = "row";
    buttonRow.alignment = "right";
    var cancelButton = buttonRow.add("button", undefined, L("button.cancel"), { name: "cancel" });
    var okButton = buttonRow.add("button", undefined, "OK", { name: "ok" });

    cancelButton.onClick = function () { mainDialog.close(0); };
    okButton.onClick = function () { mainDialog.close(1); };

    if (mainDialog.show() !== 1) {
        return;
    }

    // -----------------------------------------
    // 設定値の取得 / Read settings
    // -----------------------------------------

    var targetMode = "selection";
    if (rbTargetStory.value) targetMode = "story";
    if (rbTargetDocument.value) targetMode = "document";
    if (rbTargetSpread.value) targetMode = "spread";

    var charsetMode = "keep"; // keep / Pro / Pr5 / Pr6 / Std
    if (rbCharsetPro.value) charsetMode = "Pro";
    if (rbCharsetPr5.value) charsetMode = "Pr5";
    if (rbCharsetPr6.value) charsetMode = "Pr6";
    if (rbCharsetStd.value) charsetMode = "Std";

    var nMode = radioMode(rbNOn, rbNOff);   // keep / on / off
    var udMode = radioMode(rbUDOn, rbUDOff);
    var pMode = radioMode(rbPOn, rbPOff);
    var ntMode = radioMode(rbNTOn, rbNTOff);
    var maxPreset = maxPresetActive; // Max/MaxN プリセット中か（AXIS を ProN へ寄せる）/ whether a Max/MaxN preset is active (forces AXIS to ProN)

    var integrateGakusan = cbIntegrateGakusan.value;
    var includeStyles = cbIncludeStyles.value;
    var includeComposite = cbIncludeComposite.value;
    var includeLocked = cbIncludeLocked.value;
    var includeHidden = cbIncludeHidden.value;
    var confirmBeforeRun = CONFIRM_BEFORE_RUN;

    /* 選択範囲・ストーリーは選択が必要 / Selection and story modes require a selection */
    if ((targetMode === "selection" || targetMode === "story") && app.selection.length === 0) {
        alert(L("alert.noSelection"));
        return;
    }

    /* スプレッドモードはレイアウトウィンドウが必要 / Spread mode needs an active layout window */
    var isLayoutWindow = false;
    try { isLayoutWindow = (app.activeWindow.constructor.name === "LayoutWindow"); } catch (e) { isLayoutWindow = false; }
    if (targetMode === "spread" && !isLayoutWindow) {
        alert(L("alert.noLayoutWindow"));
        return;
    }

    // フォントインデックスをここで構築（ダイアログ確定・入力検証を通過してから）/ Build the font index now (after the dialog is confirmed and validated)
    indexFonts();

    // -----------------------------------------
    // 対象テキストの収集 / Collect target text
    // -----------------------------------------

    var seenStoryIds = {};
    var targets = []; // {story, ranges, frames, sortKey}
    collectTargets(targetMode);

    // ページ上の位置（上→下、同じ高さは左→右）で並べ替え / Sort by page position (top→bottom, then left→right)
    sortTargetsByPosition(targets);

    if (targets.length === 0 && !includeStyles && !includeComposite) {
        alert(L("alert.noTarget"));
        return;
    }

    // -----------------------------------------
    // 変更内容の事前計算 / Pre-compute changes
    // -----------------------------------------

    var processedFontNames = {};
    var directChanges = [];     // {oldName, newName}
    var weightSubChanges = [];  // {oldName, newName, oldWeight, newWeight}
    var missingChanges = [];    // {oldName, newName}

    /* 対象をスキャンして変換候補を収集 / Scan targets to collect change candidates */
    for (var ti = 0; ti < targets.length; ti++) {
        scanTargetFonts(targets[ti]);
    }

    /* 段落・文字スタイルもスキャン / Scan paragraph and character styles too */
    if (includeStyles) {
        scanStylesForChanges(activeDocument.allParagraphStyles);
        scanStylesForChanges(activeDocument.allCharacterStyles);
    }

    /* 合成フォントのエントリもスキャン / Scan composite font entries too */
    if (includeComposite) {
        scanCompositeFontsForChanges();
    }

    if (directChanges.length === 0 && weightSubChanges.length === 0 && missingChanges.length === 0) {
        alert(L("alert.noChange"));
        return;
    }

    // -----------------------------------------
    // 確認ダイアログ / Confirmation dialog
    // -----------------------------------------

    var selectedOldNames = null; // null = すべて適用 / null = apply all
    if (confirmBeforeRun) {
        var confirmResult = showConfirmDialog();
        if (!confirmResult.ok) {
            return;
        }
        selectedOldNames = confirmResult.selected;
    }

    // -----------------------------------------
    // 変換マップの構築 / Build font-name map
    // -----------------------------------------

    var fontNameMap = {}; // oldName -> newName
    addSelectedChanges(directChanges, fontNameMap, selectedOldNames);
    addSelectedChanges(weightSubChanges, fontNameMap, selectedOldNames);

    // -----------------------------------------
    // 適用（全体を 1 アンドゥにまとめる）/ Apply (wrapped as a single undo step)
    // -----------------------------------------

    var changedCount = 0;
    app.doScript(function () {
        // 同一ストーリーを複数 target で重複カウントしない / Don't count the same story twice across targets
        var countedStoryIds = {};
        for (var tj = 0; tj < targets.length; tj++) {
            var target = targets[tj];
            if (applyChangesToTarget(target, fontNameMap) > 0) {
                var storyId = null;
                try { storyId = target.story.id; } catch (e) { storyId = null; }
                if (storyId === null) {
                    changedCount++;
                } else if (!countedStoryIds[storyId]) {
                    countedStoryIds[storyId] = true;
                    changedCount++;
                }
            }
        }
        // スタイルは適用のみ（テキストオブジェクト数には数えない）/ Styles are applied but not counted as text objects
        if (includeStyles) {
            applyChangesToStyles(activeDocument.allParagraphStyles, fontNameMap);
            applyChangesToStyles(activeDocument.allCharacterStyles, fontNameMap);
        }
        // 合成フォントも適用のみ（テキストオブジェクト数には数えない）/ Composite fonts applied but not counted
        if (includeComposite) {
            applyChangesToCompositeFonts(fontNameMap);
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, L("dialog.title"));

    alert(L("alert.done") + "\n\n" + labelText("alert.changedCount") + changedCount);

    // =========================================
    // 対象収集 / Target collection
    // =========================================

    /* 対象モードに応じてターゲット（ストーリー / テキスト範囲）を集める / Gather targets by mode */
    function collectTargets(mode) {
        if (mode === "document") {
            var stories = activeDocument.stories;
            for (var i = 0; i < stories.length; i++) {
                if (storyInScope(stories[i])) addStoryTarget(stories[i]);
            }
            return;
        }
        if (mode === "story") {
            // 選択が属するストーリー全体を対象にする / Whole story that the selection belongs to
            collectStoriesFromSelection(app.selection);
            return;
        }
        if (mode === "spread") {
            // アクティブスプレッド上のテキストフレームを対象にする / Text frames on the active spread
            var spreadFrames = collectActiveSpreadTextFrames();
            for (var s = 0; s < spreadFrames.length; s++) {
                var frame = spreadFrames[s];
                if (!includeLocked && isFrameLocked(frame)) continue;
                if (!includeHidden && isFrameHidden(frame)) continue;
                addStoryTarget(frame.parentStory);
            }
            return;
        }
        // selection
        collectFromSelection(app.selection);
    }

    /* アクティブスプレッド上のテキストフレームを集める（グループ内も含む）/ Collect text frames on the active spread (incl. groups) */
    function collectActiveSpreadTextFrames() {
        var result = [];
        try {
            var spread = app.activeWindow.activeSpread;
            var items = spread.allPageItems;
            for (var i = 0; i < items.length; i++) {
                if (items[i].constructor.name === "TextFrame") result.push(items[i]);
            }
        } catch (e) { }
        return result;
    }

    /* 選択アイテム（グループ内も再帰）から部分選択を尊重して集める / Collect targets, honoring partial selection */
    function collectFromSelection(items) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            var typeName;
            try { typeName = item.constructor.name; } catch (e) { continue; }

            if (typeName === "TextFrame") {
                if (!includeLocked && isFrameLocked(item)) continue;
                if (!includeHidden && isFrameHidden(item)) continue;
                addStoryTarget(item.parentStory);
            } else if (typeName === "Cell") {
                // 表セルは cell.texts[0] 経由で扱う / Table cells are handled via cell.texts[0]
                var cellText = cellTextOf(item);
                if (cellText) addTextTarget(cellText);
            } else if (TEXT_TYPES[typeName]) {
                addTextTarget(item); // 部分選択を尊重 / honor partial text selection
            } else if (typeName === "Group") {
                collectFromSelection(item.pageItems);
            }
        }
    }

    /* 選択アイテム（グループ内も再帰）から所属ストーリー全体を集める / Collect whole stories from a selection */
    function collectStoriesFromSelection(items) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            var typeName;
            try { typeName = item.constructor.name; } catch (e) { continue; }

            if (typeName === "TextFrame") {
                if (!includeLocked && isFrameLocked(item)) continue;
                if (!includeHidden && isFrameHidden(item)) continue;
                addStoryTarget(item.parentStory);
            } else if (typeName === "Cell") {
                var cellText = cellTextOf(item);
                if (cellText) {
                    try { addStoryTarget(cellText.parentStory); } catch (e2) { }
                }
            } else if (TEXT_TYPES[typeName]) {
                try { addStoryTarget(item.parentStory); } catch (e3) { }
            } else if (typeName === "Group") {
                collectStoriesFromSelection(item.pageItems);
            }
        }
    }

    /* 表セルのテキスト（cell.texts[0]）を安全に取得 / Safely get a cell's text (cell.texts[0]) */
    function cellTextOf(cell) {
        try {
            var texts = cell.texts;
            if (texts.length > 0) return texts[0];
        } catch (e) { }
        return null;
    }

    /* ストーリー単位のターゲットを追加（重複は除外）/ Add a story-level target (deduped) */
    function addStoryTarget(story) {
        if (!story || !story.isValid) return;
        var storyId;
        try { storyId = story.id; } catch (e) { return; }
        if (seenStoryIds[storyId]) return;
        seenStoryIds[storyId] = true;

        var frames = framesArray(story.textContainers);
        targets.push({
            story: story,
            ranges: story.textStyleRanges,
            frames: frames,
            sortKey: sortKeyOfFrames(frames)
        });
    }

    /* テキスト範囲（部分選択）単位のターゲットを追加 / Add a text-range target (partial selection) */
    function addTextTarget(textObject) {
        var story;
        try { story = textObject.parentStory; } catch (e) { return; }
        if (!story || !story.isValid) return;

        var frames;
        try { frames = framesArray(textObject.parentTextFrames); } catch (e2) { frames = []; }
        targets.push({
            story: story,
            ranges: textObject.textStyleRanges,
            frames: frames,
            sortKey: sortKeyOfFrames(frames)
        });
    }

    /* ストーリーが対象範囲（ロック/非表示フィルタ）に入るか / Whether a story has at least one in-scope container */
    function storyInScope(story) {
        var frames = story.textContainers;
        if (frames.length === 0) return false;
        for (var i = 0; i < frames.length; i++) {
            var frame = frames[i];
            if (!includeLocked && isFrameLocked(frame)) continue;
            if (!includeHidden && isFrameHidden(frame)) continue;
            return true;
        }
        return false;
    }

    /* コレクション/配列を素の配列へ / Normalize a collection or array into a plain array */
    function framesArray(frames) {
        var arr = [];
        try {
            for (var i = 0; i < frames.length; i++) arr.push(frames[i]);
        } catch (e) { }
        return arr;
    }

    /* 先頭フレームの位置を並べ替えキーに（InDesign は上が小さい Y）/ Sort key from the first frame (smaller Y = higher) */
    function sortKeyOfFrames(frames) {
        try {
            if (frames.length > 0) {
                var bounds = frames[0].geometricBounds; // [y1, x1, y2, x2] = [top, left, bottom, right]
                return [bounds[0], bounds[1]];
            }
        } catch (e) { }
        return [1e9, 1e9];
    }

    /* ターゲットをページ位置順（上→下、同じ高さは左→右）に並べ替え / Sort targets by page position */
    function sortTargetsByPosition(list) {
        list.sort(function (a, b) {
            if (Math.abs(a.sortKey[0] - b.sortKey[0]) > 1.0) return a.sortKey[0] - b.sortKey[0]; // top が小さい方が上
            return a.sortKey[1] - b.sortKey[1];                                                  // left が小さい方が先
        });
    }

    // =========================================
    // ロック・非表示判定 / Lock & hidden checks
    // =========================================

    /* 自身・所属レイヤー・祖先グループのいずれかがロックされているか / Whether the frame is effectively locked */
    function isFrameLocked(frame) {
        try { if (frame.locked) return true; } catch (e) { }
        try { if (frame.itemLayer.locked) return true; } catch (e2) { }
        var parent;
        try { parent = frame.parent; } catch (e3) { parent = null; }
        while (parent) {
            var typeName;
            try { typeName = parent.constructor.name; } catch (e4) { break; }
            if (typeName !== "Group") break;
            try { if (parent.locked) return true; } catch (e5) { }
            try { parent = parent.parent; } catch (e6) { break; }
        }
        return false;
    }

    /* 自身・所属レイヤー・祖先グループのいずれかが非表示か / Whether the frame is effectively hidden */
    function isFrameHidden(frame) {
        try { if (!frame.visible) return true; } catch (e) { }
        try { if (!frame.itemLayer.visible) return true; } catch (e2) { }
        var parent;
        try { parent = frame.parent; } catch (e3) { parent = null; }
        while (parent) {
            var typeName;
            try { typeName = parent.constructor.name; } catch (e4) { break; }
            if (typeName !== "Group") break;
            try { if (!parent.visible) return true; } catch (e5) { }
            try { parent = parent.parent; } catch (e6) { break; }
        }
        return false;
    }

    // =========================================
    // 変更内容のスキャン / Change scanning
    // =========================================

    /* ターゲットの textStyleRange を走査して変換候補に分類 / Scan a target's runs */
    function scanTargetFonts(target) {
        var ranges = target.ranges;
        for (var i = 0; i < ranges.length; i++) {
            var psName = fontPsNameOf(ranges[i].appliedFont);
            if (psName) classifyFontChange(psName);
        }
    }

    /* スタイルコレクションのフォントを変換候補に分類 / Classify fonts used by a style collection */
    function scanStylesForChanges(styles) {
        for (var k = 0; k < styles.length; k++) {
            var psName = styleFontPsName(styles[k]);
            if (psName) classifyFontChange(psName);
        }
    }

    /* 合成フォントの各エントリのフォントを変換候補に分類 / Classify fonts used by composite font entries */
    function scanCompositeFontsForChanges() {
        var compositeFonts = activeDocument.compositeFonts;
        for (var c = 0; c < compositeFonts.length; c++) {
            var entries = compositeFonts[c].compositeFontEntries;
            for (var e = 0; e < entries.length; e++) {
                var psName = styleFontPsName(entries[e]); // appliedFont を持つので同じ解決でよい
                if (psName) classifyFontChange(psName);
            }
        }
    }

    /* スタイルに設定されたフォントの PostScript 名（無ければ null）/ PostScript name of a style's applied font */
    function styleFontPsName(style) {
        try {
            return fontPsNameOf(style.appliedFont);
        } catch (e) {
            return null;
        }
    }

    /* Font オブジェクトまたはフォント名文字列から PostScript 名を取得 / Resolve a PostScript name from a Font or its name string */
    function fontPsNameOf(font) {
        if (!font) return null;
        try {
            if (typeof font === "string") {
                if (!font || font.indexOf("$ID") === 0) return null;
                var resolved = app.fonts.itemByName(font);
                if (!resolved.isValid) return null;
                return resolved.postscriptName;
            }
            return font.postscriptName;
        } catch (e) {
            return null;
        }
    }

    /* 旧フォント名を direct / weightSub / missing に分類（重複は 1 回だけ）/ Classify an old font name */
    function classifyFontChange(oldName) {
        if (processedFontNames[oldName]) return;
        processedFontNames[oldName] = true;

        // 特殊シリーズ（A1明朝 など）を優先 / Special series first (e.g. A1 Mincho)
        var seriesTarget = resolveSpecialSeries(oldName);
        if (seriesTarget !== undefined) {
            if (seriesTarget === null) return; // 該当するが変更不要・対応なし
            if (fontExists(seriesTarget)) {
                directChanges.push({ oldName: oldName, newName: seriesTarget });
            } else {
                missingChanges.push({ oldName: oldName, newName: seriesTarget });
            }
            return;
        }

        // AXIS は文字セット体系が異なるため専用処理 / AXIS has its own charset scheme
        var axis = parseAxisName(oldName);
        if (axis) {
            classifyByDesired(oldName, buildAxisNameHead(axis), axis.weight);
            return;
        }

        var parsed = parseFontName(oldName);
        if (!parsed) return;

        classifyByDesired(oldName, buildConvertedNameHead(parsed), parsed.weight);
    }

    /* 変換後の先頭（ウエイト除く）＋ウエイトから direct / weightSub / missing へ分類 / Classify by desired name head + weight */
    function classifyByDesired(oldName, nameHead, weight) {
        if (nameHead === null) return; // 変換対象外（CID 非変換など）/ not convertible

        var desiredName = nameHead + weight;
        if (desiredName === oldName) return; // 変化なし / no change

        if (fontExists(desiredName)) {
            directChanges.push({ oldName: oldName, newName: desiredName });
            return;
        }

        var nearWeight = findNearestWeight(nameHead, weight);
        if (nearWeight) {
            weightSubChanges.push({
                oldName: oldName,
                newName: nameHead + nearWeight,
                oldWeight: weight,
                newWeight: nearWeight
            });
            return;
        }

        missingChanges.push({ oldName: oldName, newName: desiredName });
    }

    // =========================================
    // フォント名の分解・組み立て / Parse & build font names
    // =========================================

    /* 特殊シリーズの対応を解決 / Resolve special series mapping
       戻り値: 文字列=変換先名 / null=該当するが対応なし・変更不要 / undefined=特殊シリーズではない */
    function resolveSpecialSeries(oldName) {
        for (var seriesIndex = 0; seriesIndex < SPECIAL_SERIES.length; seriesIndex++) {
            var members = SPECIAL_SERIES[seriesIndex].members;

            // oldName がどのキー（文字セット）のメンバーか / Which charset key oldName belongs to
            var sourceKey = null;
            for (var key in members) {
                if (members.hasOwnProperty(key) && members[key] === oldName) {
                    sourceKey = key;
                    break;
                }
            }
            if (!sourceKey) continue;

            var targetKey = computeSeriesTargetKey(sourceKey);
            if (!targetKey || !members.hasOwnProperty(targetKey)) return null;

            var targetName = members[targetKey];
            if (!targetName || targetName === oldName) return null;
            return targetName;
        }
        return undefined;
    }

    /* 設定から特殊シリーズの目標キー（文字セット[+N]）を算出 / Compute target charset key for special series */
    function computeSeriesTargetKey(sourceKey) {
        var sourceHasN = sourceKey.charAt(sourceKey.length - 1) === "N";
        var sourceBaseCharset = sourceHasN ? sourceKey.substring(0, sourceKey.length - 1) : sourceKey;

        var targetBaseCharset = (charsetMode === "keep") ? sourceBaseCharset : charsetMode;

        var targetHasN = sourceHasN;
        if (nMode === "on") targetHasN = true;
        if (nMode === "off") targetHasN = false;

        return targetBaseCharset + (targetHasN ? "N" : "");
    }

    /* AXIS の PostScript 名を分解（AXIS でなければ null）/ Decompose an AXIS PostScript name (null if not AXIS) */
    function parseAxisName(name) {
        var matched = name.match(AXIS_FAMILY.regex);
        if (!matched) return null;
        return {
            width: matched[1] || "",    // Basic / Cond / Comp / ""
            charset: matched[2] || "",  // Std / Pro / Joyo / ""
            hasN: !!matched[3],
            weight: matched[4]
        };
    }

    /* AXIS の変換後フォント名の先頭（ウエイト除く）を組み立て / Build the converted AXIS name head (without weight)
       幅と Joyo は保持し、N と Std⇄Pro のみ反映。AXIS に無い Pr5/Pr6 は現状維持。 */
    function buildAxisNameHead(axis) {
        // Max/MaxN プリセット時は、設定に関わらず ProN（AXIS の収録最多）へ寄せる。AXIS は Std と ProN しか無いため。
        // Under a Max/MaxN preset, force ProN (the richest AXIS charset) regardless of settings; AXIS only has Std & ProN.
        if (maxPreset) {
            return AXIS_FAMILY.baseName + axis.width + "ProN" + "-";
        }

        // Joyo は別体系なので維持（文字セット・N の切り替え対象外）/ Joyo is a separate scheme; keep as-is
        if (axis.charset === "Joyo") {
            return AXIS_FAMILY.baseName + axis.width + "Joyo" + "-";
        }

        // 文字セット：AXIS は Std / Pro のみ。Pr5 / Pr6 や keep は現状維持 / AXIS only has Std / Pro
        var charset = axis.charset;
        if (charsetMode === "Std" || charsetMode === "Pro") charset = charsetMode;

        // N 切り替え（文字セットがある場合のみ意味を持つ）/ N toggle (only meaningful with a charset)
        var hasN = axis.hasN;
        if (nMode === "on") hasN = true;
        if (nMode === "off") hasN = false;

        var charsetCore = charset ? (charset + (hasN ? "N" : "")) : "";
        return AXIS_FAMILY.baseName + axis.width + charsetCore + "-";
    }

    /* 旧フォント名を分解（必要なら G-OTF 学参を A-OTF に読み替えて再試行）/ Decompose an old name (retry as A-OTF for G-OTF Gakusan) */
    function parseFontName(name) {
        var parsed = parseKnownFamily(name);
        if (parsed) return parsed;

        if (integrateGakusan) {
            for (var prefixIndex = 0; prefixIndex < GAKUSAN_PREFIXES.length; prefixIndex++) {
                var gakusanPrefix = GAKUSAN_PREFIXES[prefixIndex];
                if (name.indexOf(gakusanPrefix) === 0) {
                    parsed = parseKnownFamily(name.substring(gakusanPrefix.length));
                    if (parsed) return parsed;
                }
            }
        }
        return null;
    }

    /* 既知ファミリーの正規表現で旧名を P/UD/charset/N/weight に分解 / Decompose against known family regexes */
    function parseKnownFamily(name) {
        for (var i = 0; i < FONT_FAMILIES.length; i++) {
            var family = FONT_FAMILIES[i];
            var matched = name.match(family.regex);
            if (!matched) continue;

            var parsed = {
                family: family,
                isProportional: false,
                isUD: false,
                isNT: false,
                charset: "",
                hasN: false,
                weight: ""
            };

            // 最後のグループをウエイト、それ以外は値で判別 / Last group is the weight; others by value
            var lastIndex = matched.length - 1;
            parsed.weight = matched[lastIndex];

            for (var groupIndex = 1; groupIndex < lastIndex; groupIndex++) {
                var groupValue = matched[groupIndex];
                if (!groupValue) continue;
                if (groupValue === "P") parsed.isProportional = true;
                else if (groupValue === "UD") parsed.isUD = true;
                else if (groupValue === "NT") parsed.isNT = true;
                else if (groupValue === "N") parsed.hasN = true;
                else if (groupValue === "Pro" || groupValue === "Pr5" || groupValue === "Pr6" || groupValue === "Std") parsed.charset = groupValue;
            }
            return parsed;
        }
        return null;
    }

    /* 設定に基づき変換後フォント名の先頭（ウエイト除く）を組み立て / Build the converted name head (without weight) */
    function buildConvertedNameHead(parsed) {
        var family = parsed.family;
        var supportsProportional = family.regex.source.indexOf("(P)") !== -1;
        var supportsUD = family.regex.source.indexOf("(UD)") !== -1;
        var isCID = (parsed.charset === "");

        // CID（文字セット表記なし）の扱い / Handle CID (no charset token)
        if (isCID && !CONVERT_CID_TO_OTF) {
            return null; // 変換しない
        }

        // 文字セット / Character set
        var charset;
        if (charsetMode !== "keep") {
            charset = charsetMode;
        } else if (isCID) {
            charset = CID_FALLBACK_CHARSET; // 現状維持かつ CID なら Pr6 を補完
        } else {
            charset = parsed.charset;
        }

        // P / Proportional（対応ファミリーのみ）/ Proportional (supported families only)
        var isProportional = parsed.isProportional;
        if (pMode === "on") isProportional = true;
        if (pMode === "off") isProportional = false;
        if (!supportsProportional) isProportional = false;

        // UD（対応ファミリーのみ）/ UD (supported families only)
        var isUD = parsed.isUD;
        if (udMode === "on") isUD = true;
        if (udMode === "off") isUD = false;
        if (!supportsUD) isUD = false;

        // NT（対応ファミリー＝新ゴ のみ）/ NT (only families that support it, i.e. ShinGo)
        var supportsNT = family.regex.source.indexOf("(NT)") !== -1;
        var isNT = parsed.isNT;
        if (ntMode === "on") isNT = true;
        if (ntMode === "off") isNT = false;
        if (!supportsNT) isNT = false;

        // N
        var hasN = parsed.hasN;
        if (nMode === "on") hasN = true;
        if (nMode === "off") hasN = false;

        var prefix = "";
        if (isProportional) prefix += "P";
        if (isUD) prefix += "UD";

        var charsetCore = "";
        if (charset) {
            charsetCore = charset + (hasN ? "N" : "");
        }

        // 基幹名（新ゴは NT 有無を付与）/ Base name (append NT for ShinGo when enabled)
        var baseName = family.baseName + (isNT ? "NT" : "");

        return prefix + baseName + charsetCore + "-";
    }

    /* 同名ウエイトが無い場合に近いウエイトを探す / Find the nearest available weight */
    function findNearestWeight(nameHead, weight) {
        var weightIndex = indexOfArray(WEIGHT_ORDER, weight);
        if (weightIndex < 0) return null;

        for (var distance = 1; distance < WEIGHT_ORDER.length; distance++) {
            var heavierIndex = weightIndex + distance;
            var lighterIndex = weightIndex - distance;
            if (heavierIndex < WEIGHT_ORDER.length && fontExists(nameHead + WEIGHT_ORDER[heavierIndex])) {
                return WEIGHT_ORDER[heavierIndex];
            }
            if (lighterIndex >= 0 && fontExists(nameHead + WEIGHT_ORDER[lighterIndex])) {
                return WEIGHT_ORDER[lighterIndex];
            }
        }
        return null;
    }

    // =========================================
    // フォント存在判定 / Font availability
    // =========================================

    /* 指定名（PostScript 名）のフォントがインストール済みか / Whether a font (by PostScript name) is installed */
    function fontExists(name) {
        return installedFontByPs.hasOwnProperty(name);
    }

    /* インストール済み Font オブジェクトを取得 / Get an installed Font object by PostScript name */
    function getFontObject(name) {
        return installedFontByPs[name];
    }

    // =========================================
    // 確認ダイアログ / Confirmation dialog
    // =========================================

    /* 変更プレビューを表示し、各項目を ON/OFF させて結果を返す / Show preview with per-item ON/OFF and return the result */
    function showConfirmDialog() {
        var confirmDialog = new Window("dialog", L("confirm.title"));
        confirmDialog.orientation = "column";
        confirmDialog.alignChildren = "fill";
        // 左右マージンを +10 / Add 10 to left & right margins
        confirmDialog.margins.left += 10;
        confirmDialog.margins.right += 10;

        var itemCheckboxes = []; // {checkbox, oldName}

        // 「変更前」列の幅を全項目でそろえ、→ の位置を一定にする / Fix the "before" column width so arrows line up
        var beforeWidth = computeBeforeColumnWidth(confirmDialog, [directChanges, weightSubChanges]);

        // 一覧はスクロール可能なビューポートに入れる（項目が多くてもボタンが画面外に出ない）
        // Put the list in a scrollable viewport so buttons stay on-screen even with many items
        var MAX_LIST_HEIGHT = 380;
        var listGroup = confirmDialog.add("group");
        listGroup.orientation = "row";
        listGroup.alignChildren = ["fill", "fill"];
        listGroup.spacing = 2;

        var viewport = listGroup.add("group");
        viewport.orientation = "column";
        viewport.alignChildren = ["left", "top"];

        var listContent = viewport.add("group");
        listContent.orientation = "column";
        listContent.alignChildren = ["left", "top"];
        listContent.spacing = 4;

        if (directChanges.length > 0) {
            listContent.add("statictext", undefined, L("confirm.willChange"));
            addChangeCheckboxes(listContent, directChanges, itemCheckboxes, beforeWidth);
        }
        if (weightSubChanges.length > 0) {
            listContent.add("statictext", undefined, L("confirm.nearWeight"));
            addChangeCheckboxes(listContent, weightSubChanges, itemCheckboxes, beforeWidth);
        }
        if (missingChanges.length > 0) {
            listContent.add("statictext", undefined, L("confirm.notInstalled"));
            var missingNames = uniqueArray(extractNewFontNames(missingChanges));
            for (var k = 0; k < missingNames.length; k++) {
                listContent.add("statictext", undefined, "　" + missingNames[k]);
            }
        }

        confirmDialog.add("statictext", undefined, L("confirm.prompt"));

        var confirmButtonRow = confirmDialog.add("group");
        confirmButtonRow.orientation = "row";
        confirmButtonRow.alignment = "right";
        var confirmCancelButton = confirmButtonRow.add("button", undefined, L("button.cancel"), { name: "cancel" });
        var confirmRunButton = confirmButtonRow.add("button", undefined, L("button.run"), { name: "ok" });
        confirmCancelButton.onClick = function () { confirmDialog.close(0); };
        confirmRunButton.onClick = function () { confirmDialog.close(1); };

        // 一覧が高すぎる場合だけスクロールバーを付ける / Add a scrollbar only when the list is too tall
        confirmDialog.layout.layout(true);
        if (listContent.size.height > MAX_LIST_HEIGHT) {
            viewport.maximumSize.height = MAX_LIST_HEIGHT;
            viewport.preferredSize.height = MAX_LIST_HEIGHT;

            var listScrollbar = listGroup.add("scrollbar");
            listScrollbar.preferredSize.width = 16;
            listScrollbar.alignment = ["right", "fill"];
            listScrollbar.minvalue = 0;
            listScrollbar.maxvalue = listContent.size.height - MAX_LIST_HEIGHT;
            listScrollbar.value = 0;
            listScrollbar.stepdelta = 24;
            listScrollbar.jumpdelta = MAX_LIST_HEIGHT;

            confirmDialog.layout.layout(true);
            var listBaseTop = listContent.location.y;
            listScrollbar.onChanging = function () {
                listContent.location.y = listBaseTop - this.value;
            };
        }

        if (confirmDialog.show() !== 1) {
            return { ok: false, selected: null };
        }

        // チェックが ON の項目だけ採用 / Keep only checked items
        var selected = {};
        for (var i = 0; i < itemCheckboxes.length; i++) {
            if (itemCheckboxes[i].checkbox.value) {
                selected[itemCheckboxes[i].oldName] = true;
            }
        }
        return { ok: true, selected: selected };
    }

    /* 変更項目をチェックボックス（既定 ON）として追加 / Add change items as checkboxes (on by default) */
    function addChangeCheckboxes(parent, changes, itemCheckboxes, beforeWidth) {
        for (var i = 0; i < changes.length; i++) {
            var change = changes[i];
            var row = parent.add("group");
            row.orientation = "row";
            row.alignChildren = ["left", "center"];
            row.spacing = 4;

            // 「変更前」はチェックボックス。幅を固定して → の位置をそろえる / "Before" is the checkbox; fixed width aligns the arrow
            var checkbox = row.add("checkbox", undefined, toDisplayFontName(change.oldName));
            checkbox.value = true;
            checkbox.preferredSize.width = beforeWidth;

            var afterLabel = "→ " + toDisplayFontName(change.newName);
            if (change.oldWeight) {
                afterLabel += "（" + change.oldWeight + " → " + change.newWeight + "）";
            }
            row.add("statictext", undefined, afterLabel);

            itemCheckboxes.push({ checkbox: checkbox, oldName: change.oldName });
        }
    }

    /* PostScript 名を和文表示名に変換。未取得時は元の名前 / Convert PostScript name to a Japanese display name (falls back to the PostScript name) */
    function toDisplayFontName(psName) {
        var font = allFontByPs.hasOwnProperty(psName) ? allFontByPs[psName] : null;
        if (!font) return psName;
        try {
            if (currentLanguage === "ja" && font.fullNameNative) return font.fullNameNative;
            if (font.fullName) return font.fullName;
            return String(font.name).replace("\t", " ");
        } catch (e) {
            return psName;
        }
    }

    /* 「変更前」列に必要な最大幅（チェックボックスのボックス幅＋余白を加味）/ Max width needed for the "before" column, incl. checkbox box & margin */
    function computeBeforeColumnWidth(dialog, changeLists) {
        var graphics = dialog.graphics;
        var maxTextWidth = 0;
        for (var a = 0; a < changeLists.length; a++) {
            var list = changeLists[a];
            for (var i = 0; i < list.length; i++) {
                var displayName = toDisplayFontName(list[i].oldName);
                var textWidth = graphics.measureString(displayName, graphics.font).width;
                if (textWidth > maxTextWidth) maxTextWidth = textWidth;
            }
        }
        return maxTextWidth + 30; // チェックボックスのボックス＋すき間 / checkbox box + gap
    }

    /* 選択された（または全件の）変更を変換マップに追加 / Add selected (or all) changes to the name map */
    function addSelectedChanges(changes, nameMap, selectedOldNames) {
        for (var i = 0; i < changes.length; i++) {
            var change = changes[i];
            if (selectedOldNames === null || selectedOldNames[change.oldName]) {
                nameMap[change.oldName] = change.newName;
            }
        }
    }

    /* 変更リストから新フォント名だけを取り出す / Extract new font names from a change list */
    function extractNewFontNames(changes) {
        var newNames = [];
        for (var i = 0; i < changes.length; i++) {
            newNames.push(changes[i].newName);
        }
        return newNames;
    }

    // =========================================
    // 適用 / Apply
    // =========================================

    /* ターゲットにフォント変更を適用（ロックは一時解除）/ Apply font changes to a target */
    function applyChangesToTarget(target, nameMap) {
        var ranges = target.ranges;

        // 先に「文字インデックス範囲 → 新フォント」を収集（適用で textStyleRange が再構成されても安全）
        // Collect (char-index range -> new font) up front; char count is unchanged, so indices stay valid.
        var jobs = [];
        for (var i = 0; i < ranges.length; i++) {
            var range = ranges[i];
            var chars;
            try {
                chars = range.characters;
                if (chars.length === 0) continue;
            } catch (e) { continue; }

            var psName = fontPsNameOf(range.appliedFont);
            if (!psName) continue;
            var newName = nameMap[psName];
            if (!newName) continue;

            var startIndex, endIndex;
            try {
                startIndex = chars[0].index;
                endIndex = chars[chars.length - 1].index;
            } catch (e2) { continue; }

            jobs.push({ start: startIndex, end: endIndex, font: getFontObject(newName) });
        }
        if (jobs.length === 0) return 0;

        // ロックを一時解除 / Temporarily unlock
        var savedLocked = unlockFramesFor(target);

        var changed = false;
        // 途中でエラーが出てもロックを確実に戻すため try/finally で囲む / Wrap in try/finally so locks are always restored
        try {
            var storyChars = target.story.characters;
            for (var j = 0; j < jobs.length; j++) {
                try {
                    storyChars.itemByRange(jobs[j].start, jobs[j].end).appliedFont = jobs[j].font;
                    changed = true;
                } catch (e3) { }
            }
        } finally {
            // ロック状態を復元 / Restore lock states
            restoreLocks(savedLocked);
        }

        // テキストオブジェクト（ストーリー）単位でカウント / Count per story
        return changed ? 1 : 0;
    }

    /* ターゲットのフレーム・レイヤー・祖先グループのロックを一時解除 / Temporarily unlock a target's containers */
    function unlockFramesFor(target) {
        var saved = [];
        if (!includeLocked) return saved;
        for (var i = 0; i < target.frames.length; i++) {
            var frame = target.frames[i];
            pushUnlock(frame, saved);
            try { pushUnlock(frame.itemLayer, saved); } catch (e) { }
            var parent;
            try { parent = frame.parent; } catch (e2) { parent = null; }
            while (parent) {
                var typeName;
                try { typeName = parent.constructor.name; } catch (e3) { break; }
                if (typeName !== "Group") break;
                pushUnlock(parent, saved);
                try { parent = parent.parent; } catch (e4) { break; }
            }
        }
        return saved;
    }

    /* ロックされていれば解除し、復元対象として記録 / Unlock if locked and remember for restore */
    function pushUnlock(container, saved) {
        try {
            if (container.locked) {
                container.locked = false;
                saved.push(container);
            }
        } catch (e) { }
    }

    /* 記録したコンテナのロックを戻す / Restore lock state of remembered containers */
    function restoreLocks(saved) {
        for (var i = 0; i < saved.length; i++) {
            try { saved[i].locked = true; } catch (e) { }
        }
    }

    /* スタイルコレクションにフォント変更を適用 / Apply font changes to a style collection */
    function applyChangesToStyles(styles, nameMap) {
        for (var k = 0; k < styles.length; k++) {
            var psName = styleFontPsName(styles[k]);
            if (!psName) continue;
            var newName = nameMap[psName];
            if (!newName) continue;
            try {
                styles[k].appliedFont = getFontObject(newName);
            } catch (e) { }
        }
    }

    /* 合成フォントの各エントリにフォント変更を適用 / Apply font changes to composite font entries */
    function applyChangesToCompositeFonts(nameMap) {
        var compositeFonts = activeDocument.compositeFonts;
        for (var c = 0; c < compositeFonts.length; c++) {
            var entries = compositeFonts[c].compositeFontEntries;
            for (var e = 0; e < entries.length; e++) {
                var psName = styleFontPsName(entries[e]);
                if (!psName) continue;
                var newName = nameMap[psName];
                if (!newName) continue;
                try {
                    entries[e].appliedFont = getFontObject(newName);
                } catch (eApply) { }
            }
        }
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /* ダイアログにパネルを追加（共通設定）/ Add a panel to the dialog (shared setup) */
    function addPanel(parent, labelPath, spacing) {
        var panel = parent.add("panel", undefined, L(labelPath));
        setupPanel(panel, spacing);
        return panel;
    }

    /* ラジオ 3 択（あり/なし/現状維持）の状態を返す / Return tri-state radio mode (on/off/keep) */
    function radioMode(onRadio, offRadio) {
        if (onRadio.value) return "on";
        if (offRadio.value) return "off";
        return "keep";
    }

    /* 配列中の位置を返す（ES3 互換）/ indexOf for arrays (ES3 compatible) */
    function indexOfArray(arr, value) {
        for (var i = 0; i < arr.length; i++) {
            if (arr[i] === value) return i;
        }
        return -1;
    }

    /* 重複を除いた配列を返す / Return array with duplicates removed */
    function uniqueArray(arr) {
        var seen = {};
        var unique = [];
        for (var i = 0; i < arr.length; i++) {
            if (!seen[arr[i]]) {
                seen[arr[i]] = true;
                unique.push(arr[i]);
            }
        }
        return unique;
    }

})();
