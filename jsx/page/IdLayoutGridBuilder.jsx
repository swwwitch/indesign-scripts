#target indesign

/*

### 概要

アクティブページに版面・タイトルエリア・額縁・列行グリッド・区切り線などを、プレビューを見ながら一括作成します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdLayoutGridBuilder.md

### Overview

Builds the type area, title area, page frame, column and row grids and dividers on the active page in one pass, with a live preview.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdLayoutGridBuilder.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdLayoutGridBuilder";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-25";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdLayoutGridBuilder.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdLayoutGridBuilder.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

var PREVIEW_LAYER_NAME   = "__QuickLayoutPreview__";  /* プレビュー用レイヤー名 / preview layer name */
var TEMP_GRID_LAYER_NAME = "Temp Grid";               /* 仮グリッドを残すレイヤー名 / layer that keeps the temp grid */
var PAGE_FRAME_BLEED_MM  = 3;                         /* 額縁を裁ち落としへ広げる量（mm）/ page frame bleed in mm */
var DEFAULT_FONT_SIZE    = 9.5;                       /* 文字サイズが読めないときの値（pt）/ fallback font size in pt */

/* 長さの初期値（mm。定規の単位に換算して表示）/ Default lengths in mm, shown in ruler units */
var DEFAULT_LENGTHS_MM = {
    footerHeight: 30,  /* コラムエリアの高さ / footer column area height */
    offset:       10,  /* 実コンテンツ領域のオフセット / content region offset */
    columnGap:    10   /* 列の間隔 / column gap */
};

/* 区切り線の線種（ドキュメントの線種名）/ Stroke style names for dividers */
var DASHED_STROKE_STYLE_NAME = "破線 (3 & 2)";
var DOTTED_STROKE_STYLE_NAME = "点線 (1 & 1)";

/* サンプル文に使うフォントの PostScript 名（先に見つかったもの）/ PostScript names for the sample text, first match wins */
var SAMPLE_FONT_POSTSCRIPT_NAMES = ["HiraKakuProN-W3", "HiraKakuPro-W3", "HiraginoSans-W3"];

/* 作成するカラー / Colors created in the document */
var LAYOUT_COLORS = {
    titleFill:  { name: "K25", cmyk: [0, 0, 0, 25] },
    footerFill: { name: "K40", cmyk: [0, 0, 0, 40] },
    pageFrame:  { name: "K30", cmyk: [0, 0, 0, 30] },
    cellFill:   { name: "K10", cmyk: [0, 0, 0, 10] },
    tempGrid:   { name: "LayoutGrid", cmyk: [100, 0, 0, 0] }
};

// =========================================
// 単位換算 / Unit conversion
// =========================================

var MM_PER_POINT = 25.4 / 72;           /* 1pt = 0.352777…mm / millimeters per point */
var Q_PER_POINT  = MM_PER_POINT / 0.25;  /* 1Q（1H）= 0.25mm なので 1pt = 1.41111…Q / Q (and H) per point */

// =========================================
// レイアウト / Layout
// =========================================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS   = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING   = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS    = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING    = 12;                 /* パネル内の要素間隔 / panel spacing */
var TAB_MARGINS      = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */
var AUTO_BUTTON_SIZE = [70, 22];           /* ［自動調整］ボタンの寸法 / auto-adjust button size */

/**
 * ウィンドウの共通設定を適用する
 * @param {Window} targetWindow 対象ウィンドウ
 * @param {number} [spacing] 要素間隔。省略時は WINDOW_SPACING
 * @returns {void}
 */
function setupWindow(targetWindow, spacing) {
    targetWindow.orientation = "column";
    targetWindow.alignChildren = "fill";
    targetWindow.margins = WINDOW_MARGINS;
    targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルの共通設定を適用する
 * @param {Panel} targetPanel 対象パネル
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupPanel(targetPanel, spacing) {
    targetPanel.orientation = "column";
    targetPanel.alignChildren = ["fill", "top"];
    targetPanel.alignment = ["fill", "top"];
    targetPanel.margins = PANEL_MARGINS;
    targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * タブの共通設定を適用する
 * @param {Group} targetTab 対象タブ
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupTab(targetTab, spacing) {
    targetTab.orientation = "column";
    targetTab.alignChildren = ["fill", "top"];
    targetTab.margins = TAB_MARGINS;
    targetTab.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * 行グループの共通設定を適用する
 * @param {Group} targetGroup 対象グループ
 * @param {string} [alignment] 横方向の配置。省略時は "left"
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupRow(targetGroup, alignment, spacing) {
    targetGroup.orientation = "row";
    targetGroup.alignment = [alignment || "left", "center"];  /* 横と天地を対で / Pair horizontal with vertical */
    targetGroup.alignChildren = ["left", "center"];           /* 親の fill 継承を打ち消す / Cancel the inherited fill */
    targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * 入力欄に↑↓キーでの増減を付ける
 * ↑↓で ±1、Shift 併用で ±10（10 の倍数にそろえる）、Option（Alt）併用で ±0.1
 * @param {EditText} editText 対象の入力欄
 * @param {number} [minValue] 許容する最小値。省略時は下限なし（負の値も可）
 * @returns {void}
 */
function changeValueByArrowKey(editText, minValue) {
    editText.addEventListener("keydown", function (event) {
        if (event.keyName !== "Up" && event.keyName !== "Down") return;
        event.preventDefault();

        var keyboard = ScriptUI.environment.keyboardState;
        var direction = (event.keyName === "Up") ? 1 : -1;
        var value = parseFloat(editText.text);
        if (isNaN(value)) value = 0;

        if (keyboard.altKey) {
            value = Math.round((value + direction * 0.1) * 10) / 10;
        } else {
            var step = (keyboard.shiftKey || event.shiftKey) ? 10 : 1;
            value = (direction > 0)
                ? Math.ceil((value + 0.001) / step) * step
                : Math.floor((value - 0.001) / step) * step;
        }
        if (minValue !== undefined && value < minValue) value = minValue;

        editText.text = String(value);
        editText.notify("onChange");
    });

    if (minValue === undefined) return;
    editText.addEventListener("change", function () {
        var value = parseFloat(editText.text);
        if (isNaN(value) || value < minValue) editText.text = String(minValue);
    });
}

// =========================================
// ラベル定義 / Labels
// =========================================

/**
 * UI 言語を判定する
 * @returns {string} "ja" または "en"
 */
function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var currentLang = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "版面とグリッドの作成", en: "Layout Grid Builder" }
    },
    tab: {
        page:          { ja: "ページ", en: "Page" },
        areas:         { ja: "額縁・エリア", en: "Frame & Areas" },
        contentRegion: { ja: "実コンテンツ領域", en: "Content Region" }
    },
    panel: {
        units:         { ja: "単位（定規／線／文字／行送り）", en: "Units (Ruler / Stroke / Text / Leading)" },
        baseText:      { ja: "基本テキスト", en: "Base Text" },
        margin:        { ja: "マージン", en: "Margins" },
        pageFrame:     { ja: "額縁", en: "Page Frame" },
        typeArea:      { ja: "版面の罫線", en: "Type Area Border" },
        titleArea:     { ja: "タイトルエリア", en: "Title Area" },
        footerArea:    { ja: "フッターのコラムエリア", en: "Footer Column Area" },
        offset:        { ja: "オフセット", en: "Offset" },
        rowCol:        { ja: "列・行", en: "Columns / Rows" },
        cells:         { ja: "セル", en: "Cells" },
        divider:       { ja: "区切り線", en: "Dividers" }
    },
    fieldLabel: {
        cornerRadius: { ja: "角丸", en: "Corner Radius" },
        extension:    { ja: "伸縮", en: "Extension" },
        capStyle:     { ja: "線端", en: "Cap" },
        position:     { ja: "位置", en: "Position" },
        relative:     { ja: "相対", en: "Relative" },
        columnCount:  { ja: "列数", en: "Columns" },
        rowCount:     { ja: "行数", en: "Rows" },
        gap:          { ja: "間隔", en: "Gap" },
        fontSize:     { ja: "サイズ", en: "Size" },
        leading:      { ja: "行送り", en: "Leading" },
        height:       { ja: "高さ", en: "Height" },
        footerGap:    { ja: "アキ", en: "Gap" },
        tempGrid:     { ja: "仮グリッド", en: "Temp Grid" }
    },
    side: {
        top:    { ja: "天", en: "Top" },
        bottom: { ja: "地", en: "Bottom" },
        left:   { ja: "左", en: "Left" },
        right:  { ja: "右", en: "Right" }
    },
    checkbox: {
        border:        { ja: "罫線", en: "Border" },
        fill:          { ja: "塗り", en: "Fill" },
        drawBorder:    { ja: "罫線を描画", en: "Draw border" },
        drawPageFrame: { ja: "額縁を描画", en: "Draw page frame" },
        applyMargins:  { ja: "ページのマージンにも反映", en: "Apply to page margins" },
        bleed:         { ja: "裁ち落とし", en: "Use bleed" },
        frameCorner:   { ja: "角丸", en: "Corner radius" },
        drawDividers:  { ja: "区切り線を描画", en: "Draw dividers" },
        link:          { ja: "連動", en: "Link" },
        showTempGrid:  { ja: "表示", en: "Show" },
        keepTempGrid:  { ja: "残す", en: "Keep" },
        threadFrames:  { ja: "フレームを連結", en: "Thread frames" }
    },
    radio: {
        capNone:        { ja: "なし", en: "None" },
        capRound:       { ja: "丸型", en: "Round" },
        capProjecting:  { ja: "突出", en: "Projecting" },
        positionTop:    { ja: "上", en: "Top" },
        positionBottom: { ja: "下", en: "Bottom" },
        positionLeft:   { ja: "左", en: "Left" },
        positionRight:  { ja: "右", en: "Right" },
        cellFill:       { ja: "塗り", en: "Fill" },
        cellTextFrame:  { ja: "テキストフレーム", en: "Text Frame" },
        sampleNone:     { ja: "なし", en: "None" },
        sampleProse:    { ja: "サンプル文", en: "Sample text" },
        sampleDummy:    { ja: "ダミー文字", en: "Placeholder" },
        lineSolid:      { ja: "実線", en: "Solid" },
        lineDashed:     { ja: "破線", en: "Dashed" },
        lineDotted:     { ja: "点線", en: "Dotted" }
    },
    unit: {
        characters: { ja: "文字", en: "chars" }
    },
    button: {
        ok:            { ja: "OK", en: "OK" },
        cancel:        { ja: "キャンセル", en: "Cancel" },
        autoAdjust:    { ja: "自動調整", en: "Auto" },
        autoAdjustAll: { ja: "すべて自動調整", en: "Auto All" }
    },
    tooltip: {
        units: {
            ja: "線幅と基本テキストを入力する単位を切り替えます。定規の単位はドキュメントの設定に従います。",
            en: "Switches the units for stroke weights and base text. The ruler unit follows the document setting."
        },
        borderWeight: { ja: "版面とタイトルエリアの罫線の線幅です。", en: "Stroke weight for the type area and title area borders." },
        cornerRadius: {
            ja: "版面の罫線（［伸縮］が 0 のとき）とタイトルエリアの塗りに使います。",
            en: "Used for the type area border (when Extension is 0) and the title area fill."
        },
        extension: {
            ja: "0 のときは長方形で描きます。正の値で四辺の罫線を角から外へ伸ばし、負の値で内へ縮めます。",
            en: "At 0 the border is drawn as a rectangle. Positive values extend each side past the corners; negative values shorten them."
        },
        capStyle:    { ja: "［伸縮］が 0 以外のときに使えます。", en: "Available when Extension is not 0." },
        titleStroke: { ja: "線幅は［版面の罫線］と同じです。", en: "Uses the same weight as Type Area Border." },
        titleLength: {
            ja: "［位置］が左・右のときは、タイトルエリアの幅として使います。",
            en: "When Position is Left or Right, this is used as the width of the title area."
        },
        titleExtension:  { ja: "タイトルエリアの罫線を両端から伸ばす量です。", en: "How far the title area border extends past both ends." },
        autoTitleLength: { ja: "仮グリッドの線に合うように［高さ］を丸めます。", en: "Rounds Height to the nearest temp grid line." },
        footerGap:       { ja: "実コンテンツ領域とコラムエリアのあいだのアキです。", en: "Space between the content region and the column area." },
        autoFooterHeight: {
            ja: "コラムエリアの上端が仮グリッドの線に合うように［高さ］を調整します。",
            en: "Adjusts Height so the top of the column area sits on a temp grid line."
        },
        showTempGrid: {
            ja: "基本テキストの文字サイズと行送りから求めた行の位置に、水色の線をプレビューで表示します。",
            en: "Previews cyan lines at the line positions given by the base text size and leading."
        },
        keepTempGrid: {
            ja: "［OK］のあと、印刷されない「Temp Grid」レイヤーに仮グリッドを残します。",
            en: "After OK, keeps the temp grid on a non-printing \"Temp Grid\" layer."
        },
        relative: {
            ja: "入力した値の増減分を、天地左右のマージンにまとめて加えます。",
            en: "Adds the change in this value to all four margins."
        },
        drawPageFrame: { ja: "ページの周囲を K30 の塗りで囲みます。", en: "Surrounds the page with a K30 fill." },
        applyMargins: {
            ja: "［OK］のとき、このページの「マージン・段組」のマージンも同じ値に変更します。",
            en: "On OK, also sets this page's Margins and Columns margins to these values."
        },
        frameCorner: {
            ja: "額縁の内側（くり抜いた部分）の角を丸めます。外側の角は丸めません。",
            en: "Rounds the corners of the frame opening. The outer corners stay square."
        },
        bleed: {
            ja: "額縁の外側を 3mm 外へ広げます。見開きの内側には広げません。",
            en: "Extends the outer edge of the frame by 3 mm. The spine side of a spread is not extended."
        },
        linkSides: { ja: "天地左右を同じ値にそろえます。", en: "Keeps all four sides at the same value." },
        autoOffset: {
            ja: "左右は文字サイズの倍数に、天地は仮グリッドの線に合わせて調整します。",
            en: "Rounds left and right to multiples of the font size, and top and bottom to temp grid lines."
        },
        characterCount: {
            ja: "1 列に入る文字数です。変更すると列の間隔を計算し直します。",
            en: "Characters per column. Changing it recalculates the column gap."
        },
        autoColumnGap: { ja: "［文字］の値に合わせて列の間隔を計算し直します。", en: "Recalculates the column gap from the character count." },
        linkGaps:      { ja: "列と行の間隔を同じ値にそろえます。", en: "Keeps the column and row gaps the same." },
        threadFrames:  { ja: "作成したテキストフレームを順に連結します。", en: "Threads the new text frames in order." },
        sampleText: {
            ja: "［OK］のあと、最初のフレームに流し込みます（プレビューでは流し込みません）。",
            en: "Placed into the first frame after OK (not shown in the preview)."
        },
        dividers: {
            ja: "列・行の間隔の中央に線を引きます。間隔が 0 のときは使えません。",
            en: "Draws lines centered in the column and row gaps. Unavailable when both gaps are 0."
        },
        dividerStyle: {
            ja: "ドキュメントの線種「破線 (3 & 2)」「点線 (1 & 1)」を使います。",
            en: "Uses the document stroke styles \"破線 (3 & 2)\" and \"点線 (1 & 1)\"."
        },
        autoAdjustAll: { ja: "各パネルの［自動調整］をまとめて実行します。", en: "Runs every Auto button in the dialog." }
    },
    alert: {
        noDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document before running this script." },
        strokeStyleMissing: {
            ja: "線種「{name}」がドキュメントにありません。実線で描画します。",
            en: "The stroke style \"{name}\" was not found in the document. Solid lines will be drawn instead."
        }
    }
};

/**
 * ドット区切りキーでラベルを取得する
 * @param {string} labelKey 例: "panel.margin"
 * @returns {string} 現在の言語のラベル。見つからない場合はキーをそのまま返す
 */
function getLabel(labelKey) {
    var labelNode = LABELS;
    var keyParts = labelKey.split(".");
    for (var i = 0; i < keyParts.length; i++) {
        labelNode = labelNode[keyParts[i]];
        if (!labelNode) return labelKey;
    }
    return labelNode[currentLang] || labelNode.en || labelKey;
}

/**
 * 項目名に言語別のコロンを付ける（日本語は全角、英語は半角）
 * @param {string} labelKey 例: "fieldLabel.gap"
 * @returns {string} コロン付きのラベル
 */
function labelText(labelKey) {
    return getLabel(labelKey) + (currentLang === "ja" ? "：" : ":");
}

/**
 * ラベルの末尾に単位を括弧付きで添える
 * @param {string} labelKey 例: "panel.margin"
 * @param {string} unitLabel 表示する単位名
 * @returns {string} 単位付きのラベル
 */
function withUnit(labelKey, unitLabel) {
    return getLabel(labelKey) + " (" + unitLabel + ")";
}

// =========================================
// サンプル文 / Sample text
// =========================================

var SAMPLE_PROSE_TEXT = "朝、目が覚めると、枕元の端末が静かに光っていた。\r「おはようございます。昨日の記憶を同期しますか？」\r\r　私はしばらくその表示を見つめた。\r　同期ボタンは、もう三日間押していない。\r\r　窓の外には、相変わらず同じ街が広がっている。\r　高層ビルの壁面には、朝のニュースが流れていた。\r\r「政府は本日、記憶バックアップ制度の利用率が国民の92%に達したと発表しました」\r\r　人々は、もうほとんど忘れない。\r　毎晩、脳内の記憶はクラウドに保存される。事故でも病気でも、バックアップから復元できる。\r\r　昨日までの自分を、正確に続きから生きられる。\r\r　便利な世界だ。\r\r　私は端末を伏せて、キッチンへ向かった。\r　コーヒーを淹れていると、壁のディスプレイが自動で点灯する。\r\r「未同期の記憶があります」\r\r　分かっている。\r\r　その記憶のせいだ。\r\r　昨日、私は一人の老人に会った。\r\r　河川敷のベンチで、古い紙の本を読んでいた。\r　今どき珍しい。\r\r「それ、オフラインの本ですか？」\r\r　私が声をかけると、老人は少し笑った。\r\r「そうだよ。記録に残らないものが好きでね」\r\r　意味が分からなかった。\r\r　記録に残らない？\r　そんなもの、価値があるのだろうか。\r\r「今の時代、全部残せるじゃないですか」\r\r　私が言うと、老人は本を閉じて言った。\r\r「だから残らないものが必要なんだ」\r\r　風が吹いた。\r　河川敷の草が揺れる。\r\r「人はね、本当は忘れる生き物なんだよ」\r\r　私は黙っていた。\r\r「忘れるから、また会いたくなる。忘れるから、思い出になる」\r\r　老人は空を見上げた。\r\r「全部残るなら、人生はただのログだ」\r\r　ログ。\r\r　その言葉が、妙に頭に残った。\r\r　家に帰ってから、私は同期を押せなかった。\r\r　もし同期すれば、この会話は永久に保存される。\r　政府のサーバーにも、医療記録にも、私の人生ログにも。\r\r　そしてきっと、忘れられなくなる。\r\r　私は端末をもう一度見る。\r\r「記憶同期を実行しますか？」\r\r　画面の下に、小さく表示されている。\r\r「同期しない記憶は、時間とともに消失する可能性があります」\r\r　それでいい。\r\r　私は河川敷の風を思い出す。\r　老人の声を思い出す。\r\r　でも、きっと少しずつ薄れていく。\r\r　声の高さも。\r　顔の皺も。\r　本の色も。\r\r　いつか曖昧になる。\r\r　それでいいのだと思う。\r\r　私は端末の通知を閉じた。\r\r　しばらくして、端末が静かに言う。\r\r「未同期記憶の自動削除まで、残り23時間」\r\r　窓の外では、ドローンが郵便物を運んでいた。\r　街は今日も、正確に記録されている。\r\r　私はコーヒーを飲みながら、ふと思う。\r\r　もしかしたら、あの老人の顔も。\r　もう、はっきり思い出せない。\r\r　でも、不思議と安心していた。\r\r　その記憶は、私の中だけにある。\r\r　サーバーにも、政府にも、誰のログにも残らない。\r\r　ただ、私の人生のどこかに、少しだけ影響して。\r　そして、静かに消えていく。\r\r　端末の光が消える。\r\r　私は窓を開けた。\r\r　春の風が、部屋に入ってきた。";

/**
 * □□□□○□□□□● を繰り返したダミー文字を作る（5〜10 回ごとに改行）
 * @returns {string} ダミー文字列
 */
function buildDummyText() {
    var unitPattern = "□□□□○□□□□●";
    var dummyText = "";
    var repeatsInLine = 0;
    var repeatsUntilBreak = Math.floor(Math.random() * 6) + 5;
    for (var i = 0; i < 200; i++) {
        dummyText += unitPattern;
        repeatsInLine++;
        if (repeatsInLine >= repeatsUntilBreak) {
            dummyText += "\r";
            repeatsInLine = 0;
            repeatsUntilBreak = Math.floor(Math.random() * 6) + 5;
        }
    }
    return dummyText;
}

// =========================================
// 共通ユーティリティ / Utilities
// =========================================

/**
 * 定規の単位の表示名を返す
 * @param {MeasurementUnits} measurementUnit 定規の単位
 * @returns {string} 単位名
 */
function getRulerUnitLabel(measurementUnit) {
    switch (measurementUnit) {
        case MeasurementUnits.MILLIMETERS:     return "mm";
        case MeasurementUnits.CENTIMETERS:     return "cm";
        case MeasurementUnits.INCHES:
        case MeasurementUnits.INCHES_DECIMAL:  return "in";
        case MeasurementUnits.PICAS:           return "p";
        case MeasurementUnits.PIXELS:          return "px";
        case MeasurementUnits.CICEROS:         return "c";
        case MeasurementUnits.AGATES:          return "ag";
        case MeasurementUnits.AMERICAN_POINTS: return "ap";
        case MeasurementUnits.Q:               return "Q";
        case MeasurementUnits.HA:              return "H";
        default:                               return "pt";
    }
}

/**
 * ページの pt 寸法と定規の単位での寸法を比べて、1 単位あたりのポイント数を実測する
 * （換算表を持たないので、どの単位でもドキュメントと食い違わない）
 * @param {Page} targetPage 対象ページ
 * @returns {{horizontal: number, vertical: number}} 横・縦それぞれの 1 単位あたりのポイント数
 */
function measurePointsPerUnit(targetPage) {
    var topLeft = targetPage.resolve(AnchorPoint.TOP_LEFT_ANCHOR, CoordinateSpaces.SPREAD_COORDINATES)[0];
    var bottomRight = targetPage.resolve(AnchorPoint.BOTTOM_RIGHT_ANCHOR, CoordinateSpaces.SPREAD_COORDINATES)[0];
    var bounds = targetPage.bounds;  /* [上, 左, 下, 右]。縦は縦の単位、横は横の単位 / vertical and horizontal units */
    return {
        horizontal: (bottomRight[0] - topLeft[0]) / (bounds[3] - bounds[1]),
        vertical: (bottomRight[1] - topLeft[1]) / (bounds[2] - bounds[0])
    };
}

/**
 * mm の長さを定規の単位に換算する
 * @param {number} millimeters 長さ（mm）
 * @param {number} pointsPerUnit 1 単位あたりのポイント数
 * @returns {number} 定規の単位での長さ
 */
function millimetersToUnits(millimeters, pointsPerUnit) {
    return millimeters / MM_PER_POINT / pointsPerUnit;
}

/**
 * ポイント値を単位付き文字列にする（素の数値はドキュメントの単位で解釈されるため）
 * @param {number} points ポイント値
 * @returns {string} 例: "0.3pt"
 */
function toPointString(points) {
    return String(points) + "pt";
}

/**
 * 小数点以下の桁数を指定して丸める
 * @param {number} value 対象の値
 * @param {number} digits 小数点以下の桁数
 * @returns {number} 丸めた値
 */
function roundTo(value, digits) {
    var scale = Math.pow(10, digits);
    return Math.round(value * scale) / scale;
}

/**
 * 文字列を数値にし、読めなければ代わりの値を返す（0 はそのまま 0）
 * @param {string} text 入力文字列
 * @param {number} fallback 数値にならないときの値
 * @returns {number} 数値
 */
function parseNumberOr(text, fallback) {
    var value = parseFloat(text);
    return isNaN(value) ? fallback : value;
}

/**
 * 配列に値が含まれるか調べる（ES3 に indexOf がないため）
 * @param {Array<string>} values 配列
 * @param {string} target 探す値
 * @returns {boolean} 含まれていれば true
 */
function arrayContains(values, target) {
    for (var i = 0; i < values.length; i++) {
        if (values[i] === target) return true;
    }
    return false;
}

/**
 * 範囲 [上, 左, 下, 右] を内側へ縮める
 * @param {Array<number>} bounds 元の範囲
 * @param {number} top 上から縮める量
 * @param {number} left 左から縮める量
 * @param {number} bottom 下から縮める量
 * @param {number} right 右から縮める量
 * @returns {Array<number>} 縮めた範囲
 */
function insetBounds(bounds, top, left, bottom, right) {
    return [bounds[0] + top, bounds[1] + left, bounds[2] - bottom, bounds[3] - right];
}

/**
 * 範囲からタイトルエリアを差し引く
 * @param {Array<number>} bounds 版面 [上, 左, 下, 右]
 * @param {boolean} titleOn タイトルエリアを描くか
 * @param {number} titleLength タイトルエリアの長さ
 * @param {string} titlePosition "top" / "bottom" / "left" / "right"
 * @returns {Array<number>} 差し引いた範囲
 */
function subtractTitleArea(bounds, titleOn, titleLength, titlePosition) {
    var result = bounds.slice(0);
    if (!titleOn || !(titleLength > 0)) return result;
    if (titlePosition === "top") result[0] += titleLength;
    else if (titlePosition === "bottom") result[2] -= titleLength;
    else if (titlePosition === "left") result[1] += titleLength;
    else if (titlePosition === "right") result[3] -= titleLength;
    return result;
}

/**
 * 範囲の下からフッターのコラムエリア（高さ＋アキ）を差し引く
 * @param {Array<number>} bounds 元の範囲 [上, 左, 下, 右]
 * @param {boolean} footerOn コラムエリアを描くか
 * @param {number} footerHeight コラムエリアの高さ
 * @param {number} footerGap 実コンテンツ領域とのアキ
 * @returns {Array<number>} 差し引いた範囲
 */
function subtractFooterArea(bounds, footerOn, footerHeight, footerGap) {
    var result = bounds.slice(0);
    if (footerOn && footerHeight > 0) result[2] -= (footerHeight + footerGap);
    return result;
}

/**
 * 版面上端からの距離を、最も近い仮グリッドの線の位置に丸める
 * 1 本目は上端から文字サイズぶん下、以降は行送りの間隔
 * @param {number} distance 版面上端からの距離
 * @param {{fontSize: number, leading: number}} fontMetrics 文字サイズと行送り（定規の単位）
 * @returns {number} 丸めた距離
 */
function snapToTempGrid(distance, fontMetrics) {
    var lineIndex = Math.round((distance - fontMetrics.fontSize) / fontMetrics.leading);
    if (lineIndex < 0) lineIndex = 0;
    return fontMetrics.fontSize + lineIndex * fontMetrics.leading;
}

// =========================================
// レイヤーとカラー / Layers and colors
// =========================================

/**
 * 印刷しない作業用レイヤーを取得する（なければ作成）
 * @param {Document} doc 対象ドキュメント
 * @param {string} layerName レイヤー名
 * @returns {Layer} 表示・ロック解除済みのレイヤー
 */
function getOrCreateWorkLayer(doc, layerName) {
    var layer = doc.layers.itemByName(layerName);
    if (!layer.isValid) layer = doc.layers.add({ name: layerName });
    layer.visible = true;
    layer.locked = false;
    layer.printable = false;
    return layer;
}

/**
 * 作業用レイヤーがアクティブなら、ほかのレイヤーに切り替える
 * @param {Document} doc 対象ドキュメント
 * @param {Array<string>} excludedNames アクティブにしないレイヤー名
 * @returns {void}
 */
function activateOtherLayer(doc, excludedNames) {
    if (!arrayContains(excludedNames, doc.activeLayer.name)) return;
    for (var i = 0; i < doc.layers.length; i++) {
        var layer = doc.layers[i];
        if (arrayContains(excludedNames, layer.name)) continue;
        try {
            doc.activeLayer = layer;
        } catch (e) {
            /* 切り替えられなくても描画は続ける / Keep drawing even if the switch fails */
        }
        return;
    }
}

/**
 * プレビューレイヤー上のオブジェクトをすべて削除する
 * @param {Document} doc 対象ドキュメント
 * @returns {void}
 */
function clearPreviewLayer(doc) {
    var layer = doc.layers.itemByName(PREVIEW_LAYER_NAME);
    if (!layer.isValid) return;
    layer.locked = false;
    layer.visible = true;
    for (var i = layer.pageItems.length - 1; i >= 0; i--) {
        layer.pageItems[i].remove();
    }
}

/**
 * プレビューレイヤーを削除する
 * @param {Document} doc 対象ドキュメント
 * @returns {void}
 */
function removePreviewLayer(doc) {
    var layer = doc.layers.itemByName(PREVIEW_LAYER_NAME);
    if (!layer.isValid) return;
    activateOtherLayer(doc, [PREVIEW_LAYER_NAME]);
    layer.remove();
}

/**
 * 名前付きのカラーを取得する（なければ作成）
 * @param {Document} doc 対象ドキュメント
 * @param {{name: string, cmyk: Array<number>}} colorDef カラーの定義
 * @returns {Color} カラー
 */
function getOrCreateColor(doc, colorDef) {
    var color = doc.colors.itemByName(colorDef.name);
    if (color.isValid) return color;
    return doc.colors.add({
        name: colorDef.name,
        model: ColorModel.PROCESS,
        space: ColorSpace.CMYK,
        colorValue: colorDef.cmyk
    });
}

// =========================================
// 描画 / Drawing
// =========================================

var ALL_CORNERS = ["topLeft", "topRight", "bottomLeft", "bottomRight"];

/* タイトルエリアの位置ごとに角丸にする角 / Corners rounded for each title position */
var TITLE_AREA_CORNERS = {
    top:    ["topLeft", "topRight"],
    bottom: ["bottomLeft", "bottomRight"],
    left:   ["topLeft", "bottomLeft"],
    right:  ["topRight", "bottomRight"]
};

/**
 * 設定値に従ってレイアウト要素を描画する
 * 縦の単位を横にそろえ、原点をスプレッドにしてから描き、最後に元へ戻す
 * @param {object} settings readSettings() が返す設定値
 * @returns {void}
 */
function drawLayout(settings) {
    var doc = app.activeDocument;
    var viewPrefs = doc.viewPreferences;
    var savedVerticalUnits = viewPrefs.verticalMeasurementUnits;
    var savedRulerOrigin = viewPrefs.rulerOrigin;
    var savedZeroPoint = doc.zeroPoint;
    var savedActiveLayer = null;

    viewPrefs.verticalMeasurementUnits = viewPrefs.horizontalMeasurementUnits;
    viewPrefs.rulerOrigin = RulerOrigin.SPREAD_ORIGIN;  /* 見開きの右ページにも対応 / Works on right-hand pages too */
    doc.zeroPoint = [0, 0];

    try {
        if (settings.targetLayer && settings.targetLayer.isValid) {
            savedActiveLayer = doc.activeLayer;
            doc.activeLayer = settings.targetLayer;
        }

        var page = app.activeWindow.activePage;
        var drawing = {
            doc: doc,
            page: page,
            settings: settings,
            regions: computeDrawRegions(page.bounds, settings),
            blackSwatch: doc.swatches.item("Black"),
            noneSwatch: doc.swatches.item("None")
        };

        if (!settings.gridOnly) {
            if (settings.typeAreaBorder) drawTypeAreaBorder(drawing);
            if (settings.titleLength > 0) drawTitleArea(drawing);
            if (drawing.regions.footerBounds) drawFooterArea(drawing);
            if (settings.pageFrame) drawPageFrame(drawing);
            if (settings.cellFill || settings.cellTextFrame) drawCells(drawing);
            if (settings.dividers && (settings.colCount > 1 || settings.rowCount > 1)) drawDividers(drawing);
        }
        if (settings.showTempGrid && (settings.targetLayer || settings.gridOnly)) drawTempGrid(drawing);
    } finally {
        if (savedActiveLayer && savedActiveLayer.isValid) {
            try {
                doc.activeLayer = savedActiveLayer;
            } catch (e) {
                /* 戻せなくても単位と原点の復元は続ける / Still restore units and origin */
            }
        }
        viewPrefs.rulerOrigin = savedRulerOrigin;
        doc.zeroPoint = savedZeroPoint;
        viewPrefs.verticalMeasurementUnits = savedVerticalUnits;
    }
}

/**
 * 描画に使う各領域を求める
 * @param {Array<number>} pageBounds ページ [上, 左, 下, 右]
 * @param {object} settings 設定値
 * @returns {object} page / typeArea / content / footerBounds / borderBottom / grid
 */
function computeDrawRegions(pageBounds, settings) {
    var typeArea = insetBounds(pageBounds, settings.marginTop, settings.marginLeft, settings.marginBottom, settings.marginRight);
    var footerOn = (settings.footerFill || settings.footerStroke) && settings.footerHeight > 0;
    var afterTitle = subtractTitleArea(typeArea, settings.titleFill || settings.titleStroke, settings.titleLength, settings.titlePosition);
    var content = subtractFooterArea(afterTitle, footerOn, settings.footerHeight, settings.footerGap);
    if (content[2] < content[0]) content[2] = content[0];

    /* フッターのコラムエリア（実コンテンツ領域の下端＋アキから）/ Footer column area below the content region */
    var footerBounds = null;
    if (footerOn) {
        var footerTop = content[2] + settings.footerGap;
        var footerBottom = footerTop + settings.footerHeight;
        if (footerTop < content[2]) footerTop = content[2];
        if (footerBottom > afterTitle[2]) footerBottom = afterTitle[2];
        if (footerBottom > footerTop && content[3] > content[1]) {
            footerBounds = [footerTop, content[1], footerBottom, content[3]];
        }
    }

    return {
        page: pageBounds,
        typeArea: typeArea,
        content: content,
        footerBounds: footerBounds,
        /* 版面の罫線はコラムエリア（高さ＋アキ）を除いた範囲 / The border stops above the footer */
        borderBottom: footerOn ? typeArea[2] - settings.footerHeight - settings.footerGap : typeArea[2],
        grid: insetBounds(content, settings.offsetTop, settings.offsetLeft, settings.offsetBottom, settings.offsetRight)
    };
}

/**
 * 線を 1 本引く（既定は黒）
 * @param {object} drawing 描画コンテキスト
 * @param {Array<Array<number>>} points 始点と終点 [[x, y], [x, y]]
 * @param {number} weightPt 線幅（pt）
 * @param {object} [lineOptions] endCap / strokeType / strokeColor
 * @returns {GraphicLine} 作成した線
 */
function addLine(drawing, points, weightPt, lineOptions) {
    var opts = lineOptions || {};
    var line = drawing.page.graphicLines.add();
    line.paths[0].entirePath = points;
    line.strokeWeight = toPointString(weightPt);
    line.strokeColor = opts.strokeColor || drawing.blackSwatch;
    if (opts.endCap) line.endCap = opts.endCap;
    if (opts.strokeType) line.strokeType = opts.strokeType;
    return line;
}

/**
 * 塗りだけの長方形を作り、最背面へ送る
 * @param {object} drawing 描画コンテキスト
 * @param {Array<number>} bounds [上, 左, 下, 右]
 * @param {{name: string, cmyk: Array<number>}} colorDef 塗りのカラー
 * @returns {Rectangle} 作成した長方形
 */
function addFilledRectangle(drawing, bounds, colorDef) {
    var rect = drawing.page.rectangles.add({
        geometricBounds: bounds,
        strokeWeight: 0,
        strokeColor: drawing.noneSwatch,
        fillColor: getOrCreateColor(drawing.doc, colorDef)
    });
    rect.sendToBack();
    return rect;
}

/**
 * 黒の罫線だけの長方形を作る
 * @param {object} drawing 描画コンテキスト
 * @param {Array<number>} bounds [上, 左, 下, 右]
 * @param {number} weightPt 線幅（pt）
 * @returns {Rectangle} 作成した長方形
 */
function addStrokedRectangle(drawing, bounds, weightPt) {
    return drawing.page.rectangles.add({
        geometricBounds: bounds,
        strokeWeight: toPointString(weightPt),
        strokeColor: drawing.blackSwatch,
        fillColor: drawing.noneSwatch
    });
}

/**
 * 長方形の指定した角を角丸にする
 * @param {Rectangle} rect 対象の長方形
 * @param {number} radius 角丸の半径（定規の単位）。0 以下なら何もしない
 * @param {Array<string>} cornerNames "topLeft" などの角の名前
 * @param {number} pointsPerUnit 1 単位あたりのポイント数
 * @returns {void}
 */
function roundCorners(rect, radius, cornerNames, pointsPerUnit) {
    if (!(radius > 0)) return;
    var radiusText = toPointString(radius * pointsPerUnit);
    for (var i = 0; i < cornerNames.length; i++) {
        rect[cornerNames[i] + "CornerOption"] = CornerOptions.ROUNDED_CORNER;
        rect[cornerNames[i] + "CornerRadius"] = radiusText;
    }
}

/**
 * 線端の識別子を EndCap に変換する
 * @param {string} capStyle "none" / "round" / "project"
 * @returns {EndCap} 線端
 */
function toEndCap(capStyle) {
    if (capStyle === "round") return EndCap.ROUND_END_CAP;
    if (capStyle === "project") return EndCap.PROJECTING_END_CAP;
    return EndCap.BUTT_END_CAP;
}

/**
 * 版面の罫線を描く。［伸縮］が 0 なら長方形、それ以外は四辺を別々の線で描く
 * @param {object} drawing 描画コンテキスト
 * @returns {void}
 */
function drawTypeAreaBorder(drawing) {
    var settings = drawing.settings;
    var typeArea = drawing.regions.typeArea;
    var top = typeArea[0];
    var left = typeArea[1];
    var right = typeArea[3];
    var bottom = drawing.regions.borderBottom;
    var ext = settings.borderExtension;

    if (ext === 0) {
        var rect = addStrokedRectangle(drawing, [top, left, bottom, right], settings.borderWeight);
        roundCorners(rect, settings.borderCornerRadius, ALL_CORNERS, settings.pointsPerUnit);
        return;
    }

    /* 正なら角から外へ伸ばし、負なら内へ縮める / Positive extends past the corners, negative leaves gaps */
    var lineOptions = { endCap: toEndCap(settings.capStyle) };
    addLine(drawing, [[left - ext, top], [right + ext, top]], settings.borderWeight, lineOptions);
    addLine(drawing, [[left - ext, bottom], [right + ext, bottom]], settings.borderWeight, lineOptions);
    addLine(drawing, [[left, top - ext], [left, bottom + ext]], settings.borderWeight, lineOptions);
    addLine(drawing, [[right, top - ext], [right, bottom + ext]], settings.borderWeight, lineOptions);
}

/**
 * タイトルエリアの塗りと罫線を描く
 * @param {object} drawing 描画コンテキスト
 * @returns {void}
 */
function drawTitleArea(drawing) {
    var settings = drawing.settings;
    var typeArea = drawing.regions.typeArea;
    var top = typeArea[0];
    var left = typeArea[1];
    var bottom = typeArea[2];
    var right = typeArea[3];
    var length = settings.titleLength;
    var ext = settings.titleExtension;
    var position = settings.titlePosition;

    /* 塗り：版面の端から長さぶん。外側の 2 角だけ角丸 / Fill with the two outer corners rounded */
    if (settings.titleFill) {
        var fillBounds = {
            top:    [top, left, top + length, right],
            bottom: [bottom - length, left, bottom, right],
            left:   [top, left, bottom, left + length],
            right:  [top, right - length, bottom, right]
        }[position];
        var rect = addFilledRectangle(drawing, fillBounds, LAYOUT_COLORS.titleFill);
        roundCorners(rect, settings.titleCornerRadius, TITLE_AREA_CORNERS[position], settings.pointsPerUnit);
    }

    /* 罫線：タイトルエリアの内側の辺に 1 本 / One line along the inner edge */
    if (settings.titleStroke) {
        var points;
        if (position === "top" || position === "bottom") {
            var lineY = (position === "top") ? top + length : bottom - length;
            points = [[left - ext, lineY], [right + ext, lineY]];
        } else {
            var lineX = (position === "left") ? left + length : right - length;
            points = [[lineX, top - ext], [lineX, bottom + ext]];
        }
        addLine(drawing, points, settings.borderWeight);
    }
}

/**
 * フッターのコラムエリアの塗りと罫線を描く
 * @param {object} drawing 描画コンテキスト
 * @returns {void}
 */
function drawFooterArea(drawing) {
    var settings = drawing.settings;
    var bounds = drawing.regions.footerBounds;
    if (settings.footerFill) {
        roundCorners(addFilledRectangle(drawing, bounds, LAYOUT_COLORS.footerFill), settings.footerCornerRadius, ALL_CORNERS, settings.pointsPerUnit);
    }
    if (settings.footerStroke) {
        roundCorners(addStrokedRectangle(drawing, bounds, settings.footerWeight), settings.footerCornerRadius, ALL_CORNERS, settings.pointsPerUnit);
    }
}

/**
 * ページの周囲を囲む額縁を、穴あきの多角形で描く
 * @param {object} drawing 描画コンテキスト
 * @returns {void}
 */
function drawPageFrame(drawing) {
    var settings = drawing.settings;
    var pageBounds = drawing.regions.page;
    var outer = pageBounds.slice(0);

    /* 裁ち落とし：天地は常に、左右は見開きの外側だけ / Bleed on the outside edges only */
    if (settings.pageFrameBleed) {
        var pageSide = drawing.page.side;
        var bleed = millimetersToUnits(PAGE_FRAME_BLEED_MM, settings.pointsPerUnit);
        outer[0] -= bleed;
        outer[2] += bleed;
        if (pageSide !== PageSideOptions.RIGHT_HAND) outer[1] -= bleed;
        if (pageSide !== PageSideOptions.LEFT_HAND) outer[3] += bleed;
    }
    var inner = insetBounds(pageBounds, settings.pageFrameTop, settings.pageFrameLeft, settings.pageFrameBottom, settings.pageFrameRight);

    /* 外側を順回り、内側を逆回りにして型抜きにする / Reverse the inner path to cut a hole */
    var framePolygon = drawing.page.polygons.add();
    framePolygon.paths[0].entirePath = [[outer[1], outer[0]], [outer[3], outer[0]], [outer[3], outer[2]], [outer[1], outer[2]]];
    var openingPath = framePolygon.paths.add();
    openingPath.entirePath = buildOpeningPath(inner, settings.pageFrameCornerRadiusPt / settings.pointsPerUnit);
    /* 足したパスは開いたままなので閉じる（閉じないと最後の角が直線でつながり面取りになる）/ Close it, or the last corner becomes a chamfer */
    openingPath.pathType = PathType.CLOSED_PATH;
    framePolygon.fillColor = getOrCreateColor(drawing.doc, LAYOUT_COLORS.pageFrame);
    framePolygon.strokeColor = drawing.noneSwatch;
    framePolygon.strokeWeight = 0;
    framePolygon.sendToBack();
}

/**
 * 額縁のくり抜き部分のパスを、外側と逆回り（左上から下へ）で作る。半径があれば角を曲線で丸める
 * @param {Array<number>} bounds くり抜く範囲 [上, 左, 下, 右]
 * @param {number} radius 角丸の半径（定規の単位）。幅・高さの半分までに抑える
 * @returns {Array} entirePath に渡す点の配列（角丸ありは [入り方向, 基準点, 出方向] の組）
 */
function buildOpeningPath(bounds, radius) {
    var top = bounds[0];
    var left = bounds[1];
    var bottom = bounds[2];
    var right = bounds[3];
    var r = Math.min(radius, (right - left) / 2, (bottom - top) / 2);
    if (!(r > 0)) return [[left, top], [left, bottom], [right, bottom], [right, top]];

    /* 4 分の 1 円をベジェで近似する係数 / Bezier handle length for a quarter circle */
    var handle = r * 0.5522847498;
    /**
     * 直線と曲線のつなぎ目の点を作る
     * @param {number} x 基準点の x
     * @param {number} y 基準点の y
     * @param {Array<number>} inHandle 入り方向の点
     * @param {Array<number>} outHandle 出方向の点
     * @returns {Array<Array<number>>} [入り方向, 基準点, 出方向]
     */
    function point(x, y, inHandle, outHandle) {
        return [inHandle || [x, y], [x, y], outHandle || [x, y]];
    }
    return [
        point(left, top + r, [left, top + r - handle], null),           /* 左辺の上端 / left edge, top */
        point(left, bottom - r, null, [left, bottom - r + handle]),     /* 左辺の下端 / left edge, bottom */
        point(left + r, bottom, [left + r - handle, bottom], null),     /* 下辺の左端 / bottom edge, left */
        point(right - r, bottom, null, [right - r + handle, bottom]),   /* 下辺の右端 / bottom edge, right */
        point(right, bottom - r, [right, bottom - r + handle], null),   /* 右辺の下端 / right edge, bottom */
        point(right, top + r, null, [right, top + r - handle]),         /* 右辺の上端 / right edge, top */
        point(right - r, top, [right - r + handle, top], null),         /* 上辺の右端 / top edge, right */
        point(left + r, top, null, [left + r - handle, top])            /* 上辺の左端 / top edge, left */
    ];
}

/**
 * グリッドを列×行のセルに分けた範囲を求める（列ごとに上から下の順）
 * @param {Array<number>} grid グリッドの範囲 [上, 左, 下, 右]
 * @param {object} settings 設定値
 * @returns {Array<Array<number>>} セルの範囲の配列
 */
function computeCellBounds(grid, settings) {
    var cellWidth = ((grid[3] - grid[1]) - settings.colGap * (settings.colCount - 1)) / settings.colCount;
    var cellHeight = ((grid[2] - grid[0]) - settings.rowGap * (settings.rowCount - 1)) / settings.rowCount;
    var cells = [];
    for (var i = 0; i < settings.colCount; i++) {
        for (var j = 0; j < settings.rowCount; j++) {
            var cellLeft = grid[1] + i * (cellWidth + settings.colGap);
            var cellTop = grid[0] + j * (cellHeight + settings.rowGap);
            cells.push([cellTop, cellLeft, cellTop + cellHeight, cellLeft + cellWidth]);
        }
    }
    return cells;
}

/**
 * セルを塗り、またはテキストフレームで埋める
 * @param {object} drawing 描画コンテキスト
 * @returns {void}
 */
function drawCells(drawing) {
    var settings = drawing.settings;
    var cells = computeCellBounds(drawing.regions.grid, settings);
    var i;

    if (settings.cellFill) {
        for (i = 0; i < cells.length; i++) {
            addFilledRectangle(drawing, cells[i], LAYOUT_COLORS.cellFill);
        }
        return;
    }

    var textFrames = [];
    for (i = 0; i < cells.length; i++) {
        textFrames.push(drawing.page.textFrames.add({
            geometricBounds: cells[i],
            strokeWeight: 0,
            strokeColor: drawing.noneSwatch,
            fillColor: drawing.noneSwatch
        }));
    }
    if (settings.threadFrames) {
        for (i = 0; i < textFrames.length - 1; i++) {
            textFrames[i].nextTextFrame = textFrames[i + 1];
        }
    }
    /* プレビューでは流し込まない / Skip in the preview */
    if ((settings.sampleProse || settings.sampleDummy) && !settings.targetLayer) {
        placeSampleText(textFrames[0], settings);
    }
}

/**
 * テキストフレームにサンプル文かダミー文字を流し込み、基本テキストの書式を当てる
 * @param {TextFrame} textFrame 流し込み先
 * @param {object} settings 設定値
 * @returns {void}
 */
function placeSampleText(textFrame, settings) {
    textFrame.contents = settings.sampleDummy ? buildDummyText() : SAMPLE_PROSE_TEXT;

    var story = textFrame.parentStory;
    if (settings.fontSizePt > 0) story.pointSize = toPointString(settings.fontSizePt);
    if (settings.leading === "auto" || settings.leading === "") {
        story.leading = Leading.AUTO;
    } else {
        var leadingPt = parseFloat(settings.leading);
        if (leadingPt > 0) story.leading = toPointString(leadingPt);
    }

    var sampleFont = findFontByPostScriptName(SAMPLE_FONT_POSTSCRIPT_NAMES);
    if (sampleFont) story.appliedFont = sampleFont;
}

/**
 * PostScript 名の候補から、インストールされているフォントを探す
 * （app.fonts の名前は「ファミリー名＋タブ＋スタイル名」で言語によって変わるため、PostScript 名で照合する）
 * @param {Array<string>} postScriptNames 候補の PostScript 名（先頭ほど優先）
 * @returns {Font|null} 見つかったフォント。なければ null
 */
function findFontByPostScriptName(postScriptNames) {
    var installedNames = app.fonts.everyItem().postscriptName;
    for (var i = 0; i < postScriptNames.length; i++) {
        for (var j = 0; j < installedNames.length; j++) {
            if (installedNames[j] === postScriptNames[i]) return app.fonts[j];
        }
    }
    return null;
}

/**
 * 区切り線の線種を取得する。ドキュメントになければ実線（null）にする
 * @param {Document} doc 対象ドキュメント
 * @param {string} lineType "solid" / "dashed" / "dotted"
 * @param {boolean} shouldWarn 見つからないときに警告するか
 * @returns {StrokeStyle|null} 線種。実線なら null
 */
function findDividerStrokeStyle(doc, lineType, shouldWarn) {
    var styleName = null;
    if (lineType === "dashed") styleName = DASHED_STROKE_STYLE_NAME;
    else if (lineType === "dotted") styleName = DOTTED_STROKE_STYLE_NAME;
    if (!styleName) return null;

    var strokeStyle = doc.strokeStyles.itemByName(styleName);
    if (strokeStyle.isValid) return strokeStyle;
    if (shouldWarn) alert(getLabel("alert.strokeStyleMissing").replace("{name}", styleName));
    return null;
}

/**
 * 列・行の間隔の中央に区切り線を引く
 * @param {object} drawing 描画コンテキスト
 * @returns {void}
 */
function drawDividers(drawing) {
    var settings = drawing.settings;
    var grid = drawing.regions.grid;
    var cellWidth = ((grid[3] - grid[1]) - settings.colGap * (settings.colCount - 1)) / settings.colCount;
    var cellHeight = ((grid[2] - grid[0]) - settings.rowGap * (settings.rowCount - 1)) / settings.rowCount;
    var lineOptions = {
        endCap: EndCap.BUTT_END_CAP,
        /* 警告は確定時だけ（プレビューのたびに出さない）/ Warn only on the final run */
        strokeType: findDividerStrokeStyle(drawing.doc, settings.dividerLineType, !settings.targetLayer)
    };

    /* 列の区切り（縦線）/ Column dividers */
    for (var i = 1; i < settings.colCount; i++) {
        var lineX = grid[1] + i * cellWidth + (i - 0.5) * settings.colGap;
        addLine(drawing, [[lineX, grid[0]], [lineX, grid[2]]], settings.dividerWeight, lineOptions);
    }
    /* 行の区切り（横線）/ Row dividers */
    for (var j = 1; j < settings.rowCount; j++) {
        var lineY = grid[0] + j * cellHeight + (j - 0.5) * settings.rowGap;
        addLine(drawing, [[grid[1], lineY], [grid[3], lineY]], settings.dividerWeight, lineOptions);
    }
}

/**
 * 仮グリッド（行の位置の横線）をページ全幅に描く
 * 版面上端から文字サイズぶん下が 1 本目、以降は行送りの間隔
 * @param {object} drawing 描画コンテキスト
 * @returns {void}
 */
function drawTempGrid(drawing) {
    var settings = drawing.settings;
    var pageBounds = drawing.regions.page;
    var fontSizePt = (settings.fontSizePt > 0) ? settings.fontSizePt : DEFAULT_FONT_SIZE;
    var leadingPt = parseFloat(settings.leading);
    if (isNaN(leadingPt) || leadingPt <= 0) leadingPt = fontSizePt * 1.5;

    var lineOptions = { strokeColor: getOrCreateColor(drawing.doc, LAYOUT_COLORS.tempGrid) };
    var lineStep = leadingPt / settings.pointsPerUnit;
    var gridLines = [];
    for (var lineY = drawing.regions.typeArea[0] + fontSizePt / settings.pointsPerUnit; lineY <= pageBounds[2]; lineY += lineStep) {
        gridLines.push(addLine(drawing, [[pageBounds[1], lineY], [pageBounds[3], lineY]], 0.1, lineOptions));
    }
    if (gridLines.length > 1) drawing.page.groups.add(gridLines);
}

// =========================================
// ダイアログの構築 / Dialog construction
// =========================================

/**
 * 設定ダイアログを組み立てる
 * @param {object} context ドキュメント・ページ・単位・初期値
 * @returns {object} ダイアログとコントロール、入力単位の状態をまとめたオブジェクト
 */
function buildDialog(context) {
    var ui = {
        context: context,
        fontUnitIsQ: false,     /* 基本テキストを Q/H で入力中か / Base text entered in Q/H */
        strokeUnitIsMm: false,  /* 線幅を mm で入力中か / Stroke weights entered in mm */
        lastRelativeValue: 0    /* ［相対］の前回値 / Previous Relative value */
    };

    var dlg = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(dlg);
    ui.dlg = dlg;

    var settingsTabs = dlg.add("tabbedpanel");
    settingsTabs.alignChildren = ["fill", "top"];

    var pageTab = addTab(settingsTabs, getLabel("tab.page"));
    buildPageTab(ui, pageTab);

    var areasTab = addTab(settingsTabs, getLabel("tab.areas"));
    buildPageFramePanel(ui, areasTab);
    buildTitleAreaPanel(ui, areasTab);
    buildFooterAreaPanel(ui, areasTab);

    buildContentRegionTab(ui, addTab(settingsTabs, getLabel("tab.contentRegion")));
    settingsTabs.selection = pageTab;

    buildButtonRow(ui, dlg);
    return ui;
}

/**
 * 共通設定を当てたタブを追加する
 * @param {TabbedPanel} parent タブパネル
 * @param {string} title タブの見出し
 * @returns {Tab} タブ
 */
function addTab(parent, title) {
    var tab = parent.add("tab", undefined, title);
    setupTab(tab);
    return tab;
}

/**
 * 共通設定を当てたパネルを追加する
 * @param {Group|Panel} parent 親
 * @param {string} title パネルの見出し
 * @returns {Panel} パネル
 */
function addPanel(parent, title) {
    var panel = parent.add("panel", undefined, title);
    setupPanel(panel);
    return panel;
}

/**
 * 共通設定を当てた行グループを追加する
 * @param {Group|Panel} parent 親
 * @param {string} [alignment] 横方向の配置
 * @returns {Group} 行グループ
 */
function addRow(parent, alignment) {
    var row = parent.add("group");
    setupRow(row, alignment);
    return row;
}

/**
 * コロン付きの項目名を右揃えで追加する
 * @param {Group} parent 親グループ
 * @param {string} labelKey LABELS のキー
 * @returns {StaticText} 項目名
 */
function addFieldLabel(parent, labelKey) {
    var label = parent.add("statictext", undefined, labelText(labelKey));
    label.justify = "right";
    return label;
}

/**
 * ↑↓キーで増減できる数値入力欄を追加する
 * @param {Group} parent 親グループ
 * @param {number|string} defaultValue 初期値
 * @param {number} characters 表示幅（文字数）
 * @param {number} [minValue] 最小値
 * @returns {EditText} 入力欄
 */
function addNumberInput(parent, defaultValue, characters, minValue) {
    var input = parent.add("edittext", undefined, String(defaultValue));
    input.characters = characters;
    changeValueByArrowKey(input, minValue);
    return input;
}

/**
 * 項目名と入力欄を並べた小さなグループを追加する（行の設定は当てない）
 * @param {Group} parent 親グループ
 * @param {string} labelKey LABELS のキー
 * @param {number|string} defaultValue 初期値
 * @param {number} characters 表示幅（文字数）
 * @returns {EditText} 入力欄
 */
function addLabeledInput(parent, labelKey, defaultValue, characters) {
    var labeledGroup = parent.add("group");
    addFieldLabel(labeledGroup, labelKey);
    return addNumberInput(labeledGroup, defaultValue, characters);
}

/**
 * 項目名・入力欄・単位を 1 行に並べる
 * @param {Group|Panel} parent 親
 * @param {string} labelKey LABELS のキー
 * @param {number|string} defaultValue 初期値
 * @param {number} characters 表示幅（文字数）
 * @param {string|null} unitText 単位の表記。null なら付けない
 * @param {number} [minValue] 最小値
 * @returns {{row: Group, input: EditText, unitLabel: StaticText}} 作成したコントロール
 */
function addNumberRow(parent, labelKey, defaultValue, characters, unitText, minValue) {
    var row = addRow(parent);
    addFieldLabel(row, labelKey);
    var input = addNumberInput(row, defaultValue, characters, minValue);
    var unitLabel = unitText ? row.add("statictext", undefined, unitText) : null;
    return { row: row, input: input, unitLabel: unitLabel };
}

/**
 * ［自動調整］ボタンを追加する
 * @param {Group|Panel} parent 親
 * @param {string} tooltipKey ツールチップのキー
 * @returns {Button} ボタン
 */
function addAutoButton(parent, tooltipKey) {
    var button = parent.add("button", undefined, getLabel("button.autoAdjust"));
    button.preferredSize = AUTO_BUTTON_SIZE;
    button.helpTip = getLabel(tooltipKey);
    return button;
}

/**
 * ラベルキーの一覧からラジオボタンを並べ、先頭を選択しておく
 * @param {Group} parent 親グループ（同じ親の中で排他になる）
 * @param {Array<string>} labelKeys LABELS のキー
 * @param {string} [tooltipKey] すべてに付けるツールチップのキー
 * @returns {Array<RadioButton>} ラジオボタン
 */
function addRadioButtons(parent, labelKeys, tooltipKey) {
    var radios = [];
    for (var i = 0; i < labelKeys.length; i++) {
        var radio = parent.add("radiobutton", undefined, getLabel(labelKeys[i]));
        if (tooltipKey) radio.helpTip = getLabel(tooltipKey);
        radios.push(radio);
    }
    radios[0].value = true;
    return radios;
}

/**
 * チェックボックスを追加する
 * @param {Group|Panel} parent 親
 * @param {string} labelKey LABELS のキー
 * @param {boolean} checked 初期状態
 * @param {string} [tooltipKey] ツールチップのキー
 * @returns {Checkbox} チェックボックス
 */
function addCheckbox(parent, labelKey, checked, tooltipKey) {
    var checkbox = parent.add("checkbox", undefined, getLabel(labelKey));
    checkbox.value = checked;
    if (tooltipKey) checkbox.helpTip = getLabel(tooltipKey);
    return checkbox;
}

/**
 * 天地左右の入力欄を「左｜天・連動・地｜右」の形に並べる
 * @param {Panel} parent 親パネル
 * @param {number} defaultValue 初期値
 * @param {number} characters 表示幅（文字数）
 * @returns {{row: Group, top: EditText, bottom: EditText, left: EditText, right: EditText, linkCheck: Checkbox}} 作成したコントロール
 */
function addLinkedSideInputs(parent, defaultValue, characters) {
    var sidesRow = addRow(parent, "center");
    var left = addLabeledInput(sidesRow, "side.left", defaultValue, characters);

    var topBottomGroup = sidesRow.add("group");
    topBottomGroup.orientation = "column";
    topBottomGroup.alignChildren = "center";
    var top = addLabeledInput(topBottomGroup, "side.top", defaultValue, characters);
    var linkCheck = addCheckbox(topBottomGroup, "checkbox.link", true, "tooltip.linkSides");
    var bottom = addLabeledInput(topBottomGroup, "side.bottom", defaultValue, characters);

    var right = addLabeledInput(sidesRow, "side.right", defaultValue, characters);
    return { row: sidesRow, top: top, bottom: bottom, left: left, right: right, linkCheck: linkCheck };
}

/**
 * ［ページ］タブ（単位・基本テキスト・マージン・版面の罫線）を組み立てる
 * @param {object} ui UI オブジェクト
 * @param {Tab} pagePanel ［ページ］タブ
 * @returns {void}
 */
function buildPageTab(ui, pagePanel) {
    var unitLabel = ui.context.unitLabel;
    var defaults = ui.context.defaults;

    /* 単位：先頭は定規の単位（ドキュメントの設定）/ Units: the first part is the document ruler unit */
    var unitsPanel = addPanel(pagePanel, getLabel("panel.units"));
    var unitRadios = [
        unitsPanel.add("radiobutton", undefined, unitLabel + "/pt/pt/pt"),
        unitsPanel.add("radiobutton", undefined, unitLabel + "/mm/pt/pt"),
        unitsPanel.add("radiobutton", undefined, unitLabel + "/mm/Q/H")
    ];
    for (var i = 0; i < unitRadios.length; i++) unitRadios[i].helpTip = getLabel("tooltip.units");
    unitRadios[0].value = true;
    ui.rbUnitAllPt = unitRadios[0];
    ui.rbUnitStrokeMm = unitRadios[1];
    ui.rbUnitTextQ = unitRadios[2];

    /* 基本テキスト / Base text */
    var baseTextPanel = addPanel(pagePanel, getLabel("panel.baseText"));
    var fontSizeRow = addNumberRow(baseTextPanel, "fieldLabel.fontSize", DEFAULT_FONT_SIZE, 5, "pt");
    ui.baseFontSizeInput = fontSizeRow.input;
    ui.fontSizeUnitLabel = fontSizeRow.unitLabel;
    var leadingRow = addNumberRow(baseTextPanel, "fieldLabel.leading", 16, 5, "pt");
    ui.leadingInput = leadingRow.input;
    ui.leadingUnitLabel = leadingRow.unitLabel;
    var tempGridRow = addRow(baseTextPanel);
    addFieldLabel(tempGridRow, "fieldLabel.tempGrid");
    ui.showTempGridCheck = addCheckbox(tempGridRow, "checkbox.showTempGrid", true, "tooltip.showTempGrid");
    ui.keepTempGridCheck = addCheckbox(tempGridRow, "checkbox.keepTempGrid", true, "tooltip.keepTempGrid");

    /* マージン：額縁と同じ並び＋相対＋ページへの反映 / Margins: same layout as the page frame, plus Relative */
    var marginPanel = addPanel(pagePanel, withUnit("panel.margin", unitLabel));
    ui.marginSides = addLinkedSideInputs(marginPanel, 0, 4);
    ui.marginTopInput = ui.marginSides.top;
    ui.marginBottomInput = ui.marginSides.bottom;
    ui.marginLeftInput = ui.marginSides.left;
    ui.marginRightInput = ui.marginSides.right;
    ui.marginTopInput.text = String(defaults.marginTop);
    ui.marginBottomInput.text = String(defaults.marginBottom);
    ui.marginLeftInput.text = String(defaults.marginLeft);
    ui.marginRightInput.text = String(defaults.marginRight);
    /* 連動は四辺がそろっているときだけ初期オン / Link starts on only when all four sides match */
    ui.marginSides.linkCheck.value = (defaults.marginTop === defaults.marginBottom
        && defaults.marginTop === defaults.marginLeft && defaults.marginTop === defaults.marginRight);
    var relativeRow = addNumberRow(marginPanel, "fieldLabel.relative", 0, 4, unitLabel);
    relativeRow.row.alignment = ["center", "center"];
    relativeRow.input.helpTip = getLabel("tooltip.relative");
    ui.relativeInput = relativeRow.input;
    ui.applyMarginsCheck = addCheckbox(marginPanel, "checkbox.applyMargins", false, "tooltip.applyMargins");

    buildTypeAreaPanel(ui, pagePanel);
}

/**
 * ［額縁］パネルを組み立てる
 * @param {object} ui UI オブジェクト
 * @param {Tab} parent 親タブ
 * @returns {void}
 */
function buildPageFramePanel(ui, parent) {
    var pageFramePanel = addPanel(parent, withUnit("panel.pageFrame", ui.context.unitLabel));
    var pageFrameCheckRow = addRow(pageFramePanel);
    ui.pageFrameCheck = addCheckbox(pageFrameCheckRow, "checkbox.drawPageFrame", false, "tooltip.drawPageFrame");
    ui.pageFrameBleedCheck = addCheckbox(pageFrameCheckRow, "checkbox.bleed", false, "tooltip.bleed");
    ui.pageFrameSides = addLinkedSideInputs(pageFramePanel, 0, 4);
    ui.pageFrameCornerRow = addRow(pageFramePanel);
    ui.pageFrameCornerCheck = addCheckbox(ui.pageFrameCornerRow, "checkbox.frameCorner", false, "tooltip.frameCorner");
    ui.pageFrameCornerInput = addNumberInput(ui.pageFrameCornerRow, 10, 4, 0);
    ui.pageFrameCornerUnitLabel = ui.pageFrameCornerRow.add("statictext", undefined, "pt");
}

/**
 * ［版面の罫線］パネルを組み立てる
 * @param {object} ui UI オブジェクト
 * @param {Tab} parent 親タブ
 * @returns {void}
 */
function buildTypeAreaPanel(ui, parent) {
    var typeAreaPanel = addPanel(parent, getLabel("panel.typeArea"));
    var unitLabel = ui.context.unitLabel;

    var borderRow = addRow(typeAreaPanel);
    ui.borderCheck = addCheckbox(borderRow, "checkbox.drawBorder", false);
    ui.borderWeightInput = addNumberInput(borderRow, 0.3, 5);
    ui.borderWeightInput.helpTip = getLabel("tooltip.borderWeight");
    ui.borderWeightUnitLabel = borderRow.add("statictext", undefined, "pt");

    var cornerRow = addNumberRow(typeAreaPanel, "fieldLabel.cornerRadius", 0, 4, unitLabel);
    cornerRow.input.helpTip = getLabel("tooltip.cornerRadius");
    ui.cornerRadiusRow = cornerRow.row;
    ui.cornerRadiusInput = cornerRow.input;

    var extensionRow = addNumberRow(typeAreaPanel, "fieldLabel.extension", 0, 4, unitLabel);
    extensionRow.input.helpTip = getLabel("tooltip.extension");
    ui.extensionRow = extensionRow.row;
    ui.extensionInput = extensionRow.input;

    ui.capStyleRow = addRow(typeAreaPanel);
    addFieldLabel(ui.capStyleRow, "fieldLabel.capStyle");
    var capRadios = addRadioButtons(ui.capStyleRow, ["radio.capNone", "radio.capRound", "radio.capProjecting"], "tooltip.capStyle");
    ui.rbCapNone = capRadios[0];
    ui.rbCapRound = capRadios[1];
    ui.rbCapProjecting = capRadios[2];
}

/**
 * ［タイトルエリア］パネルを組み立てる
 * @param {object} ui UI オブジェクト
 * @param {Tab} parent 親タブ
 * @returns {void}
 */
function buildTitleAreaPanel(ui, parent) {
    var titleAreaPanel = addPanel(parent, getLabel("panel.titleArea"));
    var unitLabel = ui.context.unitLabel;

    var titleCheckRow = addRow(titleAreaPanel);
    ui.titleFillCheck = addCheckbox(titleCheckRow, "checkbox.fill", false);
    ui.titleStrokeCheck = addCheckbox(titleCheckRow, "checkbox.border", false, "tooltip.titleStroke");

    var lengthRow = addNumberRow(titleAreaPanel, "fieldLabel.height", ui.context.defaults.titleLength, 4, unitLabel);
    lengthRow.input.helpTip = getLabel("tooltip.titleLength");
    ui.titleLengthRow = lengthRow.row;
    ui.titleLengthInput = lengthRow.input;
    ui.btnAutoTitleLength = addAutoButton(lengthRow.row, "tooltip.autoTitleLength");

    ui.titlePositionRow = addRow(titleAreaPanel);
    addFieldLabel(ui.titlePositionRow, "fieldLabel.position");
    var positionRadios = addRadioButtons(ui.titlePositionRow, ["radio.positionTop", "radio.positionBottom", "radio.positionLeft", "radio.positionRight"]);
    ui.rbTitleTop = positionRadios[0];
    ui.rbTitleBottom = positionRadios[1];
    ui.rbTitleLeft = positionRadios[2];
    ui.rbTitleRight = positionRadios[3];

    var extensionRow = addNumberRow(titleAreaPanel, "fieldLabel.extension", 0, 4, unitLabel);
    extensionRow.input.helpTip = getLabel("tooltip.titleExtension");
    ui.titleExtensionRow = extensionRow.row;
    ui.titleExtensionInput = extensionRow.input;
}

/**
 * ［フッターのコラムエリア］パネルを組み立てる
 * @param {object} ui UI オブジェクト
 * @param {Tab} parent 親タブ
 * @returns {void}
 */
function buildFooterAreaPanel(ui, parent) {
    var footerPanel = addPanel(parent, getLabel("panel.footerArea"));
    var unitLabel = ui.context.unitLabel;

    var footerCheckRow = addRow(footerPanel);
    ui.footerFillCheck = addCheckbox(footerCheckRow, "checkbox.fill", false);
    ui.footerStrokeCheck = addCheckbox(footerCheckRow, "checkbox.border", false);
    ui.footerWeightInput = addNumberInput(footerCheckRow, 0.3, 5);
    ui.footerWeightUnitLabel = footerCheckRow.add("statictext", undefined, "pt");

    var heightRow = addNumberRow(footerPanel, "fieldLabel.height", ui.context.defaults.footerHeight, 4, unitLabel);
    ui.footerHeightRow = heightRow.row;
    ui.footerHeightInput = heightRow.input;
    ui.btnAutoFooterHeight = addAutoButton(heightRow.row, "tooltip.autoFooterHeight");

    var gapRow = addNumberRow(footerPanel, "fieldLabel.footerGap", 0, 4, unitLabel, 0);
    gapRow.input.helpTip = getLabel("tooltip.footerGap");
    ui.footerGapRow = gapRow.row;
    ui.footerGapInput = gapRow.input;

    var cornerRow = addNumberRow(footerPanel, "fieldLabel.cornerRadius", 0, 4, unitLabel);
    ui.footerCornerRow = cornerRow.row;
    ui.footerCornerInput = cornerRow.input;
}

/**
 * ［実コンテンツ領域］タブ（オフセット・列行・セル・区切り線）を組み立てる
 * @param {object} ui UI オブジェクト
 * @param {Tab} contentPanel ［実コンテンツ領域］タブ
 * @returns {void}
 */
function buildContentRegionTab(ui, contentPanel) {
    var unitLabel = ui.context.unitLabel;

    /* オフセット / Offset */
    var offsetPanel = addPanel(contentPanel, withUnit("panel.offset", unitLabel));
    ui.offsetSides = addLinkedSideInputs(offsetPanel, 0, 4);
    var offsetFields = [ui.offsetSides.top, ui.offsetSides.bottom, ui.offsetSides.left, ui.offsetSides.right];
    for (var i = 0; i < offsetFields.length; i++) setRoundedDisplayValue(offsetFields[i], ui.context.defaults.offset);
    ui.btnAutoOffset = addAutoButton(offsetPanel, "tooltip.autoOffset");
    ui.btnAutoOffset.alignment = "center";

    /* 列・行 / Columns and rows */
    var rowColPanel = addPanel(contentPanel, getLabel("panel.rowCol"));
    var colCountRow = addNumberRow(rowColPanel, "fieldLabel.columnCount", 2, 5, null, 1);
    ui.colCountInput = colCountRow.input;
    ui.charCountInput = addNumberInput(colCountRow.row, 0, 4, 1);
    ui.charCountInput.helpTip = getLabel("tooltip.characterCount");
    colCountRow.row.add("statictext", undefined, getLabel("unit.characters"));

    var colGapRow = addNumberRow(rowColPanel, "fieldLabel.gap", ui.context.defaults.columnGap, 5, unitLabel, 0);
    ui.colGapInput = colGapRow.input;
    ui.btnAutoColumnGap = addAutoButton(colGapRow.row, "tooltip.autoColumnGap");

    ui.rowCountInput = addNumberRow(rowColPanel, "fieldLabel.rowCount", 1, 5, null, 1).input;

    var rowGapRow = addNumberRow(rowColPanel, "fieldLabel.gap", 0, 5, unitLabel, 0);
    ui.rowGapInput = rowGapRow.input;
    ui.gapLinkCheck = addCheckbox(rowGapRow.row, "checkbox.link", true, "tooltip.linkGaps");

    /* セル / Cells */
    var cellsPanel = addPanel(contentPanel, getLabel("panel.cells"));
    var cellTypeRadios = addRadioButtons(addRow(cellsPanel), ["radio.cellFill", "radio.cellTextFrame"]);
    ui.rbCellFill = cellTypeRadios[0];
    ui.rbCellTextFrame = cellTypeRadios[1];
    ui.threadCheck = addCheckbox(cellsPanel, "checkbox.threadFrames", false, "tooltip.threadFrames");
    ui.sampleTextRow = addRow(cellsPanel);
    var sampleRadios = addRadioButtons(ui.sampleTextRow, ["radio.sampleNone", "radio.sampleProse", "radio.sampleDummy"], "tooltip.sampleText");
    ui.rbSampleNone = sampleRadios[0];
    ui.rbSampleProse = sampleRadios[1];
    ui.rbSampleDummy = sampleRadios[2];

    /* 区切り線 / Dividers */
    var dividerPanel = addPanel(contentPanel, getLabel("panel.divider"));
    var dividerCheckRow = addRow(dividerPanel);
    ui.dividerCheck = addCheckbox(dividerCheckRow, "checkbox.drawDividers", true, "tooltip.dividers");
    ui.dividerWeightInput = addNumberInput(dividerCheckRow, 0.3, 5);
    ui.dividerWeightUnitLabel = dividerCheckRow.add("statictext", undefined, "pt");
    ui.dividerLineTypeRow = addRow(dividerPanel);
    var lineTypeRadios = addRadioButtons(ui.dividerLineTypeRow, ["radio.lineSolid", "radio.lineDashed", "radio.lineDotted"]);
    ui.rbLineSolid = lineTypeRadios[0];
    ui.rbLineDashed = lineTypeRadios[1];
    ui.rbLineDotted = lineTypeRadios[2];
    ui.rbLineDashed.helpTip = ui.rbLineDotted.helpTip = getLabel("tooltip.dividerStyle");
}

/**
 * ボタンエリア（左：一括自動調整、右：キャンセル・OK）を組み立てる
 * @param {object} ui UI オブジェクト
 * @param {Window} dlg ダイアログ
 * @returns {void}
 */
function buildButtonRow(ui, dlg) {
    // メイングループ（横並び） / Main group (horizontal layout)
    var btnRowGroup = dlg.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.alignment = ["fill", "bottom"];

    // 左側グループ / Left-side button group
    var btnLeftGroup = btnRowGroup.add("group");
    btnLeftGroup.alignChildren = ["left", "center"];
    ui.btnAutoAdjustAll = btnLeftGroup.add("button", undefined, getLabel("button.autoAdjustAll"));
    ui.btnAutoAdjustAll.helpTip = getLabel("tooltip.autoAdjustAll");

    // スペーサー（伸縮）/ Spacer (stretchable)
    var spacer = btnRowGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 0;

    // 右側グループ / Right-side button group
    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.alignChildren = ["right", "center"];
    btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
}

// =========================================
// ダイアログの値 / Dialog values
// =========================================

/**
 * 入力欄を数値で読む（0・空欄・読めない値は代わりの値）
 * @param {EditText} editText 入力欄
 * @param {number} fallback 代わりの値
 * @returns {number} 数値
 */
function readNumber(editText, fallback) {
    return parseFloat(editText.text) || fallback;
}

/**
 * 値を丸めずに持ったまま、欄には小数点以下 1 桁までを表示する
 * @param {EditText} editText 入力欄
 * @param {number} value 丸める前の値
 * @returns {void}
 */
function setRoundedDisplayValue(editText, value) {
    editText.text = String(roundTo(value, 1));
    editText.exactValue = value;
    editText.exactValueText = editText.text;  /* 表示が書き換えられたかの目印 / Detects user edits */
}

/**
 * setRoundedDisplayValue() で入れた欄を読む
 * 表示がそのままなら丸める前の値、書き換えられていれば入力値（0・読めない値は 0）
 * @param {EditText} editText 入力欄
 * @returns {number} 数値
 */
function readExactValue(editText) {
    if (editText.exactValueText !== undefined && editText.text === editText.exactValueText) return editText.exactValue;
    return parseFloat(editText.text) || 0;
}

/**
 * 列数・行数を読む（1 未満や読めない値は 1）
 * @param {EditText} editText 入力欄
 * @returns {number} 1 以上の整数
 */
function readCount(editText) {
    var count = parseInt(editText.text, 10);
    return (isNaN(count) || count < 1) ? 1 : count;
}

/**
 * 線幅を pt で読む（mm 入力中なら換算、読めなければ 0.3pt）
 * @param {object} ui UI オブジェクト
 * @param {EditText} editText 入力欄
 * @returns {number} 線幅（pt）
 */
function readStrokeWeight(ui, editText) {
    var weight = parseFloat(editText.text);
    if (isNaN(weight)) return 0.3;
    return ui.strokeUnitIsMm ? weight / MM_PER_POINT : weight;
}

/**
 * タイトルエリアを描くか
 * @param {object} ui UI オブジェクト
 * @returns {boolean} 塗りか罫線がオンなら true
 */
function isTitleOn(ui) {
    return ui.titleFillCheck.value || ui.titleStrokeCheck.value;
}

/**
 * フッターのコラムエリアを描くか
 * @param {object} ui UI オブジェクト
 * @returns {boolean} 塗りか罫線がオンなら true
 */
function isFooterOn(ui) {
    return ui.footerFillCheck.value || ui.footerStrokeCheck.value;
}

/**
 * 選択中のタイトルエリアの位置を取得する
 * @param {object} ui UI オブジェクト
 * @returns {string} "top" / "bottom" / "left" / "right"
 */
function getSelectedTitlePosition(ui) {
    if (ui.rbTitleBottom.value) return "bottom";
    if (ui.rbTitleLeft.value) return "left";
    if (ui.rbTitleRight.value) return "right";
    return "top";
}

/**
 * 選択中の線端を取得する
 * @param {object} ui UI オブジェクト
 * @returns {string} "none" / "round" / "project"
 */
function getSelectedCapStyle(ui) {
    if (ui.rbCapRound.value) return "round";
    if (ui.rbCapProjecting.value) return "project";
    return "none";
}

/**
 * 選択中の区切り線の種類を取得する
 * @param {object} ui UI オブジェクト
 * @returns {string} "solid" / "dashed" / "dotted"
 */
function getSelectedDividerLineType(ui) {
    if (ui.rbLineDashed.value) return "dashed";
    if (ui.rbLineDotted.value) return "dotted";
    return "solid";
}

/**
 * 基本テキストの入力値を pt にする（Q/H 入力中なら換算）
 * @param {object} ui UI オブジェクト
 * @param {number} value 入力値
 * @returns {number} ポイント値
 */
function textInputToPoints(ui, value) {
    return ui.rbUnitTextQ.value ? value / Q_PER_POINT : value;
}

/**
 * 文字サイズと行送りを定規の単位で読む（行送りが読めなければ文字サイズの 1.5 倍）
 * @param {object} ui UI オブジェクト
 * @returns {{fontSize: number, leading: number}} 文字サイズと行送り
 */
function readFontMetrics(ui) {
    var fontSize = readNumber(ui.baseFontSizeInput, DEFAULT_FONT_SIZE);
    var leading = parseFloat(ui.leadingInput.text);
    if (isNaN(leading) || leading <= 0) leading = fontSize * 1.5;
    var pointsPerUnit = ui.context.pointsPerUnit;
    return {
        fontSize: textInputToPoints(ui, fontSize) / pointsPerUnit,
        leading: textInputToPoints(ui, leading) / pointsPerUnit
    };
}

/**
 * 入力中の値から版面の範囲を求める（スプレッド座標・定規の単位）
 * @param {object} ui UI オブジェクト
 * @returns {Array<number>} [上, 左, 下, 右]
 */
function readTypeAreaBounds(ui) {
    return insetBounds(ui.context.pageBounds,
        readNumber(ui.marginTopInput, 0), readNumber(ui.marginLeftInput, 0),
        readNumber(ui.marginBottomInput, 0), readNumber(ui.marginRightInput, 0));
}

/**
 * 入力中の値から実コンテンツ領域を求める
 * @param {object} ui UI オブジェクト
 * @returns {Array<number>} [上, 左, 下, 右]
 */
function readContentBounds(ui) {
    var afterTitle = subtractTitleArea(readTypeAreaBounds(ui), isTitleOn(ui), readNumber(ui.titleLengthInput, 0), getSelectedTitlePosition(ui));
    return subtractFooterArea(afterTitle, isFooterOn(ui), readNumber(ui.footerHeightInput, 0), readNumber(ui.footerGapInput, 0));
}

/**
 * 入力中の値から、オフセットを差し引いたグリッドの範囲を求める
 * @param {object} ui UI オブジェクト
 * @returns {Array<number>} [上, 左, 下, 右]
 */
function readGridBounds(ui) {
    var sides = ui.offsetSides;
    return insetBounds(readContentBounds(ui),
        readExactValue(sides.top), readExactValue(sides.left), readExactValue(sides.bottom), readExactValue(sides.right));
}

/**
 * ダイアログの入力値を描画用の設定値にまとめる
 * @param {object} ui UI オブジェクト
 * @returns {object} 描画に使う設定値（長さは定規の単位、線幅・文字サイズは pt）
 */
function readSettings(ui) {
    var isQ = ui.rbUnitTextQ.value;
    var marginTop = parseNumberOr(ui.marginTopInput.text, 0);
    var cornerRadius = parseNumberOr(ui.cornerRadiusInput.text, 0);
    var fontSizeInput = parseFloat(ui.baseFontSizeInput.text);
    var leadingText = ui.leadingInput.text;
    var leadingInput = parseFloat(leadingText);
    var frameSides = ui.pageFrameSides;
    var offsetSides = ui.offsetSides;

    return {
        marginTop: marginTop,
        marginBottom: parseNumberOr(ui.marginBottomInput.text, marginTop),
        marginLeft: parseNumberOr(ui.marginLeftInput.text, marginTop),
        marginRight: parseNumberOr(ui.marginRightInput.text, marginTop),

        applyPageMargins: ui.applyMarginsCheck.value,

        typeAreaBorder: ui.borderCheck.value,
        borderWeight: readStrokeWeight(ui, ui.borderWeightInput),
        borderCornerRadius: cornerRadius,
        borderExtension: parseNumberOr(ui.extensionInput.text, 0),
        capStyle: getSelectedCapStyle(ui),

        titleFill: ui.titleFillCheck.value,
        titleStroke: ui.titleStrokeCheck.value,
        titleLength: parseNumberOr(ui.titleLengthInput.text, 0),
        titlePosition: getSelectedTitlePosition(ui),
        titleExtension: parseNumberOr(ui.titleExtensionInput.text, 0),
        titleCornerRadius: cornerRadius,

        footerFill: ui.footerFillCheck.value,
        footerStroke: ui.footerStrokeCheck.value,
        footerWeight: readStrokeWeight(ui, ui.footerWeightInput),
        footerHeight: parseNumberOr(ui.footerHeightInput.text, 0),
        footerGap: parseNumberOr(ui.footerGapInput.text, 0),
        footerCornerRadius: parseNumberOr(ui.footerCornerInput.text, 0),

        pageFrame: ui.pageFrameCheck.value,
        pageFrameBleed: ui.pageFrameBleedCheck.value,
        pageFrameTop: parseNumberOr(frameSides.top.text, 0),
        pageFrameBottom: parseNumberOr(frameSides.bottom.text, 0),
        pageFrameLeft: parseNumberOr(frameSides.left.text, 0),
        pageFrameRight: parseNumberOr(frameSides.right.text, 0),
        /* 額縁の内側の角丸（pt）。オフなら 0 / Opening corner radius in pt, 0 when off */
        pageFrameCornerRadiusPt: ui.pageFrameCornerCheck.value ? parseNumberOr(ui.pageFrameCornerInput.text, 0) : 0,

        offsetTop: readExactValue(offsetSides.top),
        offsetBottom: readExactValue(offsetSides.bottom),
        offsetLeft: readExactValue(offsetSides.left),
        offsetRight: readExactValue(offsetSides.right),
        colCount: readCount(ui.colCountInput),
        colGap: parseNumberOr(ui.colGapInput.text, 0),
        rowCount: readCount(ui.rowCountInput),
        rowGap: parseNumberOr(ui.rowGapInput.text, 0),

        cellFill: ui.rbCellFill.value,
        cellTextFrame: ui.rbCellTextFrame.value,
        threadFrames: ui.threadCheck.value,
        sampleProse: ui.rbSampleProse.value,
        sampleDummy: ui.rbSampleDummy.value,

        dividers: ui.dividerCheck.value,
        dividerLineType: getSelectedDividerLineType(ui),
        dividerWeight: readStrokeWeight(ui, ui.dividerWeightInput),

        fontSizePt: isQ ? fontSizeInput / Q_PER_POINT : fontSizeInput,
        /* 行送りは文字列のまま（"auto" を通す）/ Keep leading as text so "auto" passes through */
        leading: (!isNaN(leadingInput) && leadingInput > 0 && isQ) ? String(leadingInput / Q_PER_POINT) : leadingText,
        showTempGrid: ui.showTempGridCheck.value,
        pointsPerUnit: ui.context.pointsPerUnit,

        targetLayer: null,  /* 描画先レイヤー。null ならアクティブレイヤー / Target layer, null for the active layer */
        gridOnly: false     /* 仮グリッドだけを描く / Draw the temp grid only */
    };
}

// =========================================
// ダイアログの更新 / Dialog updates
// =========================================

/**
 * コントロールの有効／無効をまとめて切り替える
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function updateEnabledStates(ui) {
    /* 版面：線幅と角丸はタイトルエリアでも使う / Weight and radius are shared with the title area */
    var borderOn = ui.borderCheck.value;
    var extensionIsZero = readNumber(ui.extensionInput, 0) === 0;
    ui.borderWeightInput.enabled = ui.borderWeightUnitLabel.enabled = borderOn || ui.titleStrokeCheck.value;
    ui.cornerRadiusRow.enabled = (borderOn && extensionIsZero) || ui.titleFillCheck.value;
    ui.extensionRow.enabled = borderOn;
    ui.capStyleRow.enabled = borderOn && !extensionIsZero;

    /* タイトルエリア / Title area */
    var titleOn = isTitleOn(ui);
    ui.titleLengthRow.enabled = titleOn;
    ui.titlePositionRow.enabled = titleOn;
    ui.titleExtensionRow.enabled = ui.titleStrokeCheck.value;

    /* フッターのコラムエリア / Footer column area */
    var footerOn = isFooterOn(ui);
    ui.footerWeightInput.enabled = ui.footerWeightUnitLabel.enabled = ui.footerStrokeCheck.value;
    ui.footerHeightRow.enabled = footerOn;
    ui.footerGapRow.enabled = footerOn;
    ui.footerCornerRow.enabled = footerOn;

    /* 額縁 / Page frame */
    ui.pageFrameBleedCheck.enabled = ui.pageFrameCheck.value;
    ui.pageFrameSides.row.enabled = ui.pageFrameCheck.value;
    ui.pageFrameCornerRow.enabled = ui.pageFrameCheck.value;
    ui.pageFrameCornerInput.enabled = ui.pageFrameCornerUnitLabel.enabled = ui.pageFrameCornerCheck.value;

    /* 間隔：列数・行数が 1 のときは無効 / Gaps are unavailable for a single column or row */
    var hasColumns = readCount(ui.colCountInput) > 1;
    var hasRows = readCount(ui.rowCountInput) > 1;
    ui.colGapInput.enabled = hasColumns;
    ui.rowGapInput.enabled = hasRows;
    ui.gapLinkCheck.enabled = hasColumns && hasRows;

    /* 区切り線：両方の間隔が 0 のときは無効 / Dividers need a gap */
    var hasGap = readNumber(ui.colGapInput, 0) > 0 || readNumber(ui.rowGapInput, 0) > 0;
    var dividersOn = hasGap && ui.dividerCheck.value;
    ui.dividerCheck.enabled = hasGap;
    ui.dividerWeightInput.enabled = ui.dividerWeightUnitLabel.enabled = dividersOn;
    ui.dividerLineTypeRow.enabled = dividersOn;

    /* セル / Cells */
    ui.threadCheck.enabled = ui.rbCellTextFrame.value;
    ui.sampleTextRow.enabled = ui.rbCellTextFrame.value;
}

/**
 * 文字サイズと列幅から、1 列に入る文字数を計算して表示する
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function updateCharCount(ui) {
    var grid = readGridBounds(ui);
    var gridWidth = grid[3] - grid[1];
    if (gridWidth <= 0 || grid[2] - grid[0] <= 0) {
        ui.charCountInput.text = "0";
        return;
    }
    var colCount = parseInt(ui.colCountInput.text, 10) || 1;
    var cellWidth = (gridWidth - readNumber(ui.colGapInput, 0) * (colCount - 1)) / colCount;
    ui.charCountInput.text = String(Math.floor(cellWidth / readFontMetrics(ui).fontSize));
}

/**
 * プレビューを描き直す（プレビューは常にオン）
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function updatePreview(ui) {
    var doc = ui.context.doc;
    clearPreviewLayer(doc);
    var settings = readSettings(ui);
    settings.targetLayer = getOrCreateWorkLayer(doc, PREVIEW_LAYER_NAME);
    drawLayout(settings);
}

/**
 * 値が変わったあとの共通処理（有効／無効・文字数・プレビュー）
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function handleChange(ui) {
    updateEnabledStates(ui);
    updateCharCount(ui);
    updatePreview(ui);
}

/**
 * 基本テキストの入力単位を pt と Q/H で切り替え、入力値を換算する
 * @param {object} ui UI オブジェクト
 * @param {boolean} toQ Q/H に切り替えるなら true
 * @returns {void}
 */
function switchFontUnit(ui, toQ) {
    if (toQ === ui.fontUnitIsQ) return;
    var factor = toQ ? Q_PER_POINT : (1 / Q_PER_POINT);
    var fields = [ui.baseFontSizeInput, ui.leadingInput];
    for (var i = 0; i < fields.length; i++) {
        var value = parseFloat(fields[i].text);
        if (value > 0) fields[i].text = String(roundTo(value * factor, 2));
    }
    ui.fontSizeUnitLabel.text = toQ ? "Q" : "pt";
    ui.leadingUnitLabel.text = toQ ? "H" : "pt";
    ui.fontUnitIsQ = toQ;
}

/**
 * 線幅の入力単位を pt と mm で切り替え、入力値を換算する
 * @param {object} ui UI オブジェクト
 * @param {boolean} toMm mm に切り替えるなら true
 * @returns {void}
 */
function switchStrokeUnit(ui, toMm) {
    if (toMm === ui.strokeUnitIsMm) return;
    var factor = toMm ? MM_PER_POINT : (1 / MM_PER_POINT);
    var unitText = toMm ? "mm" : "pt";
    var fields = [ui.borderWeightInput, ui.footerWeightInput, ui.dividerWeightInput];
    for (var i = 0; i < fields.length; i++) {
        var value = parseFloat(fields[i].text);
        if (value > 0) fields[i].text = String(roundTo(value * factor, 3));
    }
    ui.borderWeightUnitLabel.text = unitText;
    ui.footerWeightUnitLabel.text = unitText;
    ui.dividerWeightUnitLabel.text = unitText;
    ui.strokeUnitIsMm = toMm;
}

/**
 * 1 列の文字数から列の間隔を求めて入力する（連動中なら行の間隔も）
 * @param {object} ui UI オブジェクト
 * @param {number} charCount 1 列の文字数
 * @returns {void}
 */
function applyGapForCharCount(ui, charCount) {
    var grid = readGridBounds(ui);
    var colCount = parseInt(ui.colCountInput.text, 10) || 1;
    var newGap = 0;
    if (colCount > 1) {
        var cellWidth = charCount * readFontMetrics(ui).fontSize;
        newGap = roundTo(((grid[3] - grid[1]) - cellWidth * colCount) / (colCount - 1), 3);
        if (newGap < 0) newGap = 0;
    }
    ui.colGapInput.text = String(newGap);
    if (ui.gapLinkCheck.value) ui.rowGapInput.text = ui.colGapInput.text;
}

// =========================================
// 自動調整 / Auto adjust
// =========================================

/**
 * タイトルエリアの長さを仮グリッドの線に合わせる
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function autoAdjustTitleLength(ui) {
    var titleLength = readNumber(ui.titleLengthInput, 0);
    if (titleLength <= 0) return;
    ui.titleLengthInput.text = String(roundTo(snapToTempGrid(titleLength, readFontMetrics(ui)), 3));
}

/**
 * コラムエリアの上端が仮グリッドの線に乗るように高さを調整する
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function autoAdjustFooterHeight(ui) {
    var footerHeight = readNumber(ui.footerHeightInput, 0);
    if (footerHeight <= 0) return;
    var fontMetrics = readFontMetrics(ui);
    var typeArea = readTypeAreaBounds(ui);
    var footerGap = readNumber(ui.footerGapInput, 0);

    var footerTop = typeArea[2] - footerGap - footerHeight;
    var snappedFooterTop = typeArea[0] + snapToTempGrid(footerTop - typeArea[0], fontMetrics);
    var newHeight = typeArea[2] - footerGap - snappedFooterTop;
    if (newHeight <= 0) newHeight = fontMetrics.fontSize;
    ui.footerHeightInput.text = String(roundTo(newHeight, 3));
}

/**
 * オフセットを調整する。左右は文字サイズの倍数、天地は仮グリッドの線に合わせる
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function autoAdjustOffsets(ui) {
    var sides = ui.offsetSides;
    var fontMetrics = readFontMetrics(ui);

    /* 左右 / Left and right */
    var horizontalInputs = [sides.left, sides.right];
    for (var i = 0; i < horizontalInputs.length; i++) {
        var value = readExactValue(horizontalInputs[i]);
        if (value === 0) continue;
        var charCount = Math.round(value / fontMetrics.fontSize);
        if (charCount < 1) charCount = 1;
        setRoundedDisplayValue(horizontalInputs[i], charCount * fontMetrics.fontSize);
    }

    /* 天地：実コンテンツ領域の端＋オフセットが仮グリッドの線に乗るように / Snap top and bottom to temp grid lines */
    var typeAreaTop = readTypeAreaBounds(ui)[0];
    var content = readContentBounds(ui);

    var offsetTop = readExactValue(sides.top);
    if (offsetTop > 0) {
        var snappedTop = typeAreaTop + snapToTempGrid(content[0] + offsetTop - typeAreaTop, fontMetrics);
        setRoundedDisplayValue(sides.top, Math.max(snappedTop - content[0], 0));
    }

    var offsetBottom = readExactValue(sides.bottom);
    if (offsetBottom > 0) {
        var snappedBottom = typeAreaTop + snapToTempGrid(content[2] - offsetBottom - typeAreaTop, fontMetrics);
        setRoundedDisplayValue(sides.bottom, Math.max(content[2] - snappedBottom, 0));
    }
}

/**
 * ［文字］の値から列の間隔を計算し直す
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function autoAdjustColumnGap(ui) {
    applyGapForCharCount(ui, readCount(ui.charCountInput));
}

/**
 * 有効な［自動調整］をまとめて実行する
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function autoAdjustAll(ui) {
    if (isTitleOn(ui)) autoAdjustTitleLength(ui);
    if (isFooterOn(ui)) autoAdjustFooterHeight(ui);
    autoAdjustOffsets(ui);
    updateCharCount(ui);  /* 列の間隔は調整後の文字数から / Gap uses the updated character count */
    autoAdjustColumnGap(ui);
}

// =========================================
// イベント / Events
// =========================================

/**
 * 入力中と確定時の両方に同じ処理を結び付ける
 * @param {EditText} editText 入力欄
 * @param {Function} handler 処理
 * @returns {void}
 */
function bindEditHandler(editText, handler) {
    editText.onChanging = handler;
    editText.onChange = handler;
}

/**
 * 天地左右の入力欄を連動させる
 * @param {object} sides addLinkedSideInputs() の戻り値
 * @param {Function} afterChange 連動後の処理
 * @returns {void}
 */
function bindLinkedSides(sides, afterChange) {
    var fields = [sides.top, sides.bottom, sides.left, sides.right];
    for (var i = 0; i < fields.length; i++) {
        (function (source) {
            bindEditHandler(source, function () {
                if (sides.linkCheck.value) {
                    for (var j = 0; j < fields.length; j++) {
                        /* 丸める前の値も一緒に写す / Copy the unrounded value too */
                        fields[j].text = source.text;
                        fields[j].exactValue = source.exactValue;
                        fields[j].exactValueText = source.exactValueText;
                    }
                }
                afterChange();
            });
        })(fields[i]);
    }
}

/**
 * ダイアログのコントロールにイベントを結び付ける
 * @param {object} ui UI オブジェクト
 * @returns {void}
 */
function bindDialogEvents(ui) {
    var refresh = function () { handleChange(ui); };
    var i;

    var clickControls = [
        ui.showTempGridCheck,
        ui.borderCheck, ui.rbCapNone, ui.rbCapRound, ui.rbCapProjecting,
        ui.titleFillCheck, ui.titleStrokeCheck, ui.rbTitleTop, ui.rbTitleBottom, ui.rbTitleLeft, ui.rbTitleRight,
        ui.footerFillCheck, ui.footerStrokeCheck,
        ui.pageFrameCheck, ui.pageFrameBleedCheck, ui.pageFrameCornerCheck,
        ui.threadCheck, ui.dividerCheck, ui.rbLineSolid, ui.rbLineDashed, ui.rbLineDotted
    ];
    for (i = 0; i < clickControls.length; i++) clickControls[i].onClick = refresh;

    var editControls = [
        ui.baseFontSizeInput, ui.leadingInput,
        ui.borderWeightInput, ui.cornerRadiusInput, ui.extensionInput,
        ui.titleLengthInput, ui.titleExtensionInput,
        ui.footerWeightInput, ui.footerHeightInput, ui.footerGapInput, ui.footerCornerInput,
        ui.colCountInput, ui.rowCountInput, ui.dividerWeightInput, ui.pageFrameCornerInput
    ];
    for (i = 0; i < editControls.length; i++) bindEditHandler(editControls[i], refresh);

    bindLinkedSides(ui.marginSides, refresh);
    bindLinkedSides(ui.pageFrameSides, refresh);
    bindLinkedSides(ui.offsetSides, refresh);

    /* 入力単位 / Input units */
    ui.rbUnitAllPt.onClick = function () { switchFontUnit(ui, false); switchStrokeUnit(ui, false); refresh(); };
    ui.rbUnitStrokeMm.onClick = function () { switchFontUnit(ui, false); switchStrokeUnit(ui, true); refresh(); };
    ui.rbUnitTextQ.onClick = function () { switchFontUnit(ui, true); switchStrokeUnit(ui, true); refresh(); };

    /* テキストフレームを選んだら連結とサンプル文をオンに / Text frames default to threaded with sample text */
    ui.rbCellFill.onClick = ui.rbCellTextFrame.onClick = function () {
        if (ui.rbCellTextFrame.value) {
            ui.threadCheck.value = true;
            ui.rbSampleProse.value = true;
            ui.rbSampleNone.value = false;
            ui.rbSampleDummy.value = false;
        }
        refresh();
    };

    /* 列と行の間隔の連動 / Linked column and row gaps */
    bindEditHandler(ui.colGapInput, function () {
        if (ui.gapLinkCheck.value) ui.rowGapInput.text = ui.colGapInput.text;
        refresh();
    });
    bindEditHandler(ui.rowGapInput, function () {
        if (ui.gapLinkCheck.value) ui.colGapInput.text = ui.rowGapInput.text;
        refresh();
    });

    /* 文字数から間隔を逆算（入力中の文字数は書き戻さない）/ Derive the gap without rewriting the count being typed */
    bindEditHandler(ui.charCountInput, function () {
        var charCount = parseInt(ui.charCountInput.text, 10);
        if (isNaN(charCount) || charCount < 1) return;
        applyGapForCharCount(ui, charCount);
        updateEnabledStates(ui);
        updatePreview(ui);
    });

    /* 相対：前回からの増減分を四辺のマージンに加える / Relative: add the change to all margins */
    bindEditHandler(ui.relativeInput, function () {
        var relativeValue = readNumber(ui.relativeInput, 0);
        var delta = relativeValue - ui.lastRelativeValue;
        ui.lastRelativeValue = relativeValue;
        var marginInputs = [ui.marginTopInput, ui.marginBottomInput, ui.marginLeftInput, ui.marginRightInput];
        for (var j = 0; j < marginInputs.length; j++) {
            marginInputs[j].text = String(roundTo(readNumber(marginInputs[j], 0) + delta, 2));
        }
        refresh();
    });

    /* 自動調整 / Auto adjust */
    ui.btnAutoTitleLength.onClick = function () { autoAdjustTitleLength(ui); refresh(); };
    ui.btnAutoFooterHeight.onClick = function () { autoAdjustFooterHeight(ui); refresh(); };
    ui.btnAutoOffset.onClick = function () { autoAdjustOffsets(ui); refresh(); };
    ui.btnAutoColumnGap.onClick = function () { autoAdjustColumnGap(ui); refresh(); };
    ui.btnAutoAdjustAll.onClick = function () { autoAdjustAll(ui); refresh(); };
}

// =========================================
// メイン / Main
// =========================================

/**
 * スプレッド座標系でのページ境界を定規の単位で取得する
 * @param {Page} targetPage 対象ページ
 * @param {number} pointsPerUnit 1 単位あたりのポイント数
 * @returns {Array<number>} [上, 左, 下, 右]
 */
function getPageBoundsOnSpread(targetPage, pointsPerUnit) {
    var topLeft = targetPage.resolve(AnchorPoint.TOP_LEFT_ANCHOR, CoordinateSpaces.SPREAD_COORDINATES)[0];
    var bottomRight = targetPage.resolve(AnchorPoint.BOTTOM_RIGHT_ANCHOR, CoordinateSpaces.SPREAD_COORDINATES)[0];
    return [topLeft[1] / pointsPerUnit, topLeft[0] / pointsPerUnit, bottomRight[1] / pointsPerUnit, bottomRight[0] / pointsPerUnit];
}

/**
 * 入力欄の初期値を求める（すべて横の定規の単位）
 * @param {Page} targetPage 対象ページ
 * @param {Array<number>} pageBounds ページ [上, 左, 下, 右]
 * @param {{horizontal: number, vertical: number}} unitScale 1 単位あたりのポイント数
 * @returns {object} マージン・タイトルエリアの長さ・コラムエリアの高さ・オフセット・列の間隔
 */
function getDefaultInputs(targetPage, pageBounds, unitScale) {
    var marginPrefs = targetPage.marginPreferences;
    /* 天地のマージンは縦の単位なので横の単位に換算 / Top and bottom margins are in vertical units */
    var verticalToHorizontal = unitScale.vertical / unitScale.horizontal;
    var marginTop = roundTo(marginPrefs.top * verticalToHorizontal, 3);
    var marginBottom = roundTo(marginPrefs.bottom * verticalToHorizontal, 3);
    /* 左ページは内側（left）が右に来るので入れ替える / Swap inside and outside on left-hand pages */
    var isLeftPage = (targetPage.side === PageSideOptions.LEFT_HAND);
    var pointsPerUnit = unitScale.horizontal;
    return {
        marginTop: marginTop,
        marginBottom: marginBottom,
        marginLeft: roundTo(isLeftPage ? marginPrefs.right : marginPrefs.left, 3),
        marginRight: roundTo(isLeftPage ? marginPrefs.left : marginPrefs.right, 3),
        /* タイトルエリアの長さ：版面の高さの 1/5 / Title length: one fifth of the type area height */
        titleLength: roundTo(((pageBounds[2] - pageBounds[0]) - marginTop - marginBottom) / 5, 2),
        footerHeight: roundTo(millimetersToUnits(DEFAULT_LENGTHS_MM.footerHeight, pointsPerUnit), 2),
        offset: millimetersToUnits(DEFAULT_LENGTHS_MM.offset, pointsPerUnit),  /* 丸めずに持つ / Kept unrounded */
        columnGap: roundTo(millimetersToUnits(DEFAULT_LENGTHS_MM.columnGap, pointsPerUnit), 2)
    };
}

/**
 * ページの「マージン・段組」のマージンを書き換える
 * 左ページは left が内側（画面の右）なので入れ替え、単位の食い違いを避けて pt で入れる
 * @param {Page} targetPage 対象ページ
 * @param {object} settings 設定値（マージンは横の定規の単位）
 * @returns {void}
 */
function applyPageMargins(targetPage, settings) {
    var marginPrefs = targetPage.marginPreferences;
    var isLeftPage = (targetPage.side === PageSideOptions.LEFT_HAND);
    var toPoints = function (value) { return toPointString(value * settings.pointsPerUnit); };
    marginPrefs.top = toPoints(settings.marginTop);
    marginPrefs.bottom = toPoints(settings.marginBottom);
    marginPrefs.left = toPoints(isLeftPage ? settings.marginRight : settings.marginLeft);
    marginPrefs.right = toPoints(isLeftPage ? settings.marginLeft : settings.marginRight);
}

/**
 * 確定した設定で描画する（必要ならページのマージンも変更し、仮グリッドを残すなら専用レイヤーにも描く）
 * @param {Document} doc 対象ドキュメント
 * @param {object} settings 設定値
 * @param {boolean} keepTempGrid 仮グリッドを残すか
 * @returns {void}
 */
function drawFinalLayout(doc, settings, keepTempGrid) {
    if (settings.applyPageMargins) applyPageMargins(app.activeWindow.activePage, settings);
    activateOtherLayer(doc, [TEMP_GRID_LAYER_NAME, PREVIEW_LAYER_NAME]);
    drawLayout(settings);
    if (!keepTempGrid) return;

    settings.targetLayer = getOrCreateWorkLayer(doc, TEMP_GRID_LAYER_NAME);
    settings.gridOnly = true;
    drawLayout(settings);
}

/**
 * ドキュメントを確認し、設定ダイアログを表示してレイアウトを作成する
 * @returns {void}
 */
function main() {
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var doc = app.activeDocument;
    var page = app.activeWindow.activePage;
    var unitScale = measurePointsPerUnit(page);
    var pageBounds = getPageBoundsOnSpread(page, unitScale.horizontal);
    var ui = buildDialog({
        doc: doc,
        page: page,
        unitLabel: getRulerUnitLabel(doc.viewPreferences.horizontalMeasurementUnits),
        pointsPerUnit: unitScale.horizontal,
        pageBounds: pageBounds,
        defaults: getDefaultInputs(page, pageBounds, unitScale)
    });

    bindDialogEvents(ui);
    handleChange(ui);

    var confirmed = (ui.dlg.show() === 1);
    removePreviewLayer(doc);
    if (!confirmed) return;

    var settings = readSettings(ui);
    var keepTempGrid = ui.keepTempGridCheck.value && ui.showTempGridCheck.value;
    app.doScript(function () {
        drawFinalLayout(doc, settings, keepTempGrid);
    }, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, getLabel("dialog.title"));
}

main();
