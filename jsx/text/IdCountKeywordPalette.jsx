#target indesign
#targetengine "IdCountKeywordPalette"

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdCountKeywordPalette
Name=キーワードの件数をパレットで数える
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.0.0
Release-Date=2026-11-05
Description-BEGIN
最大5つのキーワードが、選択中のストーリーまたはドキュメント全体にそれぞれいくつあるかを、常駐パレットで数えます。

A floating palette that counts how many times each of up to five keywords appears in the selected stories or in the whole document.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

最大5つのキーワードが、選択中のストーリーまたはドキュメント全体にそれぞれいくつあるかを、常駐パレットで数えます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCountKeywordPalette.md

### Overview

A floating palette that counts how many times each of up to five keywords appears in the selected stories or in the whole document.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCountKeywordPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdCountKeywordPalette";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-05";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "キーワードの件数をパレットで数える"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCountKeywordPalette.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCountKeywordPalette.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdCountKeywordPalette.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* キーワード欄の数 / Number of keyword fields */
    var KEYWORD_FIELD_COUNT = 5;

    /* 一覧に出す前後の文字数の初期値と上限 / Default and maximum number of characters shown around each match */
    var DEFAULT_CONTEXT_LENGTH = 2;
    var MAX_CONTEXT_LENGTH = 50;

    /* 一覧に出す件数の上限（多すぎると一覧の作成に時間がかかる）/ Maximum rows in the match list (large lists are slow to build) */
    var MAX_LIST_ITEMS = 500;

    /* 前回の設定（InDesign を終了するまで。#targetengine が必要）/ Last settings, kept until InDesign quits (needs #targetengine) */
    var SESSION_KEY = SCRIPT_NAME + "_lastSettings";

    /* パレットの参照を残す $.global のキー（再実行で前回のパレットを閉じるため）
       / $.global key holding the palette, so a rerun can close the previous one */
    var PALETTE_GLOBAL_KEY = "__IdCountKeywordPalette";

    /* 冒頭の #targetengine と同じ名前 / Same name as the #targetengine directive at the top */
    var PALETTE_ENGINE_NAME = "IdCountKeywordPalette";

    /* ランチャーの $.evalFile から呼ばれると #targetengine が効かず、呼び出し元のエンジンが終わるとパレットが消える。
       このファイルを doScript で自分のエンジンに開き直す
       / Run through a launcher's $.evalFile, #targetengine is ignored and the palette vanishes with the caller's engine,
         so rerun this file with doScript in its own engine */
    if ($.engineName !== PALETTE_ENGINE_NAME) {
        app.doScript(File($.fileName), ScriptLanguage.JAVASCRIPT);
        return;
    }

    // =========================================
    // ラベル定義 / Labels
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    var LABELS = {
        palette: {
            title: { ja: "キーワードを数える", en: "Count Keywords" }
        },
        fieldLabel: {
            total: { ja: "合計", en: "Total" },
            contextLength: { ja: "前後の文字数", en: "Characters Around" }
        },
        panel: {
            keywords: { ja: "キーワード", en: "Keywords" },
            scope: { ja: "対象", en: "Scope" },
            options: { ja: "オプション", en: "Options" },
            matches: { ja: "見つかった場所", en: "Matches" }
        },
        column: {
            page: { ja: "ページ", en: "Page" },
            context: { ja: "前後の文字列", en: "Context" }
        },
        radio: {
            story: { ja: "ストーリー", en: "Story" },
            document: { ja: "ドキュメント", en: "Document" }
        },
        checkbox: {
            caseSensitive: { ja: "大文字と小文字を区別", en: "Case Sensitive" },
            wholeWord: { ja: "単語単位", en: "Whole Word" },
            masterPages: { ja: "親ページを含む", en: "Include Parent Pages" },
            excludeContained: { ja: "ほかのキーワードに含まれる分は数えない", en: "Skip matches inside other keywords" }
        },
        tooltip: {
            story: { ja: "［数える］を押した時点で選択している文字やテキストフレームを含むストーリー（表・脚注を含む）で数えます。", en: "Counts in the stories that contain the text or text frames selected when you click Count, including tables and footnotes." },
            document: { ja: "作業中のドキュメント内のすべてのテキストで数えます。", en: "Counts in all text in the active document." },
            masterPages: { ja: "ドキュメントを対象にしたとき、親ページのテキストも数えます。", en: "Also counts text on parent pages when the scope is the document." },
            excludeContained: {
                ja: "「取引」と「取引き」のように、長いキーワードの一部として見つかった分を短いキーワードの件数から除きます。",
                en: "When a keyword is found as part of a longer keyword (e.g. \"deal\" inside \"dealer\"), it is not counted for the shorter one."
            },
            useSelection: { ja: "選択している文字を、最後に入力したキーワード欄（空いていれば最初の空欄）に入れます。", en: "Puts the selected text into the keyword field you last typed in, or the first empty one." },
            keyword: { ja: "return キーでも数えます。", en: "Press Return to count." },
            contextLength: { ja: "見つかった文字の前後に表示する文字数", en: "Number of characters shown before and after each match" },
            matchList: { ja: "クリックすると、その文字を選択して表示します。", en: "Click to select and show the text." },
            /* ステップボタン用 / for the stepper */
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        button: {
            useSelection: { ja: "選択文字を入れる", en: "Use Selection" },
            close: { ja: "閉じる", en: "Close" },
            count: { ja: "数える", en: "Count" }
        },
        status: {
            initial: { ja: "-", en: "-" },
            count: { ja: "{count} 件", en: "{count}" },
            total: { ja: "{count} 件（{scope}）", en: "{count} ({scope})" },
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noStory: { ja: "テキストを選択してください。", en: "Please select text." },
            failed: { ja: "数えられませんでした: {message}", en: "Could not count: {message}" },
            truncated: { ja: "{total} 件のうち先頭の {shown} 件を表示", en: "Showing the first {shown} of {total}" },
            overset: { ja: "オーバーセット", en: "Overset" },
            pasteboard: { ja: "ペーストボード", en: "Pasteboard" },
            unavailable: { ja: "（参照できません）", en: "(unavailable)" }
        }
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    /* コントロールの寸法 / Control sizes */
    var KEYWORD_FIELD_WIDTH = 200;          /* キーワード欄の幅（px）/ width of a keyword field */
    var COUNT_TEXT_WIDTH = 70;              /* 欄の右に出す件数の幅（px）/ width of the count next to each field */
    var MATCH_LIST_WIDTH = 380;             /* 見つかった場所の一覧の幅（px）/ width of the match list */
    var MATCH_LIST_MIN_HEIGHT = 200;        /* 一覧の最小の高さ（px。左のカラムに合わせて伸びる）/ minimum list height; it grows to the left column */
    var MATCH_LIST_PAGE_COLUMN_WIDTH = 90;  /* 一覧のページ列の幅（px）/ width of the page column */

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)

    // -----------------------------------------
    // ステップボタンの寸法・増減量 / Stepper metrics and steps
    // -----------------------------------------
    var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
    var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
    var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
    var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
    var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
    var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
    var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

    /* 単位の換算は UnitValue に任せる（in / ft / yd / mm / cm / m / pt / pc / px ほか、単数形・複数形も可）。
       UnitValue に無い単位だけ、ここで UnitValue の単位に読み替える（値は「1単位＝何 unit か」）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Units UnitValue lacks, mapped onto UnitValue units (how many of `unit` make one) */
    var STEPPER_UNIT_ALIASES = {
        "q": { unit: "mm", amount: 0.25 },    /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 },    /* 歯 / H */
        "p": { unit: "pc", amount: 1 },       /* パイカ / pica */
        "ft/in": { unit: "ft", amount: 1 },   /* Illustrator の単位コード7の表示 / Illustrator unit code 7 */
        "c": { unit: "ci", amount: 1 },       /* シセロ（InDesign の表示） / ciceros as InDesign shows them */
        "ag": { unit: "in", amount: 1 / 14 }, /* アゲート / agates */
        "ap": { unit: "tpt", amount: 1 }      /* アメリカンポイント / American points */
    };

    // -----------------------------------------
    // ステップボタンの配色 / Stepper colors
    // -----------------------------------------
    var STEPPER_UI_DARK           = isDarkUI();
    /* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
       ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
       UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
       dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
    var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
    var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
    var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
    var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
    var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
    var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

    // -----------------------------------------
    // 数値欄を作る（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
     * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す。
     * 四則演算（+ - * / と括弧）を入れると、確定時に計算した値にする。欄と異なる単位で入れた値は欄の単位へ換算する（mm の欄に「1 in」→「25.4 mm」）
     * @param {Group|Panel} parent - 追加先
     * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
     *     step / min / max / integer（true で整数のみ）/ unit / onStep
     * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
     */
    function addSteppedField(parent, fieldOptions) {
        var fieldRowGroup = parent.add("group");
        fieldRowGroup.orientation = "row";
        fieldRowGroup.alignChildren = ["left", "center"];
        fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

        var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
        if (fieldOptions.labelWidth) {
            fieldLabel.preferredSize.width = fieldOptions.labelWidth;
            fieldLabel.justify = "right";
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
        numberInput.characters = fieldOptions.characters || 6;
        numberInput.fieldLabel = fieldLabel;
        numberInput.stepperGroup = stepperGroup;

        /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(numberInput, stepperGroup);

        /* 直接入力をそろえる。計算式は計算し、数値でなければ直前の値に戻す / normalize typed values; evaluate arithmetic, revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = evaluateArithmetic(numberInput.text, fieldOptions.unit);
            if (isNaN(value)) {
                numberInput.text = numberInput.lastValidText;
                return;
            }
            writeSteppedValue(numberInput, value, fieldOptions);
        };
        return numberInput;
    }

    /**
     * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedFieldEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.fieldLabel.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
        for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
            redrawStepperGroup(numberInput.stepperGroup.children[i]);
        }
    }

    /**
     * 数値欄の単位を差し替える（単位の設定やドロップダウンを切り替えたとき用）。
     * shouldConvert が true なら値を新しい単位へ換算し（10 mm → 28.35 pt）、false なら数値はそのままで単位だけ付け替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄、または bindSteppedArrowKeys() を呼んだ入力欄
     * @param {string} unit - 新しい単位（例 " pt"。単位なしは ""）
     * @param {boolean} [shouldConvert] - 値も換算するなら true
     * @returns {void}
     */
    function setSteppedFieldUnit(numberInput, unit, shouldConvert) {
        var stepOptions = numberInput.stepperGroup.stepOptions;
        var oldUnit = stepOptions.unit || "";
        var value = parseFloat(numberInput.text);
        stepOptions.unit = unit;
        if (isNaN(value)) return;
        if (shouldConvert) {
            var converted = evaluateArithmetic(String(value) + oldUnit, unit);
            if (!isNaN(converted)) value = converted;
        }
        numberInput.text = formatStepperNumber(value) + unit;
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
     * @param {Group|Panel} parent - 追加先
     * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
     * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
     * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
     */
    function addStepper(parent, getNumberInput, stepOptions) {
        var stepperGroup = parent.add("group");
        stepperGroup.orientation = "column";
        stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
        stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
        stepperGroup.alignment = ["left", "center"];

        /**
         * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
         * @param {number} direction - 増やすなら 1、減らすなら -1
         * @returns {void}
         */
        function stepBy(direction) {
            var numberInput = getNumberInput();
            if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
            var value = evaluateArithmetic(numberInput.text, stepOptions.unit); /* 確定前の計算式も計算してから増減 / evaluate an uncommitted expression first */
            if (isNaN(value)) value = parseFloat(numberInput.text); /* 計算できなければ従来どおり先頭の数値 / fall back to the leading number */
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        stepperGroup.stepOptions = stepOptions; /* 確定時の計算で欄の単位を引けるよう公開 / lets the commit-time evaluation find the unit */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し。
     * あわせて、確定時に計算式・単位付きの値を計算して書き戻す（各スクリプトの onChange より先に呼ばれるので、onChange は計算後の値を読む）
     * @param {EditText} numberInput - 対象の入力欄
     * @param {Group} stepperGroup - addStepper() で作った∧∨
     * @returns {void}
     */
    function bindSteppedArrowKeys(numberInput, stepperGroup) {
        numberInput.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
            event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
        });
        numberInput.addEventListener("change", function () {
            var fieldUnit = stepperGroup.stepOptions ? stepperGroup.stepOptions.unit : undefined;
            var value = evaluateArithmetic(numberInput.text, fieldUnit);
            if (isNaN(value)) return; /* 計算できなければ各スクリプトの処理に任せる / leave it to the script's own handler */
            /* 式か、換算で値が変わったときだけ書き戻す（ただの数値は書式を崩さない） / rewrite only expressions and converted values */
            var hasOperator = /[*\/()\u00D7\u00F7\uFF0A\uFF0F\uFF08\uFF09]|[\d.\uFF10-\uFF19][^\d.\uFF10-\uFF19]*[+\-\u2212\uFF0B\uFF0D]/.test(numberInput.text);
            if (!hasOperator && value === parseFloat(numberInput.text)) {
                /* 単位を省いて入れた数値には、欄の単位だけ付け足す（桁は丸めない） / append the field unit to a bare number */
                var trimmedText = numberInput.text.replace(/^\s+|\s+$/g, "");
                if (fieldUnit && /[\d.]$/.test(trimmedText)) numberInput.text = trimmedText + fieldUnit;
                return;
            }
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
        });
        numberInput.stepperGroup = stepperGroup; /* setSteppedFieldUnit() から∧∨の設定を引けるようにする / lets setSteppedFieldUnit() find the options */
    }

    // -----------------------------------------
    // 値の計算 / Value helpers
    // -----------------------------------------
    /**
     * 押された修飾キーに応じて、1回分増減した値を返す
     * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
     * 整数の欄では option を無視して step の倍数へ）
     * @param {number} value - 元の値
     * @param {number} direction - 増やすなら 1、減らすなら -1
     * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
     * @returns {number} 増減した値（下限・上限は未適用）
     */
    function computeSteppedValue(value, direction, stepOptions) {
        var keyState = ScriptUI.environment.keyboardState;
        if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
        if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
        return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
    }

    /**
     * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
     * @param {number} value - 元の値
     * @param {number} multiple - 倍数の単位（例 10）
     * @param {number} direction - 上げるなら 1、下げるなら -1
     * @returns {number} 移した値
     */
    function snapStepperToNextMultiple(value, multiple, direction) {
        /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
           round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
        var quotient = Math.round(value / multiple * 1e6) / 1e6;
        if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
        return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
    }

    /**
     * 値を下限・上限の範囲に収める
     * @param {number} value - 数値
     * @param {Object} rangeOptions - min / max（どちらも省略可）
     * @returns {number} 範囲に収めた値
     */
    function clampSteppedValue(value, rangeOptions) {
        if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
        if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
        return value;
    }

    /**
     * 入力欄の文字列を四則演算（+ - * / と括弧）として計算する。eval は使わない。
     * 数値の後ろの単位は UnitValue で欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
     * 全角の数字・記号と × ÷ は半角に直す
     * @param {string} text - 入力欄の文字列
     * @param {string} [fieldUnit] - 欄の単位（例 " mm"。前後の空白は無視）
     * @returns {number} 欄の単位での計算結果（式として読めない・換算できない単位・0で割ったときは NaN）
     */
    function evaluateArithmetic(text, fieldUnit) {
        var source = String(text)
            .replace(/[！-～]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
            .replace(/×/g, "*")
            .replace(/÷/g, "/")
            .replace(/[−–—]/g, "-")
            .replace(/\s/g, "");
        if (source === "") return NaN;
        var fieldUnitKey = String(fieldUnit || "").replace(/^\s+|\s+$/g, "").toLowerCase();
        var fieldUnitValue = createStepperUnitValue(1, fieldUnitKey); /* 欄の単位の1単位（換算できない欄は null） / one field unit */
        var position = 0;

        /**
         * 加減算の並び（項 ± 項 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readSum() {
            var total = readProduct();
            while (position < source.length && (source.charAt(position) === "+" || source.charAt(position) === "-")) {
                var operator = source.charAt(position++);
                var operand = readProduct();
                total = (operator === "+") ? total + operand : total - operand;
            }
            return total;
        }

        /**
         * 乗除算の並び（因子 × 因子 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readProduct() {
            var total = readFactor();
            while (position < source.length && (source.charAt(position) === "*" || source.charAt(position) === "/")) {
                var operator = source.charAt(position++);
                var operand = readFactor();
                if (operator === "/" && operand === 0) return NaN;
                total = (operator === "*") ? total * operand : total / operand;
            }
            return total;
        }

        /**
         * 符号付きの数値（単位付きなら欄の単位へ換算）か、括弧で囲んだ式を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readFactor() {
            var ch = source.charAt(position);
            if (ch === "+" || ch === "-") {
                position++;
                var signedValue = readFactor();
                return (ch === "-") ? -signedValue : signedValue;
            }
            if (ch === "(") {
                position++;
                var innerValue = readSum();
                if (source.charAt(position) !== ")") return NaN;
                position++;
                return innerValue;
            }
            var numberMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
            if (!numberMatch) return NaN;
            position += numberMatch[0].length;
            return readUnitSuffix(parseFloat(numberMatch[0]));
        }

        /**
         * 数値の直後の単位を読み、欄の単位へ換算する
         * @param {number} value - 単位の前の数値
         * @returns {number} 欄の単位での値（換算できない単位なら NaN）
         */
        function readUnitSuffix(value) {
            var unitMatch = /^(ft\/in|[A-Za-z]+|%|°)/i.exec(source.substring(position));
            if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
            position += unitMatch[0].length;
            var unitKey = unitMatch[0].toLowerCase();
            if (unitKey === fieldUnitKey) return value;
            var typedValue = createStepperUnitValue(value, unitKey);
            if (!typedValue || !fieldUnitValue) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = typedValue.as("pt");
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldUnitValue.as("pt");
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
    }

    /**
     * 数値と単位から UnitValue を作る。Q・H・p は STEPPER_UNIT_ALIASES で UnitValue の単位に読み替える。
     * %（percent）は基準の長さが無いと換算できないので扱わない
     * @param {number} value - 数値
     * @param {string} unitKey - 単位（小文字。例 "mm"、"inches"、"q"）
     * @returns {UnitValue|null} UnitValue（UnitValue が知らない単位・空・% なら null）
     */
    function createStepperUnitValue(value, unitKey) {
        if (unitKey === "" || unitKey === "%") return null;
        var alias = STEPPER_UNIT_ALIASES[unitKey];
        var unitValue = alias ? new UnitValue(value * alias.amount, alias.unit) : new UnitValue(value, unitKey);
        if (unitValue.type === "?" || unitValue.type === "%") return null; /* 知らない単位は例外にならず "?" になる。"percent" も除く / unknown units become "?" */
        return unitValue;
    }

    /**
     * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
     * @param {EditText} numberInput - 書き込む入力欄
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {void}
     */
    function writeSteppedValue(numberInput, value, valueOptions) {
        numberInput.text = formatSteppedValue(value, valueOptions);
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
     * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
     */
    function formatSteppedValue(value, valueOptions) {
        if (valueOptions.integer) value = Math.round(value);
        return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
    }

    /**
     * 小数第2位で丸めた数値を文字列で返す
     * @param {number} value - 数値
     * @returns {string} 表示用の数値文字列
     */
    function formatStepperNumber(value) {
        return String(Math.round(value * 100) / 100);
    }

    // -----------------------------------------
    // ∧∨ボタンの描画 / Drawing
    // -----------------------------------------
    /**
     * 山形（∧／∨）の極小ボタンを作成する。
     * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
     * 継ぎ目に線は引かない
     * @param {Group|Panel} parent - 追加先
     * @param {string} direction - "up" または "down"
     * @param {Function} onClickFn - クリック時の処理
     * @returns {Group} ボタンとして使う group
     */
    function makeStepperChevronButton(parent, direction, onClickFn) {
        var buttonWidth = STEPPER_BUTTON_WIDTH;
        var buttonHeight = STEPPER_BUTTON_HEIGHT;
        var isUp = (direction === "up");
        var chevronBox = parent.add("group");
        chevronBox.margins = 0;
        chevronBox.spacing = 0;
        chevronBox.preferredSize = [buttonWidth, buttonHeight];
        chevronBox.minimumSize = [buttonWidth, buttonHeight];
        chevronBox.maximumSize = [buttonWidth, buttonHeight];
        chevronBox.isPressed = false;
        chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

        chevronBox.onDraw = function () {
            var boxGraphics = chevronBox.graphics;
            /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
               Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
            var isDimmed = !isStepperEnabledInTree(chevronBox);

            /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
            var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
            boxGraphics.newPath();
            boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
            boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

            drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
            drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
        };

        /**
         * 押下状態を変えて描き直す
         * @param {boolean} isPressed - 押下中なら true
         * @returns {void}
         */
        function repaint(isPressed) {
            if (chevronBox.isPressed === isPressed) return;
            chevronBox.isPressed = isPressed;
            redrawStepperGroup(chevronBox);
        }
        chevronBox.addEventListener("mousedown", function () {
            if (!isStepperEnabledInTree(chevronBox)) return;
            repaint(true);
            if (onClickFn) onClickFn();
        });
        chevronBox.addEventListener("mouseup", function () { repaint(false); });
        /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
        chevronBox.addEventListener("mouseout", function () { repaint(false); });
        return chevronBox;
    }

    /**
     * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
     * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
     * @param {number[]} frameColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
        var frameLeft = 0.5;
        var frameRight = boxWidth - 0.5;
        var outerY = isUp ? 0.5 : boxHeight - 0.5;
        var seamY = isUp ? boxHeight : 0;
        var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
        var radius = STEPPER_CORNER_RADIUS;
        var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
        var angle, k;

        boxGraphics.newPath();
        boxGraphics.moveTo(frameLeft, seamY);
        /* 左の角丸 / left corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
        }
        /* 右の角丸 / right corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
        }
        boxGraphics.lineTo(frameRight, seamY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
    }

    /**
     * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - ∧なら true、∨なら false
     * @param {number[]} chevronColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
        var centerX = boxWidth / 2;
        var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
        var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
        var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
        boxGraphics.newPath();
        boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
        boxGraphics.lineTo(centerX, centerY + tipOffsetY);
        boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isStepperEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
     * @param {Object} container - 行・グループ・パネルなど
     * @returns {void}
     */
    function redrawSteppersIn(container) {
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            var child = container.children[i];
            if (child.isStepperButton) redrawStepperGroup(child);
            else redrawSteppersIn(child);
        }
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} targetGroup - 描き直す group
     * @returns {void}
     */
    function redrawStepperGroup(targetGroup) {
        targetGroup.hide();
        targetGroup.show();
    }

    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

    // =========================================
    // 選択の読み取り / Reading the selection
    // =========================================

    /* 文字系のオブジェクト（parentStory を持つ）/ Text objects that have parentStory */
    var TEXT_TYPES = {
        Character: true, Word: true, Line: true, Paragraph: true, TextColumn: true,
        TextStyleRange: true, Text: true, InsertionPoint: true, Story: true
    };

    /**
     * 選択からストーリーを集める（同じストーリーは1回だけ）
     * @param {Array} items - 選択、またはグループ内のページアイテム
     * @param {Array<Story>} stories - 集めたストーリーを足す配列
     * @param {Object} storyIds - 足したストーリーの ID（重複を除く）
     * @returns {void}
     */
    function collectStories(items, stories, storyIds) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            var typeName;
            try { typeName = item.constructor.name; } catch (e) { continue; }

            var story = null;
            if (typeName === "Story") story = item;
            else if (typeName === "TextFrame" || TEXT_TYPES[typeName]) story = item.parentStory;
            else if (typeName === "Table") story = item.storyOffset.parentStory;
            else if (typeName === "Cell") story = item.parent.storyOffset.parentStory; /* Cell の親は Table / a Cell's parent is its Table */
            else if (typeName === "Group") collectStories(item.pageItems.everyItem().getElements(), stories, storyIds);

            if (!story || !story.isValid || storyIds[story.id]) continue;
            storyIds[story.id] = true;
            stories.push(story);
        }
    }

    /**
     * 今の選択を返す（ドキュメントが無ければ空の配列）
     * @returns {Array} app.selection
     */
    function getCurrentSelection() {
        if (app.documents.length === 0) return [];
        return app.selection;
    }

    /**
     * 選択した文字を返す（1段落に収まる文字だけ）
     * @param {Array} items - 選択
     * @returns {string} 選択した文字。文字を選んでいなければ空文字
     */
    function getSelectedKeyword(items) {
        if (items.length !== 1) return "";
        var typeName = items[0].constructor.name;
        if (!TEXT_TYPES[typeName] || typeName === "InsertionPoint" || typeName === "Story") return "";
        var selectedText = String(items[0].contents);
        if (/[\r\n]/.test(selectedText)) return "";
        return selectedText;
    }

    // =========================================
    // 前回の設定 / Last settings
    // =========================================

    /**
     * @typedef {Object} CountSettings
     * @property {string[]} keywords - キーワード欄の文字（KEYWORD_FIELD_COUNT 個。空欄は空文字）
     * @property {string} scope - "story" または "document"
     * @property {boolean} caseSensitive - 大文字と小文字を区別するか
     * @property {boolean} wholeWord - 単語単位で探すか
     * @property {boolean} masterPages - 親ページを含むか（ドキュメントのときだけ）
     * @property {boolean} excludeContained - ほかのキーワードに含まれる分を除くか
     * @property {number} contextLength - 一覧に出す前後の文字数
     */

    /**
     * 前回の設定を読む
     * @returns {CountSettings} 設定
     */
    function loadSettings() {
        var saved = $.global[SESSION_KEY] || {};
        var keywords = [];
        for (var i = 0; i < KEYWORD_FIELD_COUNT; i++) {
            keywords.push((saved.keywords && saved.keywords[i] != null) ? String(saved.keywords[i]) : "");
        }
        return {
            keywords: keywords,
            scope: saved.scope || "story",
            caseSensitive: !!saved.caseSensitive,
            wholeWord: !!saved.wholeWord,
            masterPages: !!saved.masterPages,
            excludeContained: (saved.excludeContained == null) ? true : !!saved.excludeContained,
            contextLength: (typeof saved.contextLength === "number") ? saved.contextLength : DEFAULT_CONTEXT_LENGTH
        };
    }

    /**
     * 設定を次回のために控える
     * @param {CountSettings} settings - 設定
     * @returns {void}
     */
    function saveSettings(settings) {
        $.global[SESSION_KEY] = settings;
    }

    // =========================================
    // 検索 / Search
    // =========================================

    /* 検索時に一時的に変える検索オプション（終わったら元に戻す）/ Find options changed for the search and restored afterwards */
    var FIND_OPTION_KEYS = [
        "caseSensitive", "wholeWord", "includeMasterPages", "includeFootnotes",
        "includeHiddenLayers", "includeLockedLayersForFind", "includeLockedStoriesForFind"
    ];

    /**
     * キーワードを文字通りに探せる形にする（^ は検索のメタ文字なので ^^ にする）
     * @param {string} keyword - キーワード
     * @returns {string} 検索文字列
     */
    function toLiteralFindWhat(keyword) {
        return keyword.split("^").join("^^");
    }

    /**
     * @typedef {Object} MatchRange
     * @property {Text} text - 見つかった文字
     * @property {string} container - 入れ物（ストーリー・セル・脚注）の指定子
     * @property {number} start - 入れ物の中での開始位置
     * @property {number} end - 入れ物の中での終了位置（この位置は含まない）
     */

    /**
     * 見つかった文字の位置を、入れ物（ストーリー・セル・脚注）と範囲で表す。
     * findText の結果の index は入れ物の中での位置（2026-11-05 実測）
     * @param {Text} foundText - findText で見つかった文字
     * @returns {MatchRange} 見つかった文字と、入れ物の指定子・範囲
     */
    function toMatchRange(foundText) {
        return { text: foundText, container: foundText.parent.toSpecifier(), start: foundText.index, end: foundText.index + foundText.length };
    }

    /**
     * キーワードごとに見つかった位置を集める
     * @param {Array<Document|Story>} searchTargets - 検索する対象
     * @param {string[]} keywords - キーワード（空文字は探さない）
     * @param {CountSettings} settings - 設定
     * @returns {Array<Array<MatchRange>>} キーワードごとの位置（keywords と同じ並び）
     */
    function findKeywordRanges(searchTargets, keywords, settings) {
        var findOptions = app.findChangeTextOptions;
        var savedOptions = {};
        for (var i = 0; i < FIND_OPTION_KEYS.length; i++) {
            savedOptions[FIND_OPTION_KEYS[i]] = findOptions[FIND_OPTION_KEYS[i]];
        }

        var rangesByKeyword = [];
        try {
            findOptions.caseSensitive = settings.caseSensitive;
            findOptions.wholeWord = settings.wholeWord;
            findOptions.includeMasterPages = (settings.scope === "document") && settings.masterPages;
            findOptions.includeFootnotes = true;
            findOptions.includeHiddenLayers = true;
            findOptions.includeLockedLayersForFind = true;
            findOptions.includeLockedStoriesForFind = true;

            for (var k = 0; k < keywords.length; k++) {
                var ranges = [];
                rangesByKeyword.push(ranges);
                if (keywords[k] === "") continue;
                app.findTextPreferences = NothingEnum.NOTHING;
                app.findTextPreferences.findWhat = toLiteralFindWhat(keywords[k]);
                for (var j = 0; j < searchTargets.length; j++) {
                    var found = searchTargets[j].findText();
                    for (var m = 0; m < found.length; m++) ranges.push(toMatchRange(found[m]));
                }
            }
        } finally {
            /* 検索パネルの状態を元に戻す / Restore the Find/Change state */
            app.findTextPreferences = NothingEnum.NOTHING;
            for (var optionKey in savedOptions) {
                if (savedOptions.hasOwnProperty(optionKey)) findOptions[optionKey] = savedOptions[optionKey];
            }
        }
        return rangesByKeyword;
    }

    /**
     * 位置を入れ物ごとの配列にまとめる
     * @param {Array<MatchRange>} ranges - 位置
     * @returns {Object} 入れ物の指定子をキーにした位置の配列
     */
    function groupRangesByContainer(ranges) {
        var grouped = {};
        for (var i = 0; i < ranges.length; i++) {
            var key = ranges[i].container;
            if (!grouped.hasOwnProperty(key)) grouped[key] = [];
            grouped[key].push(ranges[i]);
        }
        return grouped;
    }

    /**
     * キーワードごとに数える位置を返す。excludeContained なら、ほかのキーワードの一致の中に収まる分を除く
     * （「取引き」の中の「取引」は「取引」に数えない）
     * @param {Array<Array<MatchRange>>} rangesByKeyword - findKeywordRanges() の戻り値
     * @param {string[]} keywords - キーワード
     * @param {boolean} excludeContained - ほかのキーワードに含まれる分を除くか
     * @returns {Array<Array<MatchRange>>} キーワードごとに数える位置
     */
    function filterRanges(rangesByKeyword, keywords, excludeContained) {
        if (!excludeContained) return rangesByKeyword;
        var keptByKeyword = [];
        for (var k = 0; k < rangesByKeyword.length; k++) {
            var ranges = rangesByKeyword[k];

            /* 自分より長い別のキーワードの一致だけが、自分の一致を包める / Only matches of a different, longer keyword can enclose this one */
            var longerRanges = [];
            for (var other = 0; other < rangesByKeyword.length; other++) {
                if (other === k || keywords[other].length <= keywords[k].length) continue;
                longerRanges = longerRanges.concat(rangesByKeyword[other]);
            }
            var longerByContainer = groupRangesByContainer(longerRanges);

            var keptRanges = [];
            for (var i = 0; i < ranges.length; i++) {
                var candidates = longerByContainer.hasOwnProperty(ranges[i].container) ? longerByContainer[ranges[i].container] : [];
                var isContained = false;
                for (var c = 0; c < candidates.length; c++) {
                    if (candidates[c].start <= ranges[i].start && ranges[i].end <= candidates[c].end) { isContained = true; break; }
                }
                if (!isContained) keptRanges.push(ranges[i]);
            }
            keptByKeyword.push(keptRanges);
        }
        return keptByKeyword;
    }

    /**
     * 今の選択とドキュメントで数える
     * @param {CountSettings} settings - 設定
     * @returns {{matchesByKeyword: Array<Array<MatchRange>>|null, message: string}} キーワードごとの位置（数えられなければ null）と合計の行に出す文言
     */
    function runCount(settings) {
        if (app.documents.length === 0) return { matchesByKeyword: null, message: getLabel(LABELS.status.noDocument) };

        var searchTargets;
        if (settings.scope === "story") {
            searchTargets = [];
            collectStories(app.selection, searchTargets, {});
            if (searchTargets.length === 0) return { matchesByKeyword: null, message: getLabel(LABELS.status.noStory) };
        } else {
            searchTargets = [app.activeDocument];
        }

        /* パレット表示中に選択やドキュメントが変わるので、失敗は合計の行に出す / The selection or document may change while the palette is open, so report failures in the total row */
        try {
            var rangesByKeyword = findKeywordRanges(searchTargets, settings.keywords, settings);
            var matchesByKeyword = filterRanges(rangesByKeyword, settings.keywords, settings.excludeContained);
            var totalCount = 0;
            for (var i = 0; i < matchesByKeyword.length; i++) totalCount += matchesByKeyword[i].length;
            return { matchesByKeyword: matchesByKeyword, message: getLabel(LABELS.status.total, { count: totalCount, scope: getLabel(LABELS.radio[settings.scope]) }) };
        } catch (e) {
            return { matchesByKeyword: null, message: getLabel(LABELS.status.failed, { message: e.message }) };
        }
    }

    // =========================================
    // 見つかった場所の一覧 / Match list
    // =========================================

    /**
     * 見つかった文字のあるページ名を返す
     * @param {Text} foundText - 見つかった文字
     * @returns {string} ページ名（親ページは「A-親ページ」の形）。オーバーセット・ペーストボードはその旨
     */
    function getPageLabel(foundText) {
        var textFrames = foundText.parentTextFrames;
        if (textFrames.length === 0) return getLabel(LABELS.status.overset);
        var page = textFrames[0].parentPage;
        if (!page) return getLabel(LABELS.status.pasteboard);
        return (page.parent.constructor.name === "MasterSpread") ? page.parent.name : page.name;
    }

    /**
     * 一覧に出せるように文字列を整える（改行は ¶、タブは空白、ほかの制御文字や目印の文字は除く）
     * @param {string} rawText - 文字列
     * @returns {string} 整えた文字列
     */
    function toDisplayText(rawText) {
        var displayText = "";
        for (var i = 0; i < rawText.length; i++) {
            var charCode = rawText.charCodeAt(i);
            if (charCode === 13 || charCode === 10) displayText += "¶";
            else if (charCode === 9) displayText += " ";
            /* 制御文字（ページ番号・脚注番号の目印など）と U+FEFF（アンカーの目印）は出さない / skip control characters (page and footnote markers) and U+FEFF (anchor markers) */
            else if (charCode >= 32 && charCode !== 0xFEFF) displayText += rawText.charAt(i);
        }
        return displayText;
    }

    /**
     * 入れ物の中の範囲の文字列を返す
     * @param {Story|Cell|Footnote} container - 入れ物
     * @param {number} firstIndex - 最初の文字の位置
     * @param {number} lastIndex - 最後の文字の位置（この位置も含む）
     * @returns {string} 文字列。範囲が空なら空文字
     */
    function getRangeText(container, firstIndex, lastIndex) {
        if (lastIndex < firstIndex) return "";
        /* itemByRange の contents は文字列の配列で返る（2026-11-05 実測）/ itemByRange().contents comes back as an array of strings */
        var rangeContents = container.characters.itemByRange(firstIndex, lastIndex).contents;
        return (rangeContents instanceof Array) ? rangeContents.join("") : String(rangeContents);
    }

    /**
     * 見つかった文字を前後の文字列と合わせて返す（見つかった文字は【】で囲む）
     * @param {MatchRange} matchRange - 見つかった位置
     * @param {number} contextLength - 前後の文字数
     * @param {Object} containerLengths - 入れ物ごとの文字数（同じ入れ物を何度も数えないための控え）
     * @returns {string} 一覧に出す文字列
     */
    function buildContextText(matchRange, contextLength, containerLengths) {
        var container = matchRange.text.parent;
        if (!containerLengths.hasOwnProperty(matchRange.container)) containerLengths[matchRange.container] = container.characters.length;
        var containerLength = containerLengths[matchRange.container];
        var beforeText = getRangeText(container, Math.max(0, matchRange.start - contextLength), matchRange.start - 1);
        var afterText = getRangeText(container, matchRange.end, Math.min(containerLength - 1, matchRange.end + contextLength - 1));
        return toDisplayText(beforeText) + "【" + toDisplayText(String(matchRange.text.contents)) + "】" + toDisplayText(afterText);
    }

    // =========================================
    // パレット / Palette
    // =========================================

    /**
     * パレットを組み立てる（左に設定、右に見つかった場所の一覧）
     * @param {CountSettings} settings - 初期値
     * @returns {Window} パレット
     */
    function buildPalette(settings) {
        var countPalette = new Window("palette", getLabel(LABELS.palette.title) + " " + SCRIPT_VERSION);
        setupWindow(countPalette);

        /* 2カラム（左右とも高さをそろえる）/ Two columns of the same height */
        var columnsGroup = countPalette.add("group");
        setupRow(columnsGroup, ["fill", "fill"], COLUMN_SPACING);
        columnsGroup.alignChildren = ["fill", "fill"];

        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];
        leftColumn.spacing = WINDOW_SPACING;

        /* キーワード（欄の右に件数）/ Keywords, each with its count on the right */
        var keywordsPanel = leftColumn.add("panel", undefined, getLabel(LABELS.panel.keywords));
        setupPanel(keywordsPanel, 6);
        var keywordInputs = [];
        var countTexts = [];
        var lastActiveIndex = -1;
        for (var i = 0; i < KEYWORD_FIELD_COUNT; i++) {
            var keywordRowGroup = keywordsPanel.add("group");
            setupRow(keywordRowGroup, "fill", 6);
            var keywordInput = keywordRowGroup.add("edittext", undefined, settings.keywords[i]);
            keywordInput.preferredSize.width = KEYWORD_FIELD_WIDTH;
            keywordInput.alignment = ["fill", "center"];
            keywordInput.helpTip = getLabel(LABELS.tooltip.keyword);
            var countText = keywordRowGroup.add("statictext", undefined, getLabel(LABELS.status.initial));
            countText.preferredSize.width = COUNT_TEXT_WIDTH;
            countText.justify = "right";
            keywordInputs.push(keywordInput);
            countTexts.push(countText);
        }

        /* 合計 / Total */
        var totalRowGroup = keywordsPanel.add("group");
        setupRow(totalRowGroup, "fill", 6);
        totalRowGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.total));
        var totalText = totalRowGroup.add("statictext", undefined, getLabel(LABELS.status.initial));
        totalText.alignment = ["fill", "center"];

        /* 対象（ラジオは横並び）/ Scope (radios in a row) */
        var scopePanel = leftColumn.add("panel", undefined, getLabel(LABELS.panel.scope));
        setupPanel(scopePanel);
        scopePanel.orientation = "row";
        scopePanel.alignChildren = ["left", "center"];
        var scopeRadios = {
            story: scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.story)),
            document: scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.document))
        };
        scopeRadios.story.helpTip = getLabel(LABELS.tooltip.story);
        scopeRadios.document.helpTip = getLabel(LABELS.tooltip.document);
        scopeRadios[scopeRadios[settings.scope] ? settings.scope : "story"].value = true;

        /* オプション / Options */
        var optionsPanel = leftColumn.add("panel", undefined, getLabel(LABELS.panel.options));
        setupPanel(optionsPanel, 6);
        var caseSensitiveCheckbox = optionsPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.caseSensitive));
        var wholeWordCheckbox = optionsPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.wholeWord));
        var excludeContainedCheckbox = optionsPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.excludeContained));
        excludeContainedCheckbox.helpTip = getLabel(LABELS.tooltip.excludeContained);
        var masterPagesCheckbox = optionsPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.masterPages));
        masterPagesCheckbox.helpTip = getLabel(LABELS.tooltip.masterPages);
        caseSensitiveCheckbox.value = settings.caseSensitive;
        wholeWordCheckbox.value = settings.wholeWord;
        excludeContainedCheckbox.value = settings.excludeContained;
        masterPagesCheckbox.value = settings.masterPages;

        /* 見つかった場所（右のカラム。左のカラムと同じ高さまで伸ばす）/ Matches (right column, stretched to the left column's height) */
        var matchesPanel = columnsGroup.add("panel", undefined, getLabel(LABELS.panel.matches));
        setupPanel(matchesPanel, 6);
        matchesPanel.alignment = ["fill", "fill"];
        var contextLengthInput = addSteppedField(matchesPanel, {
            label: labelText(LABELS.fieldLabel.contextLength),
            text: String(settings.contextLength), characters: 4, step: 1, min: 0, max: MAX_CONTEXT_LENGTH, integer: true,
            onStep: function () { renderMatchList(); }
        });
        contextLengthInput.helpTip = getLabel(LABELS.tooltip.contextLength);
        /* 1列目は短いページ名にする（Mac は1列目の長い文字列で列幅が崩れる）/ Keep the first column short (long text in column 1 breaks the widths on Mac) */
        var matchList = matchesPanel.add("listbox", undefined, [], {
            numberOfColumns: 2,
            showHeaders: true,
            columnTitles: [getLabel(LABELS.column.page), getLabel(LABELS.column.context)],
            columnWidths: [MATCH_LIST_PAGE_COLUMN_WIDTH, MATCH_LIST_WIDTH - MATCH_LIST_PAGE_COLUMN_WIDTH - 20]
        });
        matchList.preferredSize = [MATCH_LIST_WIDTH, MATCH_LIST_MIN_HEIGHT];
        matchList.alignment = ["fill", "fill"];
        matchList.helpTip = getLabel(LABELS.tooltip.matchList);
        var listNoteText = matchesPanel.add("statictext", undefined, " ");
        listNoteText.alignment = ["fill", "bottom"];

        /* ボタン：左に［選択文字を入れる］、右に［閉じる］［数える］ / Buttons: Use Selection on the left, Close and Count on the right */
        var buttonRow = addButtonRow(countPalette);
        var btnUseSelection = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.useSelection));
        btnUseSelection.helpTip = getLabel(LABELS.tooltip.useSelection);
        var btnClose = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.close));
        var btnCount = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.count));
        alignRightOnlyButtonRow(buttonRow);

        /* 最後に数えた結果（前後の文字数を変えたら、数え直さずに一覧だけ作り直す）/ Last results; changing the context length rebuilds only the list */
        var lastMatchesByKeyword = null;

        /**
         * 前後の文字数の欄の値を返す
         * @returns {number} 前後の文字数（0〜MAX_CONTEXT_LENGTH の整数）
         */
        function readContextLength() {
            var contextLength = parseInt(contextLengthInput.text, 10);
            if (isNaN(contextLength) || contextLength < 0) return 0;
            return Math.min(contextLength, MAX_CONTEXT_LENGTH);
        }

        /**
         * 画面の値を設定にまとめる
         * @returns {CountSettings} 設定
         */
        function readSettings() {
            var keywords = [];
            for (var j = 0; j < keywordInputs.length; j++) keywords.push(keywordInputs[j].text);
            return {
                keywords: keywords,
                scope: scopeRadios.story.value ? "story" : "document",
                caseSensitive: caseSensitiveCheckbox.value,
                wholeWord: wholeWordCheckbox.value,
                masterPages: masterPagesCheckbox.value,
                excludeContained: excludeContainedCheckbox.value,
                contextLength: readContextLength()
            };
        }

        /**
         * キーワード欄のどれかに文字があるかを返す
         * @returns {boolean} 1つでもあれば true
         */
        function hasAnyKeyword() {
            for (var j = 0; j < keywordInputs.length; j++) {
                if (keywordInputs[j].text.length > 0) return true;
            }
            return false;
        }

        /**
         * 親ページの項目はドキュメントを対象にしたときだけ、［数える］はキーワードがあるときだけ押せるようにする
         * @returns {void}
         */
        function updateEnabled() {
            masterPagesCheckbox.enabled = scopeRadios.document.value;
            btnCount.enabled = hasAnyKeyword();
        }

        /**
         * 見つかった場所の一覧を作り直す（キーワード欄の順、その中は見つかった順。先頭の MAX_LIST_ITEMS 件まで）
         * @returns {void}
         */
        function renderMatchList() {
            matchList.removeAll();
            listNoteText.text = " ";
            if (!lastMatchesByKeyword) return;

            var contextLength = readContextLength();
            var containerLengths = {};
            var totalCount = 0;
            var shownCount = 0;
            for (var k = 0; k < lastMatchesByKeyword.length; k++) {
                var matches = lastMatchesByKeyword[k];
                totalCount += matches.length;
                for (var m = 0; m < matches.length && shownCount < MAX_LIST_ITEMS; m++) {
                    var pageLabel, contextText;
                    /* パレット表示中に文字が消されたり、ドキュメントが閉じられたりする / the text or document may be gone while the palette is open */
                    try {
                        pageLabel = getPageLabel(matches[m].text);
                        contextText = buildContextText(matches[m], contextLength, containerLengths);
                    } catch (e) {
                        pageLabel = "-";
                        contextText = getLabel(LABELS.status.unavailable);
                    }
                    var listItem = matchList.add("item", pageLabel);
                    listItem.subItems[0].text = contextText;
                    listItem.foundText = matches[m].text;
                    shownCount++;
                }
            }
            if (shownCount < totalCount) listNoteText.text = getLabel(LABELS.status.truncated, { total: totalCount, shown: shownCount });
        }

        /**
         * 数えて、件数を各欄の右と合計の行に、見つかった場所を右の一覧に出す
         * @returns {void}
         */
        function countAndShow() {
            if (!hasAnyKeyword()) return;
            var currentSettings = readSettings();
            saveSettings(currentSettings);
            var countResult = runCount(currentSettings);
            for (var j = 0; j < countTexts.length; j++) {
                var hasCount = countResult.matchesByKeyword && currentSettings.keywords[j] !== "";
                countTexts[j].text = hasCount ? getLabel(LABELS.status.count, { count: countResult.matchesByKeyword[j].length }) : getLabel(LABELS.status.initial);
            }
            totalText.text = countResult.message;
            lastMatchesByKeyword = countResult.matchesByKeyword;
            renderMatchList();
        }

        /**
         * 選択文字を入れる欄を決める（最後に入力した欄、なければ最初の空欄、空欄も無ければ1つ目）
         * @returns {EditText} キーワード欄
         */
        function getInsertTargetInput() {
            if (lastActiveIndex >= 0) return keywordInputs[lastActiveIndex];
            for (var j = 0; j < keywordInputs.length; j++) {
                if (keywordInputs[j].text === "") return keywordInputs[j];
            }
            return keywordInputs[0];
        }

        /**
         * キーワード欄のイベントを付ける
         * @param {EditText} keywordInput - キーワード欄
         * @param {number} fieldIndex - 何番目の欄か
         * @returns {void}
         */
        function bindKeywordInput(keywordInput, fieldIndex) {
            keywordInput.onChanging = function () {
                lastActiveIndex = fieldIndex;
                updateEnabled();
            };
            keywordInput.onActivate = function () { lastActiveIndex = fieldIndex; };
            /* キーワード欄の return で数える / Return in a keyword field counts */
            keywordInput.addEventListener("keydown", function (event) {
                if (event.keyName !== "Enter") return;
                event.preventDefault();
                countAndShow();
            });
        }
        for (var n = 0; n < keywordInputs.length; n++) bindKeywordInput(keywordInputs[n], n);

        /* 前後の文字数を直接入力したときも一覧を作り直す（値をそろえる部品の onChange のあと）/ Rebuild the list after typing too (after the part's own onChange) */
        var normalizeContextLength = contextLengthInput.onChange;
        contextLengthInput.onChange = function () {
            normalizeContextLength();
            renderMatchList();
        };

        /* 一覧で選んだ文字を選択して表示する / Select and show the text picked in the list */
        matchList.onChange = function () {
            var listItem = matchList.selection;
            if (!listItem || !listItem.foundText) return;
            try {
                if (!listItem.foundText.isValid) return;
                listItem.foundText.select();
                listItem.foundText.showText();
            } catch (e) { /* 表示できない場所（閉じたドキュメントなど）は何もしない / nothing to show (e.g. a closed document) */ }
        };

        scopeRadios.story.onClick = updateEnabled;
        scopeRadios.document.onClick = updateEnabled;
        btnCount.onClick = countAndShow;
        btnUseSelection.onClick = function () {
            var selectedKeyword = getSelectedKeyword(getCurrentSelection());
            if (selectedKeyword === "") return;
            getInsertTargetInput().text = selectedKeyword;
            updateEnabled();
        };
        btnClose.onClick = function () { countPalette.close(); };

        /* パレットは Esc で閉じないので、［閉じる］と同じにする / palettes do not close on Esc, so treat it as Close */
        countPalette.addEventListener("keydown", function (event) {
            if (event.keyName !== "Escape") return;
            event.preventDefault();
            countPalette.close();
        });

        /* 閉じるときは設定を控えて参照を外すだけ（DOM には触れない）/ On close, only keep the settings and drop the reference (no DOM access) */
        countPalette.onClose = function () {
            saveSettings(readSettings());
            $.global[PALETTE_GLOBAL_KEY] = null;
        };

        countPalette.onShow = function () {
            getInsertTargetInput().active = true;
        };

        updateEnabled();
        return countPalette;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択した文字を、まだ入っていなければ最初の空欄に入れる
     * @param {CountSettings} settings - 設定（keywords を書き換える）
     * @param {string} selectedKeyword - 選択した文字
     * @returns {void}
     */
    function addSelectedKeyword(settings, selectedKeyword) {
        if (selectedKeyword === "") return;
        for (var i = 0; i < settings.keywords.length; i++) {
            if (settings.keywords[i] === selectedKeyword) return;
        }
        for (var j = 0; j < settings.keywords.length; j++) {
            if (settings.keywords[j] === "") { settings.keywords[j] = selectedKeyword; return; }
        }
    }

    /**
     * パレットを開く（前回のパレットが開いていれば閉じて作り直す）
     * @returns {void}
     */
    function main() {
        if ($.global[PALETTE_GLOBAL_KEY]) $.global[PALETTE_GLOBAL_KEY].close();

        var settings = loadSettings();
        addSelectedKeyword(settings, getSelectedKeyword(getCurrentSelection()));

        var countPalette = buildPalette(settings);
        /* パレットは show() から即座に戻るので、参照を $.global に残す / palettes return from show() at once, so keep a reference */
        $.global[PALETTE_GLOBAL_KEY] = countPalette;
        countPalette.show();
    }

    main();

})();
