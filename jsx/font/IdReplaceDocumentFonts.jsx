#target indesign

/*

### 概要

ドキュメントで使用中のフォントをファミリー／スタイル単位で一覧し、
選んだフォントを別のフォントへまとめて置き換えます。
段落スタイル・文字スタイルのフォントも同時に更新できます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdReplaceDocumentFonts.md

### Overview

Lists the fonts used in the document by family and style, and replaces
the selected ones with another font in a single pass. Paragraph and
character styles can be updated at the same time.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdReplaceDocumentFonts.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdReplaceDocumentFonts";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdReplaceDocumentFonts.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdReplaceDocumentFonts.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* スタイル行の字下げ / Indent used for style rows */
    var STYLE_ROW_INDENT = "　　";

    /* PostScript名表示の初期状態 / Initial state of the PostScript-name display */
    var SHOW_POSTSCRIPT_NAME_DEFAULT = false;

    /* ［段落・文字スタイルも更新］の初期状態 / Initial state of the style-update option */
    var UPDATE_STYLES_DEFAULT = true;

    /* 使用箇所の件数を表示するか。フォントごとに検索するため、大きなドキュメントでは一覧の作成に時間がかかる
       Whether to show usage counts. Each font is searched separately, so building the list is slower on large documents */
    var SHOW_USAGE_COUNT = true;

    /* 検索の対象範囲 / Scope of the search */
    var INCLUDE_MASTER_PAGES  = true; /* 親ページ / master pages */
    var INCLUDE_HIDDEN_LAYERS = true; /* 非表示レイヤー / hidden layers */
    var INCLUDE_FOOTNOTES     = true; /* 脚注 / footnotes */

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

    var LIST_LABEL_SPACING    = 6;    /* 見出しとリストの間隔 / spacing between a label and its list */
    var OPTION_SPACING        = 4;    /* チェックボックスの間隔 / spacing between checkboxes */
    var LISTBOX_HEIGHT        = 300;  /* リストの高さ / list height */
    var LISTBOX_WIDTH_MIN     = 200;  /* リスト幅の下限 / minimum list width */
    var LISTBOX_WIDTH_MAX     = 600;  /* リスト幅の上限 / maximum list width */
    var LISTBOX_CHAR_WIDTH    = 9;    /* 1文字あたりの概算幅 / approximate width per character */
    var LISTBOX_WIDTH_PADDING = 60;   /* リスト幅の余裕 / extra width added to the list */

    // =========================================
    // ローカライズ / Localization
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

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
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

    var LABELS = {
        dialog: {
            title: { ja: "フォント置換", en: "Replace Fonts" }
        },
        fieldLabel: {
            sourceFonts: { ja: "置換元フォント（複数選択可）", en: "Source Fonts (Multiple Selection)" },
            targetFont: { ja: "置換先フォント", en: "Target Font" }
        },
        checkbox: {
            postScriptName: { ja: "PostScript名で表示", en: "Show PostScript names" },
            updateStyles: { ja: "段落・文字スタイルも更新", en: "Update paragraph and character styles" }
        },
        button: {
            close: { ja: "閉じる", en: "Close" },
            replaceAll: { ja: "全置換", en: "Replace All" },
            replace: { ja: "フォントを置換", en: "Replace Fonts" }
        },
        list: {
            notInstalled: { ja: "未インストール", en: "not installed" }
        },
        tooltip: {
            sourceFonts: {
                ja: "置換元のフォントを選びます。ファミリー名の行を選ぶと、そのファミリーのスタイルがすべて選ばれます。",
                en: "Pick the fonts to replace. Selecting a family row selects every style in that family."
            },
            targetFont: {
                ja: "置換先のフォントを選びます。ファミリー名の行は選べません。",
                en: "Pick the font to replace them with. Family rows cannot be selected."
            },
            postScriptName: {
                ja: "ファミリー名とスタイル名の代わりに、PostScript名で一覧します。",
                en: "List the fonts by PostScript name instead of family and style."
            },
            updateStyles: {
                ja: "テキストだけでなく、段落スタイル・文字スタイルに設定されたフォントも置き換えます。",
                en: "Replace the font in paragraph and character styles as well as in the text itself."
            },
            replaceAll: {
                ja: "使用中のすべてのフォントを置換先フォントに置き換えます。置換先を選んでいないときは、置換元の1つ目のフォントに揃えます。",
                en: "Replace every font in use with the target font. With no target selected, the first source font is used instead."
            },
            replace: {
                ja: "選んだ置換元フォントを、置換先フォントに置き換えます。",
                en: "Replace the selected source fonts with the target font."
            }
        },
        alert: {
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            noFontsFound: {
                ja: "ドキュメント内に使用中のフォントが見つかりません。",
                en: "No fonts in use were found in the document."
            },
            noSourceFont: {
                ja: "置換元フォントを1つ以上選択してください。",
                en: "Please select at least one source font."
            },
            noTargetFont: {
                ja: "置換先フォントを選択してください。",
                en: "Please select a target font."
            },
            selectFonts: {
                ja: "置換先または置換元フォントを選択してください。\n（または、置換元だけ選んで全置換することも可能です）",
                en: "Please select either a target font or source fonts.\n(Alternatively, select only the source fonts to replace all.)"
            },
            targetNotFound: {
                ja: "置換先フォントが見つかりません（%1）。",
                en: "Target font not found (%1)."
            }
        }
    };

    // =========================================
    // 状態 / State
    // =========================================

    var doc = null;
    var mainDialog = null;
    var sourceFontListBox = null;
    var targetFontListBox = null;
    var postScriptNameCheckbox = null;
    var updateStylesCheckbox = null;

    /* 収集した使用中フォント / Fonts collected from the document */
    var usedFontMap = {};

    /* リストに並べるフォント（ファミリー見出しを含む）/ Fonts listed in the boxes, family headers included */
    var flatFontList = [];

    /* 選択の再帰更新を防ぐフラグ / Guard against recursive selection updates */
    var isUpdatingSelection = false;

    // =========================================
    // 検索・置換の下ごしらえ / Find and change setup
    // =========================================

    /**
     * 検索・置換の設定を初期化する
     * @returns {void}
     */
    function resetFindChangePreferences() {
        app.findTextPreferences = NothingEnum.NOTHING;
        app.changeTextPreferences = NothingEnum.NOTHING;
    }

    /**
     * 検索範囲のオプションをそろえる
     * @param {boolean} includesLocked - ロックされたレイヤー・ストーリーを含めるか
     * @returns {void}
     */
    function setupFindChangeOptions(includesLocked) {
        var options = app.findChangeTextOptions;
        options.includeMasterPages = INCLUDE_MASTER_PAGES;
        options.includeHiddenLayers = INCLUDE_HIDDEN_LAYERS;
        options.includeFootnotes = INCLUDE_FOOTNOTES;
        options.includeLockedLayersForFind = includesLocked;
        options.includeLockedStoriesForFind = includesLocked;
    }

    /**
     * 指定したフォントを使っているテキストを検索する
     * @param {Font} font - 検索するフォント
     * @returns {Array} 見つかったテキストの配列
     */
    function findTextOfFont(font) {
        resetFindChangePreferences();
        setupFindChangeOptions(true);
        app.findTextPreferences.appliedFont = font;
        var found = doc.findText();
        resetFindChangePreferences();
        return found;
    }

    // =========================================
    // フォントの収集 / Collecting fonts
    // =========================================

    /**
     * スタイルに設定されているフォント名を取り出す
     * @param {ParagraphStyle|CharacterStyle} style - 対象スタイル
     * @returns {string} 「ファミリー名＋タブ＋スタイル名」。未設定なら空文字列
     */
    function getStyleFontName(style) {
        var appliedFont = style.appliedFont;
        var familyName = "";

        if (appliedFont && appliedFont.fontFamily) {
            familyName = appliedFont.fontFamily;
        } else if (appliedFont && appliedFont.constructor === String) {
            familyName = String(appliedFont);
        }
        if (familyName === "") return "";

        var styleName = style.fontStyle;
        if (!styleName || styleName.constructor !== String) return "";
        return familyName + "\t" + String(styleName);
    }

    /**
     * ドキュメント内のすべての段落スタイル・文字スタイルを集める
     * @returns {Array} スタイルの配列（スタイルグループ内も含む）
     */
    function collectAllStyles() {
        return doc.allParagraphStyles.concat(doc.allCharacterStyles);
    }

    /**
     * 段落・文字スタイルが使っているフォントを数える
     * @returns {object} フォント名 → 件数のマップ
     */
    function collectStyleFontCounts() {
        var counts = {};
        var styles = collectAllStyles();

        for (var i = 0; i < styles.length; i++) {
            var fontName = getStyleFontName(styles[i]);
            if (fontName === "") continue;
            counts[fontName] = (counts[fontName] || 0) + 1;
        }
        return counts;
    }

    /**
     * PostScript名を取り出す（取得できないときはフォント名で代用）
     * @param {Font} font - 対象フォント
     * @param {string} fallbackName - 取得できないときに使う名前
     * @returns {string} PostScript名
     */
    function getPostScriptName(font, fallbackName) {
        /* 環境にないフォントは PostScript 名を返さないことがある / Fonts that are not installed may not expose a PostScript name */
        var postScriptName = "";
        try {
            postScriptName = font.postscriptName;
        } catch (e) {
            postScriptName = "";
        }
        if (postScriptName && postScriptName.constructor === String) return String(postScriptName);
        return fallbackName.replace("\t", " ");
    }

    /**
     * ドキュメントで使用中のフォントをファミリー別に集める
     * @returns {object} ファミリー名 → フォント名 → フォント情報のマップ
     */
    function collectUsedFonts() {
        var styleFontCounts = collectStyleFontCounts();
        var fonts = doc.fonts.everyItem().getElements();
        var fontMap = {};

        for (var i = 0; i < fonts.length; i++) {
            var font = fonts[i];
            var family = font.fontFamily;
            var styleName = font.fontStyleName;
            var fontName = family + "\t" + styleName;

            if (!fontMap[family]) fontMap[family] = {};
            fontMap[family][fontName] = {
                font: font,
                name: fontName,
                family: family,
                style: styleName,
                postScriptName: getPostScriptName(font, fontName),
                isMissing: (font.status !== FontStatus.INSTALLED),
                textCount: SHOW_USAGE_COUNT ? findTextOfFont(font).length : 0,
                styleCount: styleFontCounts[fontName] || 0
            };
        }
        return fontMap;
    }

    // =========================================
    // リストの組み立て / Building the list
    // =========================================

    /**
     * ［段落・文字スタイルも更新］がオンかを調べる
     * @returns {boolean} オンなら true
     */
    function updatesStyles() {
        return updateStylesCheckbox ? updateStylesCheckbox.value : UPDATE_STYLES_DEFAULT;
    }

    /**
     * 件数と未インストールの印をまとめた接尾辞を作る（日本語は全角括弧、英語は半角括弧）
     * @param {object} fontInfo - フォント情報
     * @returns {string} 括弧付きの接尾辞。示すものがなければ空文字列
     */
    function labelSuffix(fontInfo) {
        var parts = [];

        if (SHOW_USAGE_COUNT) {
            var count = fontInfo.textCount + (updatesStyles() ? fontInfo.styleCount : 0);
            parts.push(String(count));
        }
        if (fontInfo.isMissing) parts.push(getLabel(LABELS.list.notInstalled));
        if (parts.length === 0) return "";

        var body = parts.join(uiLang === "ja" ? "／" : " / ");
        return (uiLang === "ja") ? "（" + body + "）" : " (" + body + ")";
    }

    /**
     * リスト1行分のデータを作る
     * @param {object} fontInfo - フォント情報
     * @param {string} labelBody - 接尾辞を除いた表示名
     * @returns {object} リスト行
     */
    function createFontRow(fontInfo, labelBody) {
        return {
            label: labelBody + labelSuffix(fontInfo),
            name: fontInfo.name,
            family: fontInfo.family,
            style: fontInfo.style,
            font: fontInfo.font,
            isMissing: fontInfo.isMissing,
            isHeader: false
        };
    }

    /**
     * 収集したフォントをリスト表示用の1次元配列にする
     * @param {object} fontMap - collectUsedFonts() が返したマップ
     * @returns {Array<object>} ファミリー見出しとスタイル行を並べた配列（PostScript名表示中は見出しなしの1フォント1行）
     */
    function buildFlatFontList(fontMap) {
        var showsPostScriptName = postScriptNameCheckbox ? postScriptNameCheckbox.value : SHOW_POSTSCRIPT_NAME_DEFAULT;
        var fontList = [];

        for (var family in fontMap) {
            var styles = fontMap[family];
            var fontNames = [];
            for (var fontName in styles) {
                fontNames.push(fontName);
            }

            /* PostScript名表示のときは見出しを立てず、1フォント1行で並べる / In PostScript-name mode, list one row per font with no headers */
            if (showsPostScriptName) {
                for (var p = 0; p < fontNames.length; p++) {
                    var psFont = styles[fontNames[p]];
                    fontList.push(createFontRow(psFont, psFont.postScriptName));
                }
                continue;
            }

            /* スタイルが1つだけのファミリーは見出しを立てず1行で見せる / Show single-style families on one row */
            if (fontNames.length === 1) {
                var onlyFont = styles[fontNames[0]];
                fontList.push(createFontRow(onlyFont, onlyFont.family + " " + onlyFont.style));
                continue;
            }

            fontList.push({ label: family, family: family, isHeader: true });
            for (var i = 0; i < fontNames.length; i++) {
                var font = styles[fontNames[i]];
                fontList.push(createFontRow(font, STYLE_ROW_INDENT + font.style));
            }
        }
        return fontList;
    }

    /**
     * 表示形式に合わせてリストを作り直す（選択はフォント名で引き継ぐ）
     * @returns {void}
     */
    function rebuildFontList() {
        var previousSourceFontNames = getSelectedSourceFontNames();
        var previousTargetFontName = getSelectedTargetFontName();

        flatFontList = buildFlatFontList(usedFontMap);
        populateFontListBoxes();
        restoreSelection(previousSourceFontNames, previousTargetFontName);
    }

    /**
     * フォント一覧を集め直してリストを作り直す
     * @returns {void}
     */
    function reloadFontList() {
        usedFontMap = collectUsedFonts();
        rebuildFontList();
    }

    // =========================================
    // 置換処理 / Replacing fonts
    // =========================================

    /**
     * フォント名から Font を取得する
     * @param {string} fontName - フォント名（ファミリー名＋タブ＋スタイル名）
     * @returns {Font|null} 見つからなければ null
     */
    function findFontByName(fontName) {
        var font = doc.fonts.itemByName(fontName);
        return font.isValid ? font : null;
    }

    /**
     * 置換元として選択されているフォント名を取り出す（見出し行は除く）
     * @returns {Array<string>} フォント名の配列
     */
    function getSelectedSourceFontNames() {
        var fontNames = [];
        if (!sourceFontListBox || !sourceFontListBox.selection) return fontNames;

        for (var i = 0; i < sourceFontListBox.selection.length; i++) {
            var listEntry = flatFontList[sourceFontListBox.selection[i].index];
            if (!listEntry.isHeader) fontNames.push(listEntry.name);
        }
        return fontNames;
    }

    /**
     * 置換先リストでフォント（見出し行以外）が選ばれているか調べる
     * @returns {boolean} 選ばれていれば true
     */
    function hasTargetFontSelection() {
        if (!targetFontListBox || !targetFontListBox.selection) return false;
        return !flatFontList[targetFontListBox.selection.index].isHeader;
    }

    /**
     * 置換先として選択されているフォント名を取り出す（見出し行は除く）
     * @returns {string} 未選択・見出し行のときは空文字列
     */
    function getSelectedTargetFontName() {
        if (!hasTargetFontSelection()) return "";
        return flatFontList[targetFontListBox.selection.index].name;
    }

    /**
     * 置換先として選択されているフォントを取り出す
     * @returns {Font|null} 未選択・見出し行のときは null
     */
    function getSelectedTargetFont() {
        var fontName = getSelectedTargetFontName();
        if (fontName === "") return null;

        var targetFont = findFontByName(fontName);
        if (!targetFont) alert(getLabel(LABELS.alert.targetNotFound, [fontName.replace("\t", " ")]));
        return targetFont;
    }

    /**
     * 使用中フォントの名前をすべて集める
     * @param {string} [excludedFontName] - 除外するフォント名
     * @returns {Array<string>} フォント名の配列
     */
    function collectAllFontNames(excludedFontName) {
        var fontNames = [];
        for (var i = 0; i < flatFontList.length; i++) {
            if (flatFontList[i].isHeader) continue;
            if (flatFontList[i].name === excludedFontName) continue;
            fontNames.push(flatFontList[i].name);
        }
        return fontNames;
    }

    /**
     * テキストに使われているフォントを置き換える（ロックされたレイヤー・ストーリーは対象外）
     * @param {string} sourceFontName - 置換元のフォント名
     * @param {Font} targetFont - 置換先フォント
     * @returns {void}
     */
    function changeTextFont(sourceFontName, targetFont) {
        var sourceFont = findFontByName(sourceFontName);
        if (!sourceFont) return;

        resetFindChangePreferences();
        setupFindChangeOptions(false);
        app.findTextPreferences.appliedFont = sourceFont;
        app.changeTextPreferences.appliedFont = targetFont;
        doc.changeText();
        resetFindChangePreferences();
    }

    /**
     * 段落・文字スタイルに設定されたフォントを置き換える
     * @param {string} sourceFontName - 置換元のフォント名
     * @param {Font} targetFont - 置換先フォント
     * @returns {void}
     */
    function changeStyleFont(sourceFontName, targetFont) {
        var styles = collectAllStyles();

        for (var i = 0; i < styles.length; i++) {
            if (getStyleFontName(styles[i]) !== sourceFontName) continue;
            /* ファミリーとスタイルは別のプロパティなので、それぞれに入れる / Family and style are separate properties, so set both */
            styles[i].appliedFont = targetFont.fontFamily;
            styles[i].fontStyle = targetFont.fontStyleName;
        }
    }

    /**
     * 指定したフォントを置換先フォントに置き換える
     * @param {Array<string>} sourceFontNames - 置換元のフォント名
     * @param {Font} targetFont - 置換先フォント
     * @returns {void}
     */
    function replaceFonts(sourceFontNames, targetFont) {
        if (!targetFont || sourceFontNames.length === 0) return;

        var targetFontName = targetFont.fontFamily + "\t" + targetFont.fontStyleName;
        var updatesStyleFont = updatesStyles();

        for (var i = 0; i < sourceFontNames.length; i++) {
            if (sourceFontNames[i] === targetFontName) continue;
            changeTextFont(sourceFontNames[i], targetFont);
            if (updatesStyleFont) changeStyleFont(sourceFontNames[i], targetFont);
        }
        reloadFontList();
    }

    // =========================================
    // イベントハンドラー / Event handlers
    // =========================================

    /**
     * 置換元リストの選択を整える（見出し行はファミリー内の全スタイルに展開する）
     * @returns {void}
     */
    function handleSourceFontSelection() {
        if (isUpdatingSelection || !sourceFontListBox.selection) return;

        var expandedSelection = [];
        for (var i = 0; i < sourceFontListBox.selection.length; i++) {
            var selectedIndex = sourceFontListBox.selection[i].index;
            var listEntry = flatFontList[selectedIndex];

            if (!listEntry.isHeader) {
                expandedSelection.push(sourceFontListBox.items[selectedIndex]);
                continue;
            }
            for (var j = 0; j < flatFontList.length; j++) {
                if (!flatFontList[j].isHeader && flatFontList[j].family === listEntry.family) {
                    expandedSelection.push(sourceFontListBox.items[j]);
                }
            }
        }

        isUpdatingSelection = true;
        sourceFontListBox.selection = expandedSelection;
        isUpdatingSelection = false;
    }

    /**
     * 置換先リストで見出し行が選ばれたら選択を解除する
     * @returns {void}
     */
    function handleTargetFontSelection() {
        if (isUpdatingSelection || !targetFontListBox.selection) return;
        if (flatFontList[targetFontListBox.selection.index].isHeader) {
            targetFontListBox.selection = null;
        }
    }

    /**
     * ［PostScript名で表示］［段落・文字スタイルも更新］：表示を切り替えてリストを作り直す
     * @returns {void}
     */
    function handleDisplayModeChange() {
        rebuildFontList();
    }

    /**
     * ［フォントを置換］：選んだ置換元フォントを置換先フォントに置き換える
     * @returns {void}
     */
    function handleReplaceClick() {
        var sourceFontNames = getSelectedSourceFontNames();
        if (sourceFontNames.length === 0) {
            alert(getLabel(LABELS.alert.noSourceFont));
            return;
        }
        if (!hasTargetFontSelection()) {
            alert(getLabel(LABELS.alert.noTargetFont));
            return;
        }
        replaceFonts(sourceFontNames, getSelectedTargetFont());
    }

    /**
     * ［全置換］：使用中のすべてのフォントを1つのフォントにそろえる
     * @returns {void}
     */
    function handleReplaceAllClick() {
        /* 置換先が選ばれていれば、それにすべてをそろえる / Unify on the target font when one is selected */
        var targetFont = getSelectedTargetFont();
        if (targetFont) {
            replaceFonts(collectAllFontNames(), targetFont);
            return;
        }

        /* 置換先がなければ、置換元の1つ目にそろえる / Otherwise unify on the first source font */
        var sourceFontNames = getSelectedSourceFontNames();
        if (sourceFontNames.length === 0) {
            alert(getLabel(LABELS.alert.selectFonts));
            return;
        }

        var fallbackFont = findFontByName(sourceFontNames[0]);
        if (!fallbackFont) {
            alert(getLabel(LABELS.alert.targetNotFound, [sourceFontNames[0].replace("\t", " ")]));
            return;
        }
        replaceFonts(collectAllFontNames(sourceFontNames[0]), fallbackFont);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 見出し付きのフォントリストを1カラム分追加する
     * @param {Group} parent - 追加先のグループ
     * @param {object} labelSet - 見出しのラベル（ja/en）
     * @param {object} tooltipSet - ツールチップのラベル（ja/en）
     * @param {boolean} allowsMultiple - 複数選択を許可するか
     * @returns {ListBox} 追加したリストボックス
     */
    function addFontListColumn(parent, labelSet, tooltipSet, allowsMultiple) {
        var columnGroup = parent.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];
        columnGroup.spacing = LIST_LABEL_SPACING;
        columnGroup.add("statictext", undefined, labelText(labelSet));

        var fontListBox = columnGroup.add("listbox", undefined, [], { multiselect: allowsMultiple });
        fontListBox.preferredSize.height = LISTBOX_HEIGHT;
        fontListBox.tabEnabled = true;
        fontListBox.helpTip = getLabel(tooltipSet);
        return fontListBox;
    }

    /**
     * ダイアログを組み立てる
     * @returns {void}
     */
    function buildDialog() {
        mainDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(mainDialog);

        var listGroup = mainDialog.add("group");
        listGroup.orientation = "row";
        listGroup.alignChildren = ["fill", "top"];
        listGroup.spacing = COLUMN_SPACING;

        sourceFontListBox = addFontListColumn(listGroup, LABELS.fieldLabel.sourceFonts, LABELS.tooltip.sourceFonts, true);
        targetFontListBox = addFontListColumn(listGroup, LABELS.fieldLabel.targetFont, LABELS.tooltip.targetFont, false);

        /* 表示オプション / Display options */
        var optionGroup = mainDialog.add("group");
        optionGroup.orientation = "column";
        optionGroup.alignChildren = ["left", "center"];
        optionGroup.spacing = OPTION_SPACING;

        postScriptNameCheckbox = optionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.postScriptName));
        postScriptNameCheckbox.value = SHOW_POSTSCRIPT_NAME_DEFAULT;
        postScriptNameCheckbox.helpTip = getLabel(LABELS.tooltip.postScriptName);

        updateStylesCheckbox = optionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.updateStyles));
        updateStylesCheckbox.value = UPDATE_STYLES_DEFAULT;
        updateStylesCheckbox.helpTip = getLabel(LABELS.tooltip.updateStyles);

        /* ボタン行（閉じるは左、置換系は右）/ Button row: Close on the left, replace buttons on the right */
        var buttonRow = addButtonRow(mainDialog);
        var btnClose = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.close), { name: "cancel" });
        var btnReplaceAll = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.replaceAll));
        var btnReplace = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.replace), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        btnReplaceAll.helpTip = getLabel(LABELS.tooltip.replaceAll);
        btnReplace.helpTip = getLabel(LABELS.tooltip.replace);

        sourceFontListBox.onChange = handleSourceFontSelection;
        targetFontListBox.onChange = handleTargetFontSelection;
        postScriptNameCheckbox.onClick = handleDisplayModeChange;
        updateStylesCheckbox.onClick = handleDisplayModeChange;
        btnReplaceAll.onClick = handleReplaceAllClick;
        btnReplace.onClick = handleReplaceClick;
    }

    /**
     * 2つのリストにフォント一覧を流し込み、幅をそろえる
     * @returns {void}
     */
    function populateFontListBoxes() {
        var listBoxWidth = calculateListBoxWidth(flatFontList);

        isUpdatingSelection = true;
        sourceFontListBox.removeAll();
        targetFontListBox.removeAll();
        for (var i = 0; i < flatFontList.length; i++) {
            sourceFontListBox.add("item", flatFontList[i].label);
            targetFontListBox.add("item", flatFontList[i].label);
        }
        isUpdatingSelection = false;

        sourceFontListBox.preferredSize.width = listBoxWidth;
        targetFontListBox.preferredSize.width = listBoxWidth;
    }

    /**
     * フォント名を手がかりに、作り直したリストの選択を元に戻す
     * @param {Array<string>} sourceFontNames - 置換元として選択されていたフォント名
     * @param {string} targetFontName - 置換先として選択されていたフォント名
     * @returns {void}
     */
    function restoreSelection(sourceFontNames, targetFontName) {
        var restoredSelection = [];
        var restoredTargetIndex = -1;

        for (var i = 0; i < flatFontList.length; i++) {
            if (flatFontList[i].isHeader) continue;

            for (var j = 0; j < sourceFontNames.length; j++) {
                if (flatFontList[i].name === sourceFontNames[j]) {
                    restoredSelection.push(sourceFontListBox.items[i]);
                    break;
                }
            }
            if (restoredTargetIndex === -1 && flatFontList[i].name === targetFontName) {
                restoredTargetIndex = i;
            }
        }

        isUpdatingSelection = true;
        sourceFontListBox.selection = restoredSelection;
        targetFontListBox.selection = (restoredTargetIndex === -1) ? null : restoredTargetIndex;
        isUpdatingSelection = false;
    }

    /**
     * いちばん長いラベルからリストの幅を見積もる
     * @param {Array<object>} fontList - リストに並べるフォント
     * @returns {number} リストの幅（px）
     */
    function calculateListBoxWidth(fontList) {
        var maxLength = 0;
        for (var i = 0; i < fontList.length; i++) {
            if (fontList[i].isHeader) continue;
            if (fontList[i].label.length > maxLength) maxLength = fontList[i].label.length;
        }
        var estimatedWidth = maxLength * LISTBOX_CHAR_WIDTH + LISTBOX_WIDTH_PADDING;
        return Math.min(LISTBOX_WIDTH_MAX, Math.max(LISTBOX_WIDTH_MIN, estimatedWidth));
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * スクリプトのエントリーポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        doc = app.activeDocument;

        usedFontMap = collectUsedFonts();
        flatFontList = buildFlatFontList(usedFontMap);
        if (flatFontList.length === 0) {
            alert(getLabel(LABELS.alert.noFontsFound));
            return;
        }

        buildDialog();
        populateFontListBoxes();

        /* 先頭のフォントを選んでおく / Preselect the first font */
        sourceFontListBox.selection = 0;
        targetFontListBox.selection = 0;
        handleSourceFontSelection();

        mainDialog.show();
    }

    main();

})();
