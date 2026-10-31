#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdSplitHeadingStylesByPart
Name=部ごとに見出しの派生スタイルを当てて、目次の書式を分ける
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.0.0
Release-Date=2026-10-30
Description-BEGIN
目次に載せる見出しを、部（部の扉の段落スタイル）ごとに派生の段落スタイル（h1.blue など）へ振り分け、目次スタイルにも派生スタイルごとの書式（toc-h1.blue など）を設定します。
以後は［目次の更新］だけで、部ごとに異なる書式の目次ができます。

Assigns the headings listed in the table of contents to per-part variant paragraph styles (such as h1.blue), and gives each variant its own TOC format (such as toc-h1.blue) in the TOC style.
After that, a plain Update Table of Contents formats each part differently.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

目次に載せる見出しを、部（部の扉の段落スタイル）ごとに派生の段落スタイル（h1.blue など）へ振り分け、目次スタイルにも派生スタイルごとの書式（toc-h1.blue など）を設定します。
以後は［目次の更新］だけで、部ごとに異なる書式の目次ができます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSplitHeadingStylesByPart.md

### Overview

Assigns the headings listed in the table of contents to per-part variant paragraph styles (such as h1.blue), and gives each variant its own TOC format (such as toc-h1.blue) in the TOC style.
After that, a plain Update Table of Contents formats each part differently.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSplitHeadingStylesByPart.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdSplitHeadingStylesByPart";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-30";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "部ごとに見出しの派生スタイルを当てて、目次の書式を分ける"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdSplitHeadingStylesByPart.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdSplitHeadingStylesByPart.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/style/IdSplitHeadingStylesByPart.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 派生スタイル名の区切り（「h1.blue」の「.」） / Separator between a style name and its variant suffix ("." in "h1.blue") */
    var VARIANT_SEPARATOR = ".";

    /* 部の見出しとして表示する文字数 / Characters of each part heading shown in the dialog */
    var PART_TITLE_MAX_LENGTH = 24;

    /* InDesign がマーカー位置に入れる文字。ExtendScript では "﻿" と書くと空文字列になる
       / The character InDesign puts at marker positions; "﻿" becomes an empty string in ExtendScript */
    var MARKER_CHAR = String.fromCharCode(0xFEFF);

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
        dialog: {
            title: { ja: "部ごとに見出しスタイルを分ける", en: "Split Heading Styles by Part" }
        },
        panel: {
            toc: { ja: "目次", en: "Table of Contents" },
            parts: { ja: "部ごとの派生スタイル", en: "Variant per Part" },
            headings: { ja: "分ける見出し", en: "Headings to Split" },
            options: { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            tocStyle: { ja: "目次スタイル", en: "TOC style" },
            partStyle: { ja: "部の段落スタイル", en: "Part style" }
        },
        dropdown: {
            none: { ja: "なし", en: "None" }
        },
        checkbox: {
            updateToc: { ja: "目次を更新", en: "Update the table of contents" },
            removeBreaks: { ja: "目次の強制改行を削除", en: "Remove forced line breaks in the TOC" }
        },
        format: {
            partRow: { ja: "p.%1 %2", en: "p. %1 %2" },
            heading: { ja: "%1（レベル %2）", en: "%1 (level %2)" }
        },
        tooltip: {
            partStyle: { ja: "この段落から、次の同じスタイルの段落の手前までを1つの部とみなします", en: "Each paragraph in this style starts a part that runs up to the next one" },
            parts: { ja: "初期値は、部の中ですでに使われている派生スタイルの多数決です", en: "The default is the variant used most often within the part" },
            updateToc: { ja: "既存の目次を、選んだ目次スタイルで作り直します", en: "Rebuilds the existing table of contents with the selected TOC style" },
            removeBreaks: { ja: "目次スタイルの［強制改行を削除］がオンでも残る強制改行を削除します", en: "Removes the forced line breaks that remain even when the TOC style's Remove Forced Line Break is on" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        message: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noTocStyle: { ja: "目次スタイルがありません。", en: "The document has no TOC styles." },
            noEntries: { ja: "目次スタイル「%1」に段落スタイルが登録されていません。", en: "The TOC style \"%1\" includes no paragraph styles." },
            noVariants: {
                ja: "目次の書式の派生スタイルが見つかりません。\n「%1」のように、目次の書式の段落スタイル名に「%2接尾辞」を付けた段落スタイルを作ってから実行してください。",
                en: "No variant TOC format styles were found.\nCreate paragraph styles named like \"%1\" (the TOC format style name plus \"%2suffix\") and run the script again."
            },
            noParts: { ja: "本文に段落スタイル「%1」の段落がありません。", en: "No paragraphs in the body use the paragraph style \"%1\"." },
            nothingToDo: { ja: "変更する見出し・目次スタイルはありません。", en: "There is nothing to change." },
            done: { ja: "見出し %1 段落のスタイルを変更しました。", en: "Changed the style of %1 heading paragraphs." },
            createdStyles: { ja: "作成した段落スタイル：%1", en: "Paragraph styles created: %1" },
            entries: { ja: "目次スタイルの項目：追加 %1・書式の変更 %2", en: "TOC style entries: %1 added, %2 reformatted" },
            tocUpdated: { ja: "目次を更新しました（強制改行の削除 %1）。", en: "Updated the table of contents (%1 forced line breaks removed)." },
            tocNotFound: { ja: "目次が見つからないため、目次は更新していません。", en: "No table of contents was found, so it was not updated." },
            tocFailed: { ja: "目次を更新できませんでした：%1", en: "Could not update the table of contents: %1" }
        },
        undo: {
            split: { ja: "部ごとに見出しスタイルを分ける", en: "Split Heading Styles by Part" }
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

    var FIELD_LABEL_WIDTH = 110;   /* 項目名の幅 / width of field labels */
    var FIELD_INPUT_WIDTH = 200;   /* ドロップダウンの幅 / width of field dropdowns */
    var PART_LABEL_WIDTH  = 250;   /* 部の見出しの幅 / width of part headings */
    var VARIANT_WIDTH     = 120;   /* 派生スタイルのドロップダウンの幅 / width of variant dropdowns */

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

    // =========================================
    // 文字列 / Text
    // =========================================

    /**
     * 段落の文字列から、マーカー位置の文字と末尾の改行を除き、改行を空白にする（表示用）
     * @param {string} text - 段落の contents
     * @returns {string} 表示用の文字列
     */
    function toDisplayText(text) {
        return text.split(MARKER_CHAR).join("").replace(/\r$/, "").replace(/[\r\n]/g, " ");
    }

    /**
     * 見出しを比べるための文字列にする（マーカー位置の文字・空白・改行を除く）
     * @param {string} text - 段落の contents
     * @returns {string} 比較用の文字列
     */
    function toMatchKey(text) {
        return text.split(MARKER_CHAR).join("").replace(/[\s　]/g, "");
    }

    /**
     * 長い文字列を切り詰める
     * @param {string} text - 文字列
     * @param {number} maxLength - 最大の文字数
     * @returns {string} 切り詰めた文字列
     */
    function truncateText(text, maxLength) {
        return (text.length > maxLength) ? text.substr(0, maxLength) + "..." : text;
    }

    // =========================================
    // ドキュメントの解析 / Document analysis
    // =========================================

    /**
     * everyItem() のプロパティを必ず配列で受け取る（要素が1つだと配列にならない）
     * @param {*} value - everyItem().xxx の値
     * @returns {Array} 配列
     */
    function toArray(value) {
        return (value instanceof Array) ? value : [value];
    }

    /**
     * 段落スタイルを名前で引ける表にする（グループ内も含む。同名は先のものを使う）
     * @param {Document} doc - ドキュメント
     * @returns {Object} 名前 → ParagraphStyle
     */
    function buildStyleMap(doc) {
        var styleMap = {};
        var allStyles = doc.allParagraphStyles;
        for (var i = 0; i < allStyles.length; i++) {
            if (!styleMap.hasOwnProperty(allStyles[i].name)) styleMap[allStyles[i].name] = allStyles[i];
        }
        return styleMap;
    }

    /**
     * 目次エントリーの書式の段落スタイル名。［同一スタイル］などで段落スタイルでなければ空文字
     * @param {TOCStyleEntry} entry - 目次スタイルの項目
     * @returns {string} 段落スタイル名
     */
    function getEntryFormatName(entry) {
        try {
            var formatStyle = entry.formatStyle;
            if (formatStyle && typeof formatStyle === "object" && formatStyle.constructor.name === "ParagraphStyle") return formatStyle.name;
        } catch (e) {}
        return "";
    }

    /**
     * 段落の載っているページを返す。ページ上にない（親ページ・ペーストボード・オーバーセット）なら null
     * @param {Paragraph} para - 段落
     * @returns {Page|null} ページ
     */
    function getParagraphPage(para) {
        try {
            var frames = para.parentTextFrames;
            if (frames.length === 0 || !frames[0].parentPage) return null;
            var page = frames[0].parentPage;
            return (page.parent.constructor.name === "Spread") ? page : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * 目次スタイルと部の段落スタイルから、見出し・部・派生スタイルの候補を調べる
     * @param {Document} doc - ドキュメント
     * @param {TOCStyle} tocStyle - 目次スタイル
     * @param {string|null} partBaseName - 部の段落スタイル名（null なら最上位レベルの項目）
     * @returns {Object} { styleMap, baseEntries, partBaseName, candidates, tocStories, records, parts }
     */
    function analyzeDocument(doc, tocStyle, partBaseName) {
        var analysis = { styleMap: buildStyleMap(doc), baseEntries: [], partBaseName: "", candidates: [], tocStories: [], records: [], parts: [] };
        var styleMap = analysis.styleMap;

        /* 項目のうち、ほかの項目の派生（「h1」に対する「h1.blue」）ではないものを元の見出しとする
           / Entries that are not variants of another entry are the base headings */
        var entries = tocStyle.tocStyleEntries.everyItem().getElements();
        var entryNameSet = {};
        for (var i = 0; i < entries.length; i++) entryNameSet[entries[i].name] = true;
        for (i = 0; i < entries.length; i++) {
            var separatorIndex = entries[i].name.lastIndexOf(VARIANT_SEPARATOR);
            if (separatorIndex > 0 && entryNameSet[entries[i].name.substr(0, separatorIndex)]) continue;
            analysis.baseEntries.push({ name: entries[i].name, entry: entries[i], level: entries[i].level, formatName: getEntryFormatName(entries[i]) });
        }
        if (analysis.baseEntries.length === 0) return analysis;

        var baseSet = {};
        var topEntry = analysis.baseEntries[0];
        for (i = 0; i < analysis.baseEntries.length; i++) {
            baseSet[analysis.baseEntries[i].name] = analysis.baseEntries[i];
            if (analysis.baseEntries[i].level < topEntry.level) topEntry = analysis.baseEntries[i];
        }
        analysis.partBaseName = (partBaseName && baseSet[partBaseName]) ? partBaseName : topEntry.name;

        /* 派生の候補は、目次の書式の段落スタイルに「.接尾辞」を付けたものから集める
           / Variant suffixes come from the TOC format styles ("toc-h1" → "toc-h1.blue") */
        var candidateSet = {};
        var tocFormatSet = {};
        for (i = 0; i < analysis.baseEntries.length; i++) {
            var formatName = analysis.baseEntries[i].formatName;
            if (!formatName) continue;
            tocFormatSet[formatName] = true;
            var prefix = formatName + VARIANT_SEPARATOR;
            for (var styleName in styleMap) {
                if (!styleMap.hasOwnProperty(styleName) || styleName.indexOf(prefix) !== 0 || styleName.length === prefix.length) continue;
                tocFormatSet[styleName] = true;
                var suffix = styleName.substr(prefix.length);
                if (!candidateSet[suffix]) {
                    candidateSet[suffix] = true;
                    analysis.candidates.push(suffix);
                }
            }
        }

        var stories = doc.stories.everyItem().getElements();
        for (var s = 0; s < stories.length; s++) {
            if (stories[s].paragraphs.length === 0) continue;
            var paraStyles = toArray(stories[s].paragraphs.everyItem().appliedParagraphStyle);

            /* 目次の書式を使った段落があるストーリーは目次とみなし、本文の見出しから外す
               / A story using a TOC format style is a table of contents, not body text */
            var isTocStory = false;
            for (var t = 0; t < paraStyles.length; t++) {
                if (tocFormatSet[paraStyles[t].name]) { isTocStory = true; break; }
            }
            if (isTocStory) {
                analysis.tocStories.push(stories[s]);
                continue;
            }

            for (var p = 0; p < paraStyles.length; p++) {
                var parsed = parseVariantName(paraStyles[p].name, baseSet, candidateSet);
                if (!parsed) continue;
                var para = stories[s].paragraphs[p];
                var page = getParagraphPage(para);
                if (!page) continue;
                analysis.records.push({ para: para, styleName: paraStyles[p].name, base: parsed.base, suffix: parsed.suffix, offset: page.documentOffset, pageName: page.name, part: null });
            }
        }

        buildParts(analysis);
        return analysis;
    }

    /**
     * 段落スタイル名を、元の見出しと派生の接尾辞に分ける
     * @param {string} styleName - 段落スタイル名
     * @param {Object} nameSet - 元の名前の集合
     * @param {Object} candidateSet - 接尾辞の候補の集合
     * @returns {{base: string, suffix: string}|null} 当てはまらなければ null
     */
    function parseVariantName(styleName, nameSet, candidateSet) {
        if (nameSet[styleName]) return { base: styleName, suffix: "" };
        var separatorIndex = styleName.lastIndexOf(VARIANT_SEPARATOR);
        if (separatorIndex <= 0) return null;
        var baseName = styleName.substr(0, separatorIndex);
        var suffix = styleName.substr(separatorIndex + VARIANT_SEPARATOR.length);
        return (nameSet[baseName] && candidateSet[suffix]) ? { base: baseName, suffix: suffix } : null;
    }

    /**
     * 部の段落をページ順に並べて部を作り、見出しを部に振り分け、部ごとの派生の初期値を多数決で決める
     * @param {Object} analysis - analyzeDocument の途中結果
     * @returns {void}
     */
    function buildParts(analysis) {
        var parts = [];
        for (var i = 0; i < analysis.records.length; i++) {
            var record = analysis.records[i];
            if (record.base !== analysis.partBaseName) continue;
            parts.push({ start: record.offset, pageName: record.pageName, title: toDisplayText(record.para.contents), matchKey: toMatchKey(record.para.contents), votes: {}, variant: "" });
        }
        /* 部は多くても数十なので比較関数つきの sort で足りる / Only a few dozen parts at most, so a comparator sort is fine */
        parts.sort(function (a, b) { return a.start - b.start; });
        for (i = 0; i < parts.length; i++) parts[i].index = i;

        for (i = 0; i < analysis.records.length; i++) {
            var rec = analysis.records[i];
            for (var k = parts.length - 1; k >= 0; k--) {
                if (parts[k].start <= rec.offset) { rec.part = parts[k]; break; }
            }
            if (rec.part && rec.suffix) rec.part.votes[rec.suffix] = (rec.part.votes[rec.suffix] || 0) + 1;
        }

        for (i = 0; i < parts.length; i++) {
            var bestCount = 0;
            for (var suffix in parts[i].votes) {
                if (parts[i].votes.hasOwnProperty(suffix) && parts[i].votes[suffix] > bestCount) {
                    parts[i].variant = suffix;
                    bestCount = parts[i].votes[suffix];
                }
            }
        }
        analysis.parts = parts;
    }

    // =========================================
    // スタイルと目次の変更 / Changing styles and the TOC
    // =========================================

    /**
     * 派生スタイルを返す。無ければ元のスタイルを基準にして同じ場所（グループ）に作る
     * @param {Object} styleMap - 名前 → ParagraphStyle
     * @param {string} baseName - 元の段落スタイル名
     * @param {string} variantName - 派生スタイル名
     * @param {string[]} createdNames - 作成したスタイル名を足す先
     * @returns {ParagraphStyle} 派生スタイル
     */
    function ensureVariantStyle(styleMap, baseName, variantName, createdNames) {
        if (styleMap.hasOwnProperty(variantName)) return styleMap[variantName];
        var baseStyle = styleMap[baseName];
        var variantStyle = baseStyle.parent.paragraphStyles.add({ name: variantName, basedOn: baseStyle });
        /* 次のスタイルが自分自身なら派生でも自分自身に / Keep "next style = itself" pointing at the new style */
        variantStyle.nextStyle = (baseStyle.nextStyle.id === baseStyle.id) ? variantStyle : baseStyle.nextStyle;
        styleMap[variantName] = variantStyle;
        createdNames.push(variantName);
        return variantStyle;
    }

    /**
     * 目次スタイルの項目を、元の項目の設定を写して追加する
     * @param {TOCStyle} tocStyle - 目次スタイル
     * @param {TOCStyleEntry} baseEntry - 元の項目
     * @param {string} entryName - 追加する項目（段落スタイル名）
     * @returns {TOCStyleEntry} 追加した項目
     */
    function addEntryLike(tocStyle, baseEntry, entryName) {
        var newEntry = tocStyle.tocStyleEntries.add({ name: entryName });
        var copyKeys = ["level", "pageNumberPosition", "pageNumberStyle", "separator", "separatorStyle", "sortAlphabet"];
        for (var i = 0; i < copyKeys.length; i++) {
            try { newEntry[copyKeys[i]] = baseEntry[copyKeys[i]]; } catch (e) {}
        }
        return newEntry;
    }

    /**
     * ストーリー内の強制改行を削除する（検索設定は元に戻す）
     * @param {Story} story - ストーリー
     * @returns {number} 削除した数
     */
    function removeForcedLineBreaks(story) {
        var savedFind = app.findGrepPreferences.properties;
        var savedChange = app.changeGrepPreferences.properties;
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.findWhat = "\\n";
        app.changeGrepPreferences.changeTo = "";
        var removedCount = story.changeGrep().length;
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences.properties = savedFind;
        app.changeGrepPreferences.properties = savedChange;
        return removedCount;
    }

    /**
     * 見出しを部ごとの派生スタイルへ振り分け、目次スタイルの項目を整え、必要なら目次を更新する
     * @param {Document} doc - ドキュメント
     * @param {TOCStyle} tocStyle - 目次スタイル
     * @param {Object} analysis - analyzeDocument の戻り値
     * @param {Object} options - { variants: string[], bases: Object, updateToc: boolean, removeBreaks: boolean }
     * @returns {Object} 結果
     */
    function applySplit(doc, tocStyle, analysis, options) {
        var result = { paragraphs: 0, createdNames: [], addedEntries: 0, changedEntries: 0, tocUpdated: false, tocError: "", removedBreaks: 0 };
        var styleMap = analysis.styleMap;

        for (var i = 0; i < analysis.records.length; i++) {
            var record = analysis.records[i];
            if (!record.part || !options.bases[record.base]) continue;
            var variant = options.variants[record.part.index];
            var targetName = variant ? record.base + VARIANT_SEPARATOR + variant : record.base;
            if (record.styleName === targetName) continue;
            var targetStyle = variant ? ensureVariantStyle(styleMap, record.base, targetName, result.createdNames) : styleMap[record.base];
            /* 手で付けた書式は残す / Keep local overrides */
            record.para.applyParagraphStyle(targetStyle, false);
            result.paragraphs++;
        }

        for (var b = 0; b < analysis.baseEntries.length; b++) {
            var baseEntry = analysis.baseEntries[b];
            if (!options.bases[baseEntry.name] || !baseEntry.formatName) continue;
            for (var c = 0; c < analysis.candidates.length; c++) {
                var variantName = baseEntry.name + VARIANT_SEPARATOR + analysis.candidates[c];
                if (!styleMap.hasOwnProperty(variantName)) continue;
                var formatStyle = ensureVariantStyle(styleMap, baseEntry.formatName, baseEntry.formatName + VARIANT_SEPARATOR + analysis.candidates[c], result.createdNames);
                var entry = tocStyle.tocStyleEntries.itemByName(variantName);
                if (!entry.isValid) {
                    entry = addEntryLike(tocStyle, baseEntry.entry, variantName);
                    result.addedEntries++;
                }
                if (getEntryFormatName(entry) !== formatStyle.name) {
                    entry.formatStyle = formatStyle;
                    result.changedEntries++;
                }
            }
        }

        if (options.updateToc && analysis.tocStories.length > 0) {
            try {
                var tocStories = doc.createTOC(tocStyle, true);
                result.tocUpdated = true;
                if (options.removeBreaks) {
                    for (var s = 0; s < tocStories.length; s++) result.removedBreaks += removeForcedLineBreaks(tocStories[s]);
                }
            } catch (e) {
                result.tocError = e.message;
            }
        }
        return result;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名つきの行を作る
     * @param {Panel|Group} parent - 行を足す先
     * @param {string} labelKey - 項目名のラベルのパス
     * @returns {Group} 行
     */
    function addFieldRow(parent, labelKey) {
        var rowGroup = parent.add("group");
        setupRow(rowGroup);
        var fieldLabel = rowGroup.add("statictext", undefined, labelText(labelKey));
        fieldLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        fieldLabel.justify = "right";
        return rowGroup;
    }

    /**
     * 設定ダイアログボックスを表示する
     * @param {Document} doc - ドキュメント
     * @param {Object} analysis - analyzeDocument の戻り値
     * @param {Object} dialogState - { tocStyleIndex, partBaseName }。選び直したときはここを書き換える
     * @param {boolean} canRun - false なら［OK］を押せなくする（目次スタイルの選び直しだけできる）
     * @returns {Object|null} options。キャンセルなら null、目次スタイルか部の段落スタイルを選び直したら { rebuild: true }
     */
    function showOptionsDialog(doc, analysis, dialogState, canRun) {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        var tocPanel = dialog.add("panel", undefined, getLabel("panel.toc"));
        setupPanel(tocPanel, 6);
        var ddTocStyle = addFieldRow(tocPanel, "fieldLabel.tocStyle").add("dropdownlist", undefined, toArray(doc.tocStyles.everyItem().name));
        ddTocStyle.preferredSize.width = FIELD_INPUT_WIDTH;
        ddTocStyle.selection = dialogState.tocStyleIndex;

        var partStyleRow = addFieldRow(tocPanel, "fieldLabel.partStyle");
        var baseNames = [];
        var partStyleIndex = 0;
        for (var i = 0; i < analysis.baseEntries.length; i++) {
            baseNames.push(analysis.baseEntries[i].name);
            if (analysis.baseEntries[i].name === analysis.partBaseName) partStyleIndex = i;
        }
        var ddPartStyle = partStyleRow.add("dropdownlist", undefined, baseNames);
        ddPartStyle.preferredSize.width = FIELD_INPUT_WIDTH;
        if (baseNames.length > 0) ddPartStyle.selection = partStyleIndex;
        ddPartStyle.helpTip = getLabel("tooltip.partStyle");

        var partsPanel = dialog.add("panel", undefined, getLabel("panel.parts"));
        setupPanel(partsPanel, 6);
        partsPanel.helpTip = getLabel("tooltip.parts");
        var variantItems = [getLabel("dropdown.none")].concat(analysis.candidates);
        var variantDropdowns = [];
        for (i = 0; i < analysis.parts.length; i++) {
            var part = analysis.parts[i];
            var partRow = partsPanel.add("group");
            setupRow(partRow);
            var partLabel = partRow.add("statictext", undefined, getLabel("format.partRow", [part.pageName, truncateText(part.title, PART_TITLE_MAX_LENGTH)]));
            partLabel.preferredSize.width = PART_LABEL_WIDTH;
            var ddVariant = partRow.add("dropdownlist", undefined, variantItems);
            ddVariant.preferredSize.width = VARIANT_WIDTH;
            var variantIndex = 0;
            for (var c = 0; c < analysis.candidates.length; c++) {
                if (analysis.candidates[c] === part.variant) variantIndex = c + 1;
            }
            ddVariant.selection = variantIndex;
            variantDropdowns.push(ddVariant);
        }

        var headingsPanel = dialog.add("panel", undefined, getLabel("panel.headings"));
        setupPanel(headingsPanel, 6);
        var headingCheckboxes = {};
        for (i = 0; i < analysis.baseEntries.length; i++) {
            var baseEntry = analysis.baseEntries[i];
            var chkHeading = headingsPanel.add("checkbox", undefined, getLabel("format.heading", [baseEntry.name, baseEntry.level]));
            chkHeading.value = true;
            headingCheckboxes[baseEntry.name] = chkHeading;
        }

        var optionsPanel = dialog.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel, 6);
        var chkUpdateToc = optionsPanel.add("checkbox", undefined, getLabel("checkbox.updateToc"));
        chkUpdateToc.helpTip = getLabel("tooltip.updateToc");
        chkUpdateToc.value = analysis.tocStories.length > 0;
        chkUpdateToc.enabled = analysis.tocStories.length > 0;
        var chkRemoveBreaks = optionsPanel.add("checkbox", undefined, getLabel("checkbox.removeBreaks"));
        chkRemoveBreaks.helpTip = getLabel("tooltip.removeBreaks");
        chkRemoveBreaks.value = true;
        chkRemoveBreaks.enabled = chkUpdateToc.value;
        chkUpdateToc.onClick = function () { chkRemoveBreaks.enabled = chkUpdateToc.value; };

        var buttonRow = addButtonRow(dialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        btnOK.enabled = canRun;
        alignRightOnlyButtonRow(buttonRow);

        /* 目次スタイル・部の段落スタイルを選び直したら、調べ直して開き直す / Re-analyze and reopen after another TOC style or part style is chosen */
        var REBUILD_CODE = 3;
        ddTocStyle.onChange = function () {
            dialogState.tocStyleIndex = ddTocStyle.selection.index;
            dialogState.partBaseName = null;
            dialog.close(REBUILD_CODE);
        };
        ddPartStyle.onChange = function () {
            dialogState.partBaseName = ddPartStyle.selection.text;
            dialog.close(REBUILD_CODE);
        };

        var dialogResult = dialog.show();
        if (dialogResult === REBUILD_CODE) return { rebuild: true };
        if (dialogResult !== 1) return null;

        var options = { variants: [], bases: {}, updateToc: chkUpdateToc.value, removeBreaks: chkRemoveBreaks.value };
        for (i = 0; i < variantDropdowns.length; i++) {
            options.variants.push(variantDropdowns[i].selection.index === 0 ? "" : analysis.candidates[variantDropdowns[i].selection.index - 1]);
        }
        for (var baseName in headingCheckboxes) {
            if (headingCheckboxes.hasOwnProperty(baseName)) options.bases[baseName] = headingCheckboxes[baseName].value;
        }
        return options;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 解析結果に問題があればメッセージを返す
     * @param {TOCStyle} tocStyle - 目次スタイル
     * @param {Object} analysis - analyzeDocument の戻り値
     * @returns {string} メッセージ。問題なければ空文字
     */
    function getAnalysisProblem(tocStyle, analysis) {
        if (analysis.baseEntries.length === 0) return getLabel("message.noEntries", [tocStyle.name]);
        if (analysis.candidates.length === 0) {
            var sampleFormat = analysis.baseEntries[0].formatName || "toc-h1";
            return getLabel("message.noVariants", [sampleFormat + VARIANT_SEPARATOR + "blue", VARIANT_SEPARATOR]);
        }
        if (analysis.parts.length === 0) return getLabel("message.noParts", [analysis.partBaseName]);
        return "";
    }

    /**
     * 結果のメッセージを組み立てる
     * @param {Object} result - applySplit の戻り値
     * @param {boolean} updateRequested - 目次の更新を選んだか
     * @returns {string} メッセージ
     */
    function buildResultMessage(result, updateRequested) {
        var lines = [getLabel("message.done", [result.paragraphs])];
        if (result.createdNames.length > 0) lines.push(getLabel("message.createdStyles", [result.createdNames.join(", ")]));
        lines.push(getLabel("message.entries", [result.addedEntries, result.changedEntries]));
        if (result.tocUpdated) lines.push(getLabel("message.tocUpdated", [result.removedBreaks]));
        else if (result.tocError) lines.push(getLabel("message.tocFailed", [result.tocError]));
        else if (updateRequested) lines.push(getLabel("message.tocNotFound"));
        return lines.join("\n");
    }

    /**
     * エントリポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("message.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        if (doc.tocStyles.length === 0) {
            alert(getLabel("message.noTocStyle"));
            return;
        }

        var dialogState = { tocStyleIndex: 0, partBaseName: null };
        var tocStyle, analysis, options;
        while (true) {
            tocStyle = doc.tocStyles[dialogState.tocStyleIndex];
            analysis = analyzeDocument(doc, tocStyle, dialogState.partBaseName);
            var problem = getAnalysisProblem(tocStyle, analysis);
            if (problem) {
                alert(problem);
                /* 目次スタイルがほかにあれば、選び直せるようにダイアログボックスは開く / With other TOC styles available, still open the dialog so another can be chosen */
                if (doc.tocStyles.length === 1) return;
            }
            options = showOptionsDialog(doc, analysis, dialogState, !problem);
            if (!options) return;
            if (!options.rebuild) break;
        }

        var result;
        /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
        app.doScript(function () { result = applySplit(doc, tocStyle, analysis, options); }, ScriptLanguage.JAVASCRIPT, undefined,
            UndoModes.ENTIRE_SCRIPT, getLabel("undo.split") + " " + SCRIPT_VERSION);

        if (result.paragraphs === 0 && result.createdNames.length === 0 && result.addedEntries === 0 && result.changedEntries === 0 && !result.tocUpdated) {
            alert(getLabel("message.nothingToDo"));
            return;
        }
        alert(buildResultMessage(result, options.updateToc));
    }

    main();

})();
