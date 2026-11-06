#target indesign
#targetengine "IdCountKeywordPalette"

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdCountKeywordPalette
Name=キーワードの件数を数えて表記の揺れを統一
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.4.0
Release-Date=2026-11-05
Description-BEGIN
最大5つのキーワードが、選択中のストーリーまたはドキュメント全体にそれぞれいくつあるかを常駐パレットで数え、該当箇所を一覧しながら表記の揺れを統一します。

A floating palette that counts how many times each of up to five keywords appears in the selected stories or in the whole document, lists where they are, and unifies spelling variants.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

最大5つのキーワードが、選択中のストーリーまたはドキュメント全体にそれぞれいくつあるかを常駐パレットで数え、該当箇所を一覧しながら表記の揺れを統一します。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCountKeywordPalette.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n9f6d4d2bbeed

### Overview

A floating palette that counts how many times each of up to five keywords appears in the selected stories or in the whole document, lists where they are, and unifies spelling variants.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCountKeywordPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdCountKeywordPalette";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-07";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "キーワードの件数を数えて表記の揺れを統一"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdCountKeywordPalette.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdCountKeywordPalette.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n9f6d4d2bbeed"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdCountKeywordPalette.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* キーワード欄の数 / Number of keyword fields */
    var KEYWORD_FIELD_COUNT = 5;

    /* 一覧に出す前後の文字数の初期値と上限 / Default and maximum number of characters shown around each match */
    var DEFAULT_CONTEXT_LENGTH = 3;
    var MAX_CONTEXT_LENGTH = 20;

    /* 一覧で選んだ文字を表示するときの表示倍率（％）の初期値と、スライダーの範囲
       / Default zoom (%) used to show the picked text, and the slider range */
    var DEFAULT_ZOOM_PERCENT = 200;
    var MIN_ZOOM_PERCENT = 50;
    var MAX_ZOOM_PERCENT = 800;
    var ZOOM_STEP_PERCENT = 10; /* スライダーの値をこの単位に丸める / the slider value snaps to this step */

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
            title: { ja: "表記統一ヘルパー", en: "Spelling Consistency Helper" }
        },
        fieldLabel: {
            total: { ja: "合計", en: "Total" },
            contextLength: { ja: "前後の文字数", en: "Characters Around" },
            zoom: { ja: "表示倍率", en: "Zoom" },
            changeTo: { ja: "置換文字列", en: "Change to" }
        },
        panel: {
            variants: { ja: "揺れ候補", en: "Variants" },
            keywords: { ja: "キーワード（○＝統一先）", en: "Keywords (○ = unify target)" },
            scope: { ja: "対象", en: "Scope" },
            options: { ja: "オプション", en: "Options" },
            matches: { ja: "該当箇所", en: "Matches" }
        },
        column: {
            page: { ja: "ページ", en: "Page" },
            context: { ja: "前後の文字", en: "Context" }
        },
        radio: {
            story: { ja: "ストーリー", en: "Story" },
            document: { ja: "ドキュメント", en: "Document" }
        },
        checkbox: {
            caseSensitive: { ja: "大文字と小文字を区別", en: "Case Sensitive" },
            wholeWord: { ja: "単語単位", en: "Whole Word" },
            masterPages: { ja: "親ページを含む", en: "Include Parent Pages" },
            excludeToc: { ja: "目次を除く", en: "Exclude Tables of Contents" },
            excludeContained: { ja: "長いキーワードの一部は除く", en: "Skip parts of longer keywords" },
            bracketMatch: { ja: "該当文字を【】で囲む", en: "Wrap matches in 【】" },
            withinParagraph: { ja: "段落内に限る", en: "Within Paragraph" }
        },
        variantKind: {
            all: { ja: "すべて", en: "All" },
            katakana: { ja: "カタカナ", en: "Katakana" },
            kanji: { ja: "漢字", en: "Kanji" },
            latin: { ja: "英数字", en: "Latin" },
            suru: { ja: "する／をする", en: "する／をする" },
            dictionary: { ja: "辞書", en: "Dictionary" },
            review: { ja: "要確認", en: "To Review" }
        },
        variantSort: {
            frequent: { ja: "多い順", en: "Most first" },
            rare: { ja: "少ない順", en: "Fewest first" }
        },
        variantTable: {
            group: { ja: "組", en: "Group" },
            spelling: { ja: "表記", en: "Spelling" },
            count: { ja: "件数", en: "Count" },
            target: { ja: "統一先", en: "Target" }
        },
        tooltip: {
            story: { ja: "［検索］を押した時点で選択している文字やテキストフレームを含むストーリー（表・脚注を含む）で検索します。", en: "Searches the stories that contain the text or text frames selected when you click Find, including tables and footnotes." },
            document: { ja: "作業中のドキュメント内のすべてのテキストを検索します。", en: "Searches all text in the active document." },
            masterPages: { ja: "ドキュメントを対象にしたとき、親ページのテキストも検索します。", en: "Also searches text on parent pages when the scope is the document." },
            excludeToc: { ja: "［目次の作成］で作った目次のストーリーは検索しません。", en: "Does not search table-of-contents stories made with Table of Contents." },
            excludeContained: {
                ja: "「取引」と「取引き」のように、長いキーワードの一部として見つかった分を短いキーワードの件数から除きます。",
                en: "When a keyword is found as part of a longer keyword (e.g. \"deal\" inside \"dealer\"), it is not counted for the shorter one."
            },
            loadSelection: { ja: "選択している文字を、最後に入力したキーワード欄（空いていれば最初の空欄）に入れます。", en: "Puts the selected text into the keyword field you last typed in, or the first empty one." },
            keyword: { ja: "return キーでも検索します。", en: "Press Return to find." },
            clearKeyword: { ja: "キーワード欄を空にします。option キーを押しながらクリックすると、すべての欄を空にします。", en: "Clears the keyword field. Option-click clears all the fields." },
            wholeWord: { ja: "前後が同じ文字種（英数字・漢字・カタカナ）に続くものを除きます。「仕事」の中の「事」、pineapple の中の apple などです。", en: "Skips matches that run on into the same script (letters, kanji, katakana), like apple in pineapple." },
            withinParagraph: { ja: "一覧の前後の文字を段落の中だけにします。段落の区切り（¶）と、その先の文字は表示しません。", en: "Keeps the characters around a match inside its paragraph. The paragraph mark (¶) and what lies beyond it are not shown." },
            collectVariants: { ja: "［対象］の文字を読み取り、同じ語の表記違いと、辞書にある揺れを一覧にします。キーワードは入れなくてかまいません。", en: "Reads the text of the scope and lists words written in more than one way, including the variants in the dictionary. No keywords are needed." },
            variantList: { ja: "組を選ぶとキーワード欄に入り、検索します。◉ の表記が統一先になります（前回の統一先、辞書の推奨、いちばん多い表記の順。要確認の組は統一先を選びません）。", en: "Pick a group to put it in the keyword fields and find it. The spelling marked ◉ becomes the unify target (the last one used, the dictionary's, then the most used; to-review groups have none)." },
            variantFilter: { ja: "一覧に出す組の種類を絞ります。", en: "Limits the list to one kind of group." },
            variantSort: { ja: "組の並び順を、出現数の多い順か少ない順にします。", en: "Orders the groups by how often they appear." },
            ignore: { ja: "選んだ組を出さないようにします（InDesign を終了するまで）。", en: "Stops showing the selected group until InDesign quits." },
            copyTable: { ja: "表示している組を、組・表記・件数・統一先の表としてコピーします。表計算アプリに貼り付けられます。", en: "Copies the groups shown as a table of group, spelling, count and target, ready to paste into a spreadsheet." },
            find: { ja: "キーワードを検索し、件数と該当箇所を表示します（キーワード欄の return キーでも検索）。", en: "Finds the keywords and shows the counts and matches (Return in a keyword field also finds)." },
            contextLength: { ja: "見つかった文字の前後に表示する文字数", en: "Number of characters shown before and after each match" },
            matchList: {
                ja: "クリックでその文字を選択して表示し、［置換文字列］に入れます。ダブルクリックで除外する／しないを切り替えます（チェックなし＝除外）。",
                en: "Click to select and show the text and put it in Change to. Double-click to exclude or include the row (unchecked = excluded)."
            },
            exclude: {
                ja: "一覧で選んだ行を除外します。除外した行はチェックが外れ、件数と［すべてを統一］の対象から外れます。除外した行を選んで押すと元に戻します（行のダブルクリックでも切り替え）。",
                en: "Excludes the selected row: it is unchecked and left out of the counts and of Unify All. Select an excluded row and click again to include it (or double-click the row)."
            },
            zoom: { ja: "一覧で選んだ文字を、この表示倍率で表示します。", en: "The text picked in the list is shown at this zoom." },
            bracketMatch: { ja: "一覧で、見つかった文字を【】で囲んで前後の文字列と見分けやすくします。", en: "Wraps each match in 【】 in the list to set it apart from the surrounding text." },
            unifyTarget: { ja: "統一先にするキーワード。［すべてを統一］で、ほかの欄で見つかった文字をこのキーワードに置き換えます。", en: "The keyword to unify to. Unify All replaces the matches of the other fields with this keyword." },
            unify: { ja: "除外していない（チェックの付いた）行を、統一先のキーワードに置き換えます（1回の取り消しで戻せます）。", en: "Replaces the checked rows with the unify target (undoable in one step)." },
            changeTo: { ja: "一覧で選んだ文字を、この文字に置き換えます。", en: "The text that replaces the match picked in the list." },
            change: { ja: "一覧で選んだ1か所を、［置換文字列］の文字に置き換えます。", en: "Changes the match picked in the list to the Change to text." }
        },
        button: {
            loadSelection: { ja: "選択文字を取り込む", en: "Load Selection" },
            close: { ja: "閉じる", en: "Close" },
            find: { ja: "検索", en: "Find" },
            unify: { ja: "すべてを統一", en: "Unify All" },
            exclude: { ja: "除外する", en: "Exclude" },
            include: { ja: "除外を解除", en: "Include" },
            change: { ja: "置換", en: "Change" },
            collectVariants: { ja: "揺れ候補を拾う", en: "Find Variants" },
            ignore: { ja: "無視", en: "Ignore" },
            copyTable: { ja: "一覧をコピー", en: "Copy Table" }
        },
        status: {
            initial: { ja: "-", en: "-" },
            count: { ja: "{count} 件", en: "{count}" },
            total: { ja: "{count} 件（{scope}）", en: "{count} ({scope})" },
            totalWithExcluded: { ja: "{count} 件（{scope}、除外 {excluded} 件）", en: "{count} ({scope}, {excluded} excluded)" },
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noStory: { ja: "テキストを選択してください。", en: "Please select text." },
            failed: { ja: "検索できませんでした: {message}", en: "Could not search: {message}" },
            truncated: { ja: "{total} 件のうち先頭の {shown} 件を表示", en: "Showing the first {shown} of {total}" },
            overset: { ja: "オーバーセット", en: "Overset" },
            pasteboard: { ja: "ペーストボード", en: "Pasteboard" },
            unavailable: { ja: "（参照できません）", en: "(unavailable)" },
            stale: { ja: "キーワードか設定が変わりました。［検索］で検索し直してください。", en: "Keywords or settings changed. Click Find again." },
            unified: { ja: "{count} 件を「{target}」に置き換えました", en: "Replaced {count} with \"{target}\"" },
            changed: { ja: "「{before}」を「{after}」に置き換えました", en: "Replaced \"{before}\" with \"{after}\"" },
            skipped: { ja: "{count} 件は文字が変わっていたため置き換えませんでした", en: "{count} skipped because the text had changed" },
            nothingToReplace: { ja: "置き換える文字がありません", en: "Nothing to replace" },
            variants: { ja: "{count} 組", en: "{count} groups" },
            noVariants: { ja: "揺れ候補はありません", en: "No variants found" },
            copied: { ja: "{count} 組をコピーしました", en: "Copied {count} groups" }
        },
        alert: {
            confirmUnify: { ja: "{count} 件を「{target}」に置き換えますか？", en: "Replace {count} matches with \"{target}\"?" },
            copyFailed: { ja: "クリップボードにコピーできませんでした。", en: "Could not copy to the clipboard." }
        },
        undo: {
            unify: { ja: "表記を統一", en: "Unify Spelling" },
            change: { ja: "表記を置換", en: "Change Spelling" }
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

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPT_VERSION）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = String(SCRIPT_VERSION);
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

    /* コントロールの寸法 / Control sizes */
    var CLEAR_BUTTON_SIZE = [16, 16];       /* キーワード欄の × の大きさ（px）/ size of the × next to a keyword field */
    var CLEAR_BUTTON_CIRCLE_COLOR = isDarkUI() ? [1, 1, 1, 0.30] : [0, 0, 0, 0.30]; /* × の丸の色（macOS の検索欄の × に近い薄い灰色）/ circle color, a light gray like macOS search fields */
    var CLEAR_BUTTON_MARK_COLOR = isDarkUI() ? [0.2, 0.2, 0.2, 1] : [1, 1, 1, 1]; /* × の線の色（丸を抜いたように見せる）/ color of the × lines, cut out of the circle */
    var KEYWORD_FIELD_WIDTH = 120;          /* キーワード欄の幅（px）/ width of a keyword field */
    var COUNT_TEXT_WIDTH = 52;              /* 欄の右に出す件数の幅（px）/ width of the count next to each field */
    var MATCH_LIST_WIDTH = 300;             /* 該当箇所の一覧の幅（px）/ width of the match list */
    var MATCH_LIST_MIN_HEIGHT = 200;        /* 一覧の最小の高さ（px。左のカラムに合わせて伸びる）/ minimum list height; it grows to the left column */
    var MATCH_LIST_PAGE_COLUMN_WIDTH = 90;  /* 一覧のページ列の幅（px。これより狭いとチェック欄が消える、2026-11-07 実測）/ width of the page column; narrower hides the check boxes */
    var CHANGE_TO_FIELD_WIDTH = 180;          /* ［置換文字列］の幅（px）/ width of the Change to field */
    var TOTAL_ROW_TOP_MARGIN = 5;           /* ［合計］の行の上の余白（px）/ top margin of the Total row */
    var TOTAL_ROW_BOTTOM_MARGIN = 10;       /* ［合計］の行の下の余白（px）/ bottom margin of the Total row */
    var ZOOM_ROW_BOTTOM_MARGIN = 10;        /* ［表示倍率］の行の下の余白（px）/ bottom margin of the Zoom row */
    var CONTEXT_ROW_BOTTOM_MARGIN = 4;      /* ［前後の文字数］の行の下の余白（px）/ bottom margin of the Characters Around row */
    var SLIDER_LABEL_WIDTH = 90;            /* スライダーの項目名の幅（px。2本のスライダーの左端をそろえる）/ width of slider labels, so both sliders line up */
    var SLIDER_VALUE_WIDTH = 40;            /* スライダーの右の数字の幅（px。「800%」が入る幅で2本そろえる）/ width of the value next to a slider, fitting "800%" for both */
    var EXCLUDE_BUTTON_WIDTH = 94;         /* ［除外する］［除外を解除］の幅（px。長いほうの文言の幅、実測93px。広げない）/ width of Exclude / Include, fitted to the longer label */
    var VARIANT_LIST_WIDTH = 230;           /* 揺れ候補の一覧の幅（px）/ width of the variant list */
    var VARIANT_LIST_MIN_HEIGHT = 200;      /* 揺れ候補の一覧の最小の高さ（px）/ minimum height of the variant list */
    var LIST_NOTE_HEIGHT = 36;              /* 一覧の下のメモの高さ（px。2行分）/ height of the note below the list, two lines */
    var CHANGE_BUTTON_WIDTH = 64;            /* ［置換］の幅（px。文字数分にして、［置換文字列］を広くとる）/ width of Change, kept to its label so the field gets the room */

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
            var typeName = item.constructor.name;

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
     * @typedef {Object} HelperSettings
     * @property {string[]} keywords - キーワード欄の文字（KEYWORD_FIELD_COUNT 個。空欄は空文字）
     * @property {string} scope - "story" または "document"
     * @property {boolean} caseSensitive - 大文字と小文字を区別するか
     * @property {boolean} wholeWord - 単語単位で探すか
     * @property {boolean} masterPages - 親ページを含むか（ドキュメントのときだけ）
     * @property {boolean} excludeToc - 目次のストーリーにあるものを除くか
     * @property {boolean} excludeContained - ほかのキーワードに含まれる分を除くか
     * @property {number} contextLength - 一覧に出す前後の文字数
     * @property {boolean} bracketMatch - 一覧で見つかった文字を【】で囲むか
     * @property {boolean} withinParagraph - 一覧の前後の文字を段落の中に限るか
     * @property {Object} variantRules - 組の ID → 前回の統一先
     * @property {Object} ignoredGroups - 無視している組の ID（値は true）
     * @property {number} variantFilter - 揺れ候補の種類の番号（VARIANT_KIND_KEYS の並び）
     * @property {number} variantSort - 揺れ候補の並び順の番号（0 は多い順、1 は少ない順）
     * @property {number} zoomPercent - 一覧で選んだ文字を表示するときの表示倍率（％）
     * @property {number} targetIndex - 統一先の欄の番号（選んでいなければ -1）
     */

    /**
     * 前回の設定を読む
     * @returns {HelperSettings} 設定
     */
    function loadSettings() {
        var saved = $.global[SESSION_KEY] || {};
        var keywords = [];
        for (var i = 0; i < KEYWORD_FIELD_COUNT; i++) {
            keywords.push((saved.keywords && saved.keywords[i] != null) ? String(saved.keywords[i]) : "");
        }
        return {
            keywords: keywords,
            scope: saved.scope || "document",
            caseSensitive: readSavedFlag(saved.caseSensitive, false),
            wholeWord: readSavedFlag(saved.wholeWord, false),
            masterPages: readSavedFlag(saved.masterPages, false),
            excludeToc: readSavedFlag(saved.excludeToc, true),
            excludeContained: readSavedFlag(saved.excludeContained, true),
            contextLength: (typeof saved.contextLength === "number") ? saved.contextLength : DEFAULT_CONTEXT_LENGTH,
            bracketMatch: readSavedFlag(saved.bracketMatch, false),
            withinParagraph: readSavedFlag(saved.withinParagraph, true),
            variantRules: saved.variantRules || {},
            ignoredGroups: saved.ignoredGroups || {},
            variantFilter: (typeof saved.variantFilter === "number") ? saved.variantFilter : 0,
            variantSort: (typeof saved.variantSort === "number") ? saved.variantSort : 0,
            zoomPercent: (typeof saved.zoomPercent === "number") ? saved.zoomPercent : DEFAULT_ZOOM_PERCENT,
            targetIndex: (typeof saved.targetIndex === "number") ? saved.targetIndex : -1
        };
    }

    /**
     * 控えた真偽値を読む（まだ控えていなければ初期値）
     * @param {*} savedValue - 控えた値
     * @param {boolean} defaultValue - 初期値
     * @returns {boolean} 値
     */
    function readSavedFlag(savedValue, defaultValue) {
        return (savedValue == null) ? defaultValue : !!savedValue;
    }

    /**
     * 設定を次回のために控える
     * @param {HelperSettings} settings - 設定
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
     * @property {string} matchedText - 見つけたときの文字（置き換える前に、変わっていないかを確かめる）
     */

    /**
     * 見つかった文字の位置を、入れ物（ストーリー・セル・脚注）と範囲で表す。
     * findText の結果の index は入れ物の中での位置（2026-11-05 実測）
     * @param {Text} foundText - findText で見つかった文字
     * @returns {MatchRange} 見つかった文字と、入れ物の指定子・範囲
     */
    function toMatchRange(foundText) {
        return {
            text: foundText,
            container: foundText.parent.toSpecifier(),
            start: foundText.index,
            end: foundText.index + foundText.length,
            matchedText: String(foundText.contents)
        };
    }

    /**
     * 見つかった文字が目次のストーリー（［目次の作成］で作ったもの）にあるかを返す
     * @param {Text} foundText - 見つかった文字
     * @returns {boolean} 目次の中なら true
     */
    function isInTocStory(foundText) {
        return foundText.parentStory.storyType === StoryTypes.TOC_STORY;
    }

    /**
     * 入れ物の1文字の語の種類を返す（特殊文字は文字種なし）
     * @param {Story|Cell|Footnote} container - 入れ物
     * @param {number} characterIndex - 文字の位置
     * @returns {number} getWordClass() の値
     */
    function getWordClassAt(container, characterIndex) {
        var characterContents = container.characters.item(characterIndex).contents;
        if (typeof characterContents !== "string" || characterContents.length === 0) return 0;
        return getWordClass(characterContents.charCodeAt(0));
    }

    /**
     * 見つかった文字が、前後の同じ文字種に続いていないかを返す（「仕事」の「事」、pineapple の apple は続いている）
     * @param {MatchRange} matchRange - 見つかった位置
     * @param {Object} containerLengths - 入れ物ごとの文字数の控え
     * @returns {boolean} 語として切れていれば true
     */
    function isWholeWordRange(matchRange, containerLengths) {
        var matchedText = matchRange.matchedText;
        var container = matchRange.text.parent;
        var firstClass = getEdgeWordClass(matchedText, 0);
        if (firstClass !== 0 && matchRange.start > 0 && getWordClassAt(container, matchRange.start - 1) === firstClass) return false;
        var lastClass = getEdgeWordClass(matchedText, matchedText.length - 1);
        if (lastClass === 0 || matchRange.end >= getContainerLength(matchRange, containerLengths)) return true;
        return getWordClassAt(container, matchRange.end) !== lastClass;
    }

    /**
     * キーワードごとに見つかった位置を集める
     * @param {Array<Document|Story>} searchTargets - 検索する対象
     * @param {string[]} keywords - キーワード（空文字は探さない）
     * @param {HelperSettings} settings - 設定
     * @returns {Array<Array<MatchRange>>} キーワードごとの位置（keywords と同じ並び）
     */
    function findKeywordRanges(searchTargets, keywords, settings) {
        var findOptions = app.findChangeTextOptions;
        var savedOptions = {};
        for (var i = 0; i < FIND_OPTION_KEYS.length; i++) {
            savedOptions[FIND_OPTION_KEYS[i]] = findOptions[FIND_OPTION_KEYS[i]];
        }

        var rangesByKeyword = [];
        var containerLengths = {};
        try {
            findOptions.caseSensitive = settings.caseSensitive;
            /* InDesign の wholeWord は日本語の語を返さない（2026-11-07 実測）ので、見つけたあとに自前で判定する
               / InDesign's wholeWord returns nothing for Japanese words, so the edges are checked after finding */
            findOptions.wholeWord = false;
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
                    for (var foundIndex = 0; foundIndex < found.length; foundIndex++) {
                        if (settings.excludeToc && isInTocStory(found[foundIndex])) continue;
                        var matchRange = toMatchRange(found[foundIndex]);
                        if (settings.wholeWord && !isWholeWordRange(matchRange, containerLengths)) continue;
                        ranges.push(matchRange);
                    }
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
    function removeContainedRanges(rangesByKeyword, keywords, excludeContained) {
        if (!excludeContained) return rangesByKeyword;
        var keptByKeyword = [];
        for (var k = 0; k < rangesByKeyword.length; k++) {
            var ranges = rangesByKeyword[k];

            /* 自分より長い別のキーワードの一致だけが、自分の一致を包める / Only matches of a different, longer keyword can enclose this one */
            var longerRanges = [];
            for (var otherIndex = 0; otherIndex < rangesByKeyword.length; otherIndex++) {
                if (otherIndex === k || keywords[otherIndex].length <= keywords[k].length) continue;
                longerRanges = longerRanges.concat(rangesByKeyword[otherIndex]);
            }
            var longerByContainer = groupRangesByContainer(longerRanges);

            var keptRanges = [];
            for (var i = 0; i < ranges.length; i++) {
                var candidates = longerByContainer.hasOwnProperty(ranges[i].container) ? longerByContainer[ranges[i].container] : [];
                var isContained = false;
                for (var j = 0; j < candidates.length; j++) {
                    if (candidates[j].start <= ranges[i].start && ranges[i].end <= candidates[j].end) { isContained = true; break; }
                }
                if (!isContained) keptRanges.push(ranges[i]);
            }
            keptByKeyword.push(keptRanges);
        }
        return keptByKeyword;
    }

    /**
     * 今の選択とドキュメントで検索する
     * @param {HelperSettings} settings - 設定
     * @returns {{matchesByKeyword: Array<Array<MatchRange>>|null, message: string}} キーワードごとの位置（数えられなければ null）と合計の行に出す文言
     */
    function runFind(settings) {
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
            var matchesByKeyword = removeContainedRanges(rangesByKeyword, settings.keywords, settings.excludeContained);
            var totalCount = 0;
            for (var i = 0; i < matchesByKeyword.length; i++) totalCount += matchesByKeyword[i].length;
            return { matchesByKeyword: matchesByKeyword, message: getLabel(LABELS.status.total, { count: totalCount, scope: getLabel(LABELS.radio[settings.scope]) }) };
        } catch (e) {
            return { matchesByKeyword: null, message: getLabel(LABELS.status.failed, { message: e.message }) };
        }
    }

    // =========================================
    // 揺れ候補 / Variants
    // =========================================

    /* 表記ゆれ辞書。1行に1組（ここに行を足せます）
       「誤 → 正」は誤記だけあれば候補に出す。「? 誤 → 正」も候補に出すが、文脈で正しいこともあるので統一先を選ばない。
       「A、B」は2つ以上使われているときに候補に出す。先頭に「= 」を付けると先頭の表記を統一先に推奨する
       / The variant dictionary, one group per line (add your own):
         "wrong → right" appears as soon as the wrong spelling is found; "? wrong → right" too, but no target is chosen because the word
         can be right in context; "A、B" appears when two are used; a leading "= " recommends the first spelling */
    var VARIANT_DICTIONARY_LINES = [
        "= ください、下さい",
        "= いただく、頂く",
        "= いただき、頂き",
        "= いただきます、頂きます",
        "= いただいた、頂いた",
        "= できる、出来る",
        "= できます、出来ます",
        "= できない、出来ない",
        "= できた、出来た",
        "= できて、出来て",
        "= こと、事",
        "= とき、時",
        "= ところ、所",
        "= ほか、他",
        "= など、等",
        "= ただし、但し",
        "= および、及び",
        "= または、又は",
        "= ならびに、並びに",
        "= したがって、従って",
        "= すでに、既に",
        "= さらに、更に",
        "= まず、先ず",
        "= あらかじめ、予め",
        "= おそらく、恐らく",
        "= たとえば、例えば",
        "= ぜひ、是非",
        "= よろしく、宜しく",
        "= ありがとう、有難う、有り難う",
        "= お願い、御願い",
        "= ご確認、御確認",
        "= ご連絡、御連絡",
        "= わかる、分かる、解る",
        "= わかりやすい、分かりやすい、分り易い",
        "= スマートフォン、スマホ",
        "= アプリ、アプリケーション",
        "：、:",
        "？、?",
        "！、!",
        "（、(",
        "）、)",
        "％、%",
        "＆、&",
        "〜、～",
        "? 波線 → 破線",
        "パス状文字 → パス上文字",
        "? 補完 → 補間",
        "役物 → 約物",
        "? 抑える → 押さえる",
        "? 改定 → 改訂",
        "? 改定版 → 改訂版",
        "? 意外 → 以外",
        "? 内臓 → 内蔵",
        "? 初稿 → 初校",
        "個条書き → 箇条書き",
        "? 製作 → 制作",
        "? 対象 → 対照",
        "ドラック → ドラッグ",
        "WIn → Win",
        "WIndows → Windows",
        "QuickLook → Quick Look",
        "しずらい → しづらい",
        "日本工業規格 → 日本産業規格",
        "シャギー → ジャギー",
        "ガーター → ガター",
        "レスポンシブル → レスポンシブ",
        "墨付き括弧 → 隅付き括弧",
        "激的 → 劇的",
        "グリーンインストール → クリーンインストール",
        "? タイポ → タイポグラフィ",
        "字詰め → 文字詰め",
        "ぶら下げインデント → 突き出しインデント",
        "ワイアーフレーム → ワイヤーフレーム",
        "ヒヤリング → ヒアリング",
        "ダイアル → ダイヤル",
        "エンジニヤ → エンジニア",
        "フォロアー → フォロワー",
        "クリエーティブ、クリエイティブ",
        "クリエーター、クリエイター",
        "アクティベート、アクティベイト",
        "インターフェイス、インタフェース",
        "プレフィックス、プリフィックス",
        "ウイルス、ウィルス",
        "ウエイト、ウェイト",
        "ウェッブ → ウェブ",
        "アナリティックス → アナリティクス",
        "ソフトウエア、ソフトウェア",
        "メイン、メーン",
        "? することができる → できる",
        "? することが可能 → できる",
        "? より → から",
        "? おります → います",
        "? にて → で",
        "? については → は",
        "? なります → です",
        "? という",
        "? さわる → 操作する",
        "? 触る → 操作する",
        "? 立ち上げる → 起動する",
        "? 打つ → 入力する",
        "? 落ちる → 強制終了する",
        "? 凍る → フリーズする",
        "? 拾う → 取得する",
        "? おく → 保存する、移動する",
        "起点、基点、基準点",
        "フチ、ふち、エッジ、キワ、境界線",
        "ビフォア、ビフォー",
        "白、ホワイト",
        "黒、ブラック",
        "アンカー、アンカーポイント、ポイント",
        "※、＊、*",
        "スペースバー、スペースキー",
        "アンダーバー、アンダースコア",
        "クオリティ、クォリティ、クオリティー",
        "つける、付ける",
        "名詞、名刺",
        "下線、アンダーライン",
        "やさしい、易しい、優しい",
        "半径、判型",
        "クリック、クイック",
        "x、×",
        "ドロップシャドウ、影、シャドウ",
        "ビュレット、バレット、ブレット",
        "カラム、コラム、段組み、段組",
        "コンマ、カンマ",
        "ルビ、読みがな、フリガナ",
        "マス目、升目、枡目、桝目",
        "音引き、伸ばし棒、長音",
        "拗促音、小書きのかな",
        "ビットマップ画像、ラスター画像",
        "切り抜き、トリミング",
        "レイアウト、レーアウト",
        "取り消し線、打ち消し線",
        "スモールキャップ、スモールキャップス、スモールキャピタル",
        "チャネル、チャンネル",
        "PC、Mac、パソコン、マシン、端末",
        "左揃え、左寄せ",
        "右揃え、右寄せ",
        "中央揃え、センター揃え、中央寄せ",
        "行頭行末揃え、ジャスティファイ",
        "マージン、余白",
        "裁断、断裁",
        "トレス、トレース",
        "パタン、パターン",
        "カラーモード、カラースペース",
        "スクリーンショット、ショット",
        "スクリーンキャプチャ、キャプチャ、スクリーンダンプ",
        "アプリケーション、ソフト、ソフトウェア",
        "Indesign → InDesign",
        "PhotoShop → Photoshop",
        "FireFly → Firefly",
        "Firefry → Firefly",
        "LightRoom → Lightroom",
        "AfterEffects → After Effects",
        "Adobe Max → Adobe MAX",
        "ツールパネル → ツールバー",
        "サムネイル → サムネール",
        "ポイントテキスト → ポイント文字",
        "エリア内テキスト → エリア内文字",
        "? パレット → パネル",
        "ウィジット → ウィジェット",
        "キャンバス → カンバス",
        "インターフェース → インターフェイス",
        "ウインドウ → ウィンドウ",
        "フォルダ → フォルダー",
        "? 切り替え → 切り換え",
        "シェープ → シェイプ",
        "? 書く → 描く",
        "? 先端 → 線端",
        "Vanishing Point → 消点",
        "? 注釈 → コメント",
        "クローズドパス → クローズパス",
        "? ベクトル → ベクター",
        "Adobe Systems → Adobe",
        "? Adobe → アドビ",
        "複製を保存 → コピーを保存",
        "? 環境設定 → 設定",
        "アクティベート、追加、インストール",
        "ディアクティベート、削除",
        "PDF変換、PDF書き出し、PDF保存、PDF作成",
        "ジャスティファイ → 均等配置（最終行左揃え）",
        "見積り → 見積もり、見積",
        "打合せ → 打ち合わせ",
        "問合せ → 問い合わせ",
        "? ベット → ベッド",
        "? ドック → ドッグ",
        "ビック → ビッグ",
        "グット → グッド",
        "? バット → バッド",
        "スプレットシート → スプレッドシート",
        "? バック → バッグ",
        "アボガド → アボカド",
        "ギブス → ギプス",
        "? バッチ → バッジ",
        "カッコ → 括弧",
        "ナナメ → 斜め",
        "? アキ → 空き",
        "? ケイ → 罫",
        "カコミ → 囲み",
        "ナリユキ → 成り行き",
        "ケヌキ → 毛抜き",
        "永遠と → 延々と",
        "? 脈略と → 脈々と",
        "? 誠実に → 切実に",
        "AdobeFonts → Adobe Fonts",
        "AdobeFont → Adobe Fonts",
        "AdobeStock → Adobe Stock",
        "amazon → Amazon",
        "android → Android",
        "Balmuda → BALMUDA",
        "BALMUDA、バルミューダ",
        "Bang & olufsen → Bang & Olufsen",
        "Bang&Olufsen → Bang & Olufsen",
        "BOSE → Bose",
        "bose → Bose",
        "CookPad → Cookpad",
        "cookPad → Cookpad",
        "cookpad → Cookpad",
        "css → CSS",
        "DropBox → Dropbox",
        "facebook → Facebook",
        "FaceBook → Facebook",
        "figma → Figma",
        "Freee → freee",
        "github → GitHub",
        "Github → GitHub",
        "google → Google",
        "Google Map → Google Maps",
        "Google Maps、Google マップ",
        "html → HTML",
        "illustrator → Illustrator",
        "instagram → Instagram",
        "javascript → JavaScript",
        "Javascript → JavaScript",
        "keynote → Keynote",
        "KEYNOTE → Keynote",
        "? line → LINE",
        "? Line → LINE",
        "Linkedin → LinkedIn",
        "logicool → Logicool",
        "Mac book → MacBook",
        "Macbook → MacBook",
        "Mac Book Pro → MacBook Pro",
        "MacbookPro → MacBook Pro",
        "mac os → macOS",
        "macos → macOS",
        "miro → Miro",
        "? Note → note",
        "portforio → portfolio",
        "plofile → profile",
        "Protopie → ProtoPie",
        "slack → Slack",
        "snapchat → Snapchat",
        "SnapChat → Snapchat",
        "SnowPeak → Snow Peak",
        "spotify → Spotify",
        "Tiktok → TikTok",
        "tiktok → TikTok",
        "TIPS → Tips",
        "User Voice → UserVoice",
        "WEB → Web",
        "Web、ウェブ",
        "WiFi → Wi-Fi",
        "wifi → Wi-Fi",
        "wikipedia → Wikipedia",
        "WikiPedia → Wikipedia",
        "XCode → Xcode",
        "xcode → Xcode",
        "Yahoo! Japan → Yahoo! JAPAN",
        "Yahoo!JAPAN → Yahoo! JAPAN",
        "Yahoo! News → Yahoo!ニュース",
        "Yahoo! ニュース → Yahoo!ニュース",
        "Youtube → YouTube",
        "zoom → Zoom",
        "ZOOM → Zoom",
        "data base → database",
        "time schedule → timetable",
        "口コミ、クチコミ"
    ];

    /* 組の種類の順（絞り込みの項目の並び）/ Kinds of group, in the order of the filter */
    var VARIANT_KIND_KEYS = ["all", "katakana", "kanji", "latin", "suru", "dictionary", "review"];

    /* 語の種類 / Word classes */
    var WORD_CLASS_KATAKANA = 1;
    var WORD_CLASS_KANJI = 2;
    var WORD_CLASS_LATIN = 3;

    /* 語の切り出し（カタカナの連続・英数字の語・漢字の連続）/ Cuts out katakana runs, Latin words and kanji runs */
    var VARIANT_WORD_PATTERN = /[ァ-ヺー・ｦ-ﾟ]+|[A-Za-z0-9０-９Ａ-Ｚａ-ｚ\-]+|[㐀-鿿豈-﫿々]+/g;

    /* 漢字のあいだの1文字（取り引き・申し込み）。「入力する場合」と「入力した場合」を同じ語にしないよう、名詞になる形だけ
       / One kana between kanji in a noun; only noun forms, so that 入力する場合 and 入力した場合 stay apart */
    var INNER_OKURIGANA = "きりしちみいぎびけえせめれげべて";
    /* 語の終わりの送り仮名（取引き・売上げ）/ Okurigana that ends a noun */
    var ENDING_OKURIGANA = "きりいみけげえちれてせめ";
    /* 2文字の語尾（打ち合わせ・見積もり）/ Two-kana endings */
    var LONG_OKURIGANA = ["わせ", "もり"];
    /* 名詞のあとに続くかな。ほかのかなが続けば活用語とみなす / Kana that may follow a noun; any other means an inflected word */
    var NOUN_FOLLOWERS = "はがをにのでともへやかねよなだま";
    /* 複合語に続く「する」の形（変更する・変更して・変更され・変更せず）/ Forms of する after a compound */
    var SURU_FORMS = "すしさせ";
    /* 「を」の有無で比べる「する」の形。長いものから。形が異なると置き換えられないので形ごとに別の組にする
       / Forms of する compared with and without を, longest first; each form is its own group */
    var SURU_ENDINGS = ["しました", "しません", "します", "しない", "すれば", "して", "した", "する"];
    /* 文頭で大文字になるだけの語 / Words that change case only at the start of a sentence */
    var COMMON_WORDS = "|a|an|and|are|as|at|be|but|by|for|from|he|i|if|in|is|it|its|no|not|of|on|or|our|she|so|that|the|their|there|these|they|this|to|we|what|when|which|with|you|your|";

    /* 半角カタカナの変換表（U+FF66 からの並び）/ Half-width katakana, in order from U+FF66 */
    var HALF_KATAKANA_TABLE = "ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン";

    /**
     * 文字コードの語の種類を返す（語が続いているかの判定に使う。ひらがなには語の切れ目が無いので 0）
     * @param {number} charCode - 文字コード
     * @returns {number} WORD_CLASS_* または 0
     */
    function getWordClass(charCode) {
        if ((charCode >= 0x30A1 && charCode <= 0x30FA) || charCode === 0x30FC || (charCode >= 0xFF66 && charCode <= 0xFF9F)) return WORD_CLASS_KATAKANA;
        if ((charCode >= 0x3400 && charCode <= 0x9FFF) || (charCode >= 0xF900 && charCode <= 0xFAFF) || charCode === 0x3005) return WORD_CLASS_KANJI;
        if ((charCode >= 0x30 && charCode <= 0x39) || (charCode >= 0x41 && charCode <= 0x5A) || (charCode >= 0x61 && charCode <= 0x7A) || charCode === 0x5F) return WORD_CLASS_LATIN;
        if ((charCode >= 0xFF10 && charCode <= 0xFF19) || (charCode >= 0xFF21 && charCode <= 0xFF3A) || (charCode >= 0xFF41 && charCode <= 0xFF5A)) return WORD_CLASS_LATIN;
        return 0;
    }

    /**
     * 語の端の文字の語の種類を返す。漢字の端を見るのはすべて漢字の語だけ（「確認下さい」の「下さい」は語のまま）
     * @param {string} wordText - 語
     * @param {number} edgeIndex - 端の文字の位置
     * @returns {number} WORD_CLASS_* または 0
     */
    function getEdgeWordClass(wordText, edgeIndex) {
        var edgeClass = getWordClass(wordText.charCodeAt(edgeIndex));
        if (edgeClass !== WORD_CLASS_KANJI) return edgeClass;
        for (var i = 0; i < wordText.length; i++) {
            if (getWordClass(wordText.charCodeAt(i)) !== WORD_CLASS_KANJI) return 0;
        }
        return edgeClass;
    }

    /**
     * 文字列の前後の空白を除く
     * @param {string} rawText - 文字列
     * @returns {string} 除いた文字列
     */
    function trimSpaces(rawText) {
        return rawText.replace(/^\s+/, "").replace(/\s+$/, "");
    }

    /**
     * 「A、B」の形の文字列を表記の配列にする
     * @param {string} namesText - 文字列
     * @returns {string[]} 表記（空は除く）
     */
    function splitDictionaryNames(namesText) {
        var parts = namesText.split("、");
        var names = [];
        for (var i = 0; i < parts.length; i++) {
            var name = trimSpaces(parts[i]);
            if (name !== "") names.push(name);
        }
        return names;
    }

    /**
     * @typedef {Object} DictionaryEntry
     * @property {string} id - 組の ID
     * @property {string} mode - "variants"（2つ以上で候補）、"correction"（誤記だけで候補）、"review"（要確認）
     * @property {string[]} spellings - 表記
     * @property {string[]} wrong - あれば候補になる表記（correction と review）
     * @property {string|null} target - 統一先（無ければ null）
     */

    /**
     * 辞書の1行を読む
     * @param {string} rawLine - 1行
     * @returns {DictionaryEntry|null} 組。組にならない行は null
     */
    function parseDictionaryLine(rawLine) {
        var line = trimSpaces(rawLine);
        var isReview = line.charAt(0) === "?";
        var prefersFirst = line.charAt(0) === "=";
        if (isReview || prefersFirst) line = line.substring(1);
        var arrowIndex = line.indexOf("→");
        if (arrowIndex >= 0 || isReview) {
            var wrong = splitDictionaryNames(arrowIndex >= 0 ? line.substring(0, arrowIndex) : line);
            var right = (arrowIndex >= 0) ? splitDictionaryNames(line.substring(arrowIndex + 1)) : [];
            if (wrong.length === 0) return null;
            var target = (right.length > 0) ? right[0] : null;
            var mode = isReview ? "review" : "correction";
            return { id: (isReview ? "r:" : "c:") + (target || wrong[0]), mode: mode, spellings: wrong.concat(right), wrong: wrong, target: target };
        }
        var spellings = splitDictionaryNames(line);
        if (spellings.length < 2) return null;
        return { id: "d:" + spellings.join("|"), mode: "variants", spellings: spellings, wrong: [], target: prefersFirst ? spellings[0] : null };
    }

    /**
     * 表記の組み合わせを比べるための文字列を返す（並びによらない）
     * @param {string[]} spellings - 表記
     * @returns {string} 比較用の文字列
     */
    function getSpellingSetKey(spellings) {
        return spellings.slice(0).sort().join("\u0001");
    }

    /**
     * 辞書を読み、同じ統一先の行を1つの組にまとめ、修正と同じ内容の揺れの行を除く
     * @param {string[]} dictionaryLines - 辞書の行
     * @returns {Array<DictionaryEntry>} 組
     */
    function loadVariantDictionary(dictionaryLines) {
        var entries = [];
        var entryById = {};
        var fixedSets = {};
        for (var i = 0; i < dictionaryLines.length; i++) {
            var entry = parseDictionaryLine(dictionaryLines[i]);
            if (!entry) continue;
            if (entryById.hasOwnProperty(entry.id)) {
                mergeDictionaryNames(entryById[entry.id].wrong, entry.wrong);
                mergeDictionaryNames(entryById[entry.id].spellings, entry.spellings);
            } else {
                entryById[entry.id] = entry;
                entries.push(entry);
            }
        }
        for (var j = 0; j < entries.length; j++) {
            if (entries[j].mode !== "variants") fixedSets[getSpellingSetKey(entries[j].spellings)] = true;
        }
        var kept = [];
        for (var k = 0; k < entries.length; k++) {
            if (entries[k].mode === "variants" && fixedSets.hasOwnProperty(getSpellingSetKey(entries[k].spellings))) continue;
            kept.push(entries[k]);
        }
        return kept;
    }

    /**
     * 配列に無い表記だけを足す
     * @param {string[]} names - 足す先
     * @param {string[]} additions - 足す表記
     * @returns {void}
     */
    function mergeDictionaryNames(names, additions) {
        for (var i = 0; i < additions.length; i++) {
            if (indexOfText(names, additions[i]) < 0) names.push(additions[i]);
        }
    }

    /**
     * 配列の中の文字列の位置を返す（ExtendScript の配列に indexOf は無い）
     * @param {string[]} texts - 配列
     * @param {string} wantedText - 探す文字列
     * @returns {number} 位置。無ければ -1
     */
    function indexOfText(texts, wantedText) {
        for (var i = 0; i < texts.length; i++) {
            if (texts[i] === wantedText) return i;
        }
        return -1;
    }

    /**
     * 全角英数字を半角にし、小文字にする
     * @param {string} rawText - 文字列
     * @returns {string} そろえた文字列
     */
    function foldLatinWord(rawText) {
        var folded = "";
        for (var i = 0; i < rawText.length; i++) {
            var charCode = rawText.charCodeAt(i);
            folded += String.fromCharCode((charCode >= 0xFF10 && charCode <= 0xFF5A) ? charCode - 0xFEE0 : charCode);
        }
        return folded.toLowerCase();
    }

    /**
     * 英数字の語の組の名前を返す
     * @param {string} wordText - 語
     * @returns {string|null} 組の名前。比べない語は null
     */
    function getLatinKey(wordText) {
        var word = foldLatinWord(wordText).split("-").join("");
        /* 数字は 12 と １２ が両方あるときだけ組になる / Numbers group only when 12 and １２ both appear */
        if (word.length < 2 && !/^[0-9]+$/.test(word)) return null;
        if (COMMON_WORDS.indexOf("|" + word + "|") >= 0) return null;
        return "l:" + word;
    }

    /**
     * カタカナの語の文字をそろえる（半角→全角、小さいかな→大きいかな、ヴァ→バ、中黒を除く）
     * @param {string} wordText - 語
     * @returns {string[]} そろえた文字の配列
     */
    function foldKatakanaWord(wordText) {
        var letters = [];
        var vuForms = { "ァ": "バ", "ィ": "ビ", "ェ": "ベ", "ォ": "ボ" };
        var smallForms = { "ァ": "ア", "ィ": "イ", "ゥ": "ウ", "ェ": "エ", "ォ": "オ" };
        for (var i = 0; i < wordText.length; i++) {
            var charCode = wordText.charCodeAt(i);
            var letter = wordText.charAt(i);
            if (letter === "・") continue;
            if (charCode === 0xFF9E || charCode === 0xFF9F) {
                /* 半角の濁点・半濁点は前の字に足す（ウの濁点はヴ）/ Half-width voiced marks join the previous letter */
                var previous = letters.pop() || "";
                var markStep = (charCode === 0xFF9E) ? 1 : 2;
                letters.push((previous === "ウ") ? "ヴ" : String.fromCharCode(previous.charCodeAt(0) + markStep));
                continue;
            }
            if (charCode >= 0xFF66 && charCode <= 0xFF9D) letter = HALF_KATAKANA_TABLE.charAt(charCode - 0xFF66);
            if (letters.length > 0 && letters[letters.length - 1] === "ヴ" && vuForms.hasOwnProperty(letter)) {
                letters[letters.length - 1] = vuForms[letter];
                continue;
            }
            letters.push(smallForms.hasOwnProperty(letter) ? smallForms[letter] : letter);
        }
        for (var j = 0; j < letters.length; j++) {
            if (letters[j] === "ヴ") letters[j] = "ブ";
        }
        return letters;
    }

    /**
     * カタカナの語の組の名前を返す。長音は長い語の中だけ無視する（ユーザビリティとユーザービリティは同じ語、ビルとビールは別の語）
     * @param {string} wordText - 語
     * @returns {string|null} 組の名前。短すぎる語は null
     */
    function getKatakanaKey(wordText) {
        var letters = foldKatakanaWord(wordText);
        while (letters.length > 0 && letters[letters.length - 1] === "ー") letters.pop();
        var withoutLong = [];
        for (var i = 0; i < letters.length; i++) {
            if (letters[i] !== "ー") withoutLong.push(letters[i]);
        }
        if (withoutLong.length < 2) return null;
        return "k:" + (withoutLong.length >= 4 ? withoutLong : letters).join("");
    }

    /**
     * ひらがなかどうかを返す
     * @param {string} sourceText - 本文
     * @param {number} textIndex - 位置
     * @returns {boolean} ひらがななら true
     */
    function isHiraganaAt(sourceText, textIndex) {
        var charCode = sourceText.charCodeAt(textIndex);
        return charCode >= 0x3041 && charCode <= 0x309F;
    }

    /**
     * 漢字の語の終わりを探す。活用の途中の語（取り消す・切り替える）は語幹が欠けた表記になるので語にしない。
     * 「する」が続く複合語（変更せず・確認して）は漢字の部分を語にする
     * @param {string} sourceText - 本文
     * @param {number} runEnd - 漢字の連続の終わり
     * @returns {{end: number, isNoun: boolean}} 語の終わりと、語として数えるか
     */
    function findKanjiWordEnd(sourceText, runEnd) {
        var end = runEnd;
        var hasInner = false;
        while (end + 1 < sourceText.length && INNER_OKURIGANA.indexOf(sourceText.charAt(end)) >= 0 && getWordClass(sourceText.charCodeAt(end + 1)) === WORD_CLASS_KANJI) {
            end += 2;
            hasInner = true;
            while (end < sourceText.length && getWordClass(sourceText.charCodeAt(end)) === WORD_CLASS_KANJI) end++;
        }
        var stem = end;
        if (indexOfText(LONG_OKURIGANA, sourceText.substr(end, 2)) >= 0) end += 2;
        else if (end < sourceText.length && (ENDING_OKURIGANA.indexOf(sourceText.charAt(end)) >= 0 || (hasInner && sourceText.charAt(end) === "し"))) end++;
        if (end >= sourceText.length || !isHiraganaAt(sourceText, end) || NOUN_FOLLOWERS.indexOf(sourceText.charAt(end)) >= 0) return { end: end, isNoun: true };
        return { end: stem, isNoun: !hasInner && stem < sourceText.length && SURU_FORMS.indexOf(sourceText.charAt(stem)) >= 0 };
    }

    /**
     * 名詞のあとの「する」の形を返す（調整する・調整をする のどちらも「する」）
     * @param {string} sourceText - 本文
     * @param {number} wordEnd - 名詞の終わり
     * @returns {{ending: string, end: number}|null} 形と、その終わりの位置。続かなければ null
     */
    function findSuruEnding(sourceText, wordEnd) {
        var start = (sourceText.charAt(wordEnd) === "を") ? wordEnd + 1 : wordEnd;
        for (var i = 0; i < SURU_ENDINGS.length; i++) {
            if (sourceText.substr(start, SURU_ENDINGS[i].length) === SURU_ENDINGS[i]) return { ending: SURU_ENDINGS[i], end: start + SURU_ENDINGS[i].length };
        }
        return null;
    }

    /**
     * 組の名前ごとに表記の件数を足す
     * @param {Object} foundWords - 組の名前 → { counts, order }
     * @param {string} wordKey - 組の名前
     * @param {string} surfaceText - 表記
     * @returns {void}
     */
    function addFoundWord(foundWords, wordKey, surfaceText) {
        if (!foundWords.hasOwnProperty(wordKey)) foundWords[wordKey] = { counts: {}, order: [] };
        var word = foundWords[wordKey];
        if (!word.counts.hasOwnProperty(surfaceText)) {
            word.counts[surfaceText] = 0;
            word.order.push(surfaceText);
        }
        word.counts[surfaceText]++;
    }

    /**
     * 本文の語を切り出して、組の名前ごとに表記の件数を数える
     * @param {string} sourceText - 本文
     * @returns {Object} 組の名前 → { counts, order }
     */
    function collectWords(sourceText) {
        var foundWords = {};
        var wordPattern = new RegExp(VARIANT_WORD_PATTERN.source, "g");
        var wordMatch;
        while ((wordMatch = wordPattern.exec(sourceText)) !== null) {
            var start = wordMatch.index;
            var end = start + wordMatch[0].length;
            var firstClass = getWordClass(sourceText.charCodeAt(start));
            var wordKey = null;
            var isPlainNoun = true;
            if (firstClass === WORD_CLASS_KANJI) {
                var kanjiWord = findKanjiWordEnd(sourceText, end);
                end = kanjiWord.end;
                wordPattern.lastIndex = end;
                var kanjiLetters = sourceText.substring(start, end).replace(/[ぁ-ゟ]/g, "");
                isPlainNoun = kanjiLetters.length === end - start;
                if (kanjiWord.isNoun && kanjiLetters.length >= 2) wordKey = "j:" + kanjiLetters;
            } else {
                /* 端の中黒とハイフンは語に含めない / Middle dots and hyphens at the edges are not part of the word */
                while (start < end && /[・\-]/.test(sourceText.charAt(start))) start++;
                while (end > start && /[・\-]/.test(sourceText.charAt(end - 1))) end--;
                if (end > start) {
                    var wordText = sourceText.substring(start, end);
                    wordKey = (getWordClass(wordText.charCodeAt(0)) === WORD_CLASS_KATAKANA) ? getKatakanaKey(wordText) : getLatinKey(wordText);
                    if (getWordClass(wordText.charCodeAt(0)) === WORD_CLASS_LATIN) isPlainNoun = false;
                }
            }
            if (wordKey === null) continue;
            addFoundWord(foundWords, wordKey, sourceText.substring(start, end));
            var suru = isPlainNoun ? findSuruEnding(sourceText, end) : null;
            if (suru) addFoundWord(foundWords, "s:" + wordKey.substring(2) + suru.ending, sourceText.substring(start, suru.end));
        }
        return foundWords;
    }

    /**
     * 数値を、大きいものが先に並ぶ文字列にする（比較関数つき sort() を避けるためのキー）
     * @param {number} value - 値
     * @returns {string} 8桁の文字列
     */
    function toDescendingKey(value) {
        return String(100000000 + 99999999 - value).substring(1);
    }

    /**
     * 表記を件数の多い順に並べる（同じ件数は元の並びのまま）
     * @param {Array<{text: string, count: number}>} spellings - 表記
     * @returns {Array<{text: string, count: number}>} 並べた新しい配列
     */
    function sortSpellings(spellings) {
        var sortKeys = [];
        for (var i = 0; i < spellings.length; i++) {
            sortKeys.push(toDescendingKey(spellings[i].count) + "\u0001" + toDescendingKey(99999999 - i) + "\u0001" + i);
        }
        sortKeys.sort();
        var sorted = [];
        for (var j = 0; j < sortKeys.length; j++) sorted.push(spellings[Number(sortKeys[j].split("\u0001")[2])]);
        return sorted;
    }

    /**
     * 本文の中の表記の位置を返す（語として切れているものだけ）
     * @param {string} sourceText - 本文
     * @param {string} spelling - 表記
     * @returns {number[]} 開始位置
     */
    function findSpellingHits(sourceText, spelling) {
        var hits = [];
        var firstClass = getEdgeWordClass(spelling, 0);
        var lastClass = getEdgeWordClass(spelling, spelling.length - 1);
        var hitIndex = sourceText.indexOf(spelling);
        while (hitIndex >= 0) {
            var hitEnd = hitIndex + spelling.length;
            var runsOn = (firstClass !== 0 && hitIndex > 0 && getWordClass(sourceText.charCodeAt(hitIndex - 1)) === firstClass)
                || (lastClass !== 0 && hitEnd < sourceText.length && getWordClass(sourceText.charCodeAt(hitEnd)) === lastClass);
            if (!runsOn) hits.push(hitIndex);
            hitIndex = sourceText.indexOf(spelling, hitEnd);
        }
        return hits;
    }

    /**
     * 辞書の組のうち、本文に現れるものを返す。長い表記の中の短い表記は数えない（「文字詰め」の中の「字詰め」）
     * @param {Array<DictionaryEntry>} entries - 辞書
     * @param {string} sourceText - 本文
     * @returns {Array<VariantGroup>} 組
     */
    function findDictionaryGroups(entries, sourceText) {
        var groups = [];
        for (var i = 0; i < entries.length; i++) {
            var entry = entries[i];
            var hitsBySpelling = [];
            for (var j = 0; j < entry.spellings.length; j++) hitsBySpelling.push(findSpellingHits(sourceText, entry.spellings[j]));
            var spellings = [];
            var usedCount = 0;
            var hasWrong = false;
            for (var k = 0; k < entry.spellings.length; k++) {
                var count = countUncontainedHits(entry.spellings, hitsBySpelling, k);
                if (count > 0) usedCount++;
                if (count > 0 && indexOfText(entry.wrong, entry.spellings[k]) >= 0) hasWrong = true;
                /* 正しい表記は、本文で使われていなくても組に入れる / The right spelling stays even when the text does not use it */
                if (count > 0 || (entry.mode !== "variants" && entry.spellings[k] === entry.target)) spellings.push({ text: entry.spellings[k], count: count });
            }
            if (entry.mode === "variants" ? usedCount < 2 : !hasWrong) continue;
            groups.push({ id: entry.id, spellings: sortSpellings(spellings), recommended: (entry.mode === "review") ? null : entry.target });
        }
        return groups;
    }

    /**
     * 同じ組の、より長い表記の中に収まる分を除いた件数を返す
     * @param {string[]} spellings - 組の表記
     * @param {Array<number[]>} hitsBySpelling - 表記ごとの位置
     * @param {number} spellingIndex - 数える表記の番号
     * @returns {number} 件数
     */
    function countUncontainedHits(spellings, hitsBySpelling, spellingIndex) {
        var hits = hitsBySpelling[spellingIndex];
        var ownLength = spellings[spellingIndex].length;
        var count = 0;
        for (var i = 0; i < hits.length; i++) {
            var isContained = false;
            for (var j = 0; j < spellings.length && !isContained; j++) {
                if (spellings[j].length <= ownLength) continue;
                for (var k = 0; k < hitsBySpelling[j].length; k++) {
                    var longerStart = hitsBySpelling[j][k];
                    if (longerStart <= hits[i] && hits[i] + ownLength <= longerStart + spellings[j].length) { isContained = true; break; }
                }
            }
            if (!isContained) count++;
        }
        return count;
    }

    /**
     * @typedef {Object} VariantGroup
     * @property {string} id - 組の ID（先頭2文字が種類。k: カタカナ、j: 漢字、l: 英数字、s: する／をする、c:・d: 辞書、r: 要確認）
     * @property {Array<{text: string, count: number}>} spellings - 表記（多い順）
     * @property {string|null} [recommended] - 辞書が推奨する統一先
     */

    /**
     * 組の種類を返す
     * @param {VariantGroup} group - 組
     * @returns {string} VARIANT_KIND_KEYS のいずれか
     */
    function getVariantKind(group) {
        var kinds = { "k:": "katakana", "j:": "kanji", "l:": "latin", "s:": "suru", "r:": "review" };
        var prefix = group.id.substring(0, 2);
        return kinds.hasOwnProperty(prefix) ? kinds[prefix] : "dictionary";
    }

    /**
     * 組が、すでにある組の表記に含まれるかを返す
     * @param {Array<{text: string}>} spellings - 組の表記
     * @param {Array<VariantGroup>} groups - すでにある組
     * @returns {boolean} 含まれれば true
     */
    function isCoveredByGroups(spellings, groups) {
        for (var i = 0; i < groups.length; i++) {
            var groupTexts = [];
            for (var j = 0; j < groups[i].spellings.length; j++) groupTexts.push(groups[i].spellings[j].text);
            var isCovered = true;
            for (var k = 0; k < spellings.length && isCovered; k++) {
                if (indexOfText(groupTexts, spellings[k].text) < 0) isCovered = false;
            }
            if (isCovered) return true;
        }
        return false;
    }

    /**
     * 本文の中で表記が揺れている語の組を、出現数の多い順に返す
     * @param {string} sourceText - 本文
     * @param {Array<DictionaryEntry>} dictionaryEntries - 辞書
     * @returns {Array<VariantGroup>} 組
     */
    function findVariantGroups(sourceText, dictionaryEntries) {
        var groups = findDictionaryGroups(dictionaryEntries, sourceText);
        var dictionaryGroupCount = groups.length;
        var foundWords = collectWords(sourceText);
        for (var wordKey in foundWords) {
            if (!foundWords.hasOwnProperty(wordKey) || foundWords[wordKey].order.length < 2) continue;
            var spellings = [];
            for (var i = 0; i < foundWords[wordKey].order.length; i++) {
                var surfaceText = foundWords[wordKey].order[i];
                spellings.push({ text: surfaceText, count: foundWords[wordKey].counts[surfaceText] });
            }
            if (isCoveredByGroups(spellings, groups.slice(0, dictionaryGroupCount))) continue;
            groups.push({ id: wordKey, spellings: sortSpellings(spellings), recommended: null });
        }

        var sortKeys = [];
        for (var j = 0; j < groups.length; j++) {
            var total = 0;
            for (var k = 0; k < groups[j].spellings.length; k++) total += groups[j].spellings[k].count;
            sortKeys.push(toDescendingKey(total) + "\u0001" + groups[j].id + "\u0001" + j);
        }
        sortKeys.sort();
        var sorted = [];
        for (var n = 0; n < sortKeys.length; n++) {
            var parts = sortKeys[n].split("\u0001");
            sorted.push(groups[Number(parts[parts.length - 1])]);
        }
        return sorted;
    }

    /**
     * ストーリーが親ページだけにあるかを返す
     * @param {Story} story - ストーリー
     * @returns {boolean} 親ページのストーリーなら true
     */
    function isMasterStory(story) {
        if (story.textContainers.length === 0) return false;
        var firstFrame = story.textContainers[0];
        /* パス上文字はパスの親からページをたどる / Text on a path finds its page through the path's owner */
        var page = (firstFrame.constructor.name === "TextPath") ? firstFrame.parent.parentPage : firstFrame.parentPage;
        return !!page && page.parent.constructor.name === "MasterSpread";
    }

    /**
     * 表のセルと脚注の文字をストーリーの本文に足す
     * @param {Story} story - ストーリー
     * @param {string[]} textParts - 足す先
     * @returns {void}
     */
    function addStoryTexts(story, textParts) {
        textParts.push(String(story.contents));
        for (var i = 0; i < story.tables.length; i++) textParts.push(story.tables[i].cells.everyItem().contents.join("\r"));
        if (story.footnotes.length > 0) textParts.push(story.footnotes.everyItem().contents.join("\r"));
    }

    /**
     * ［対象］の文字をまとめて返す（検索と同じ範囲。表・脚注を含み、［親ページを含む］［目次を除く］に従う）
     * @param {HelperSettings} settings - 設定
     * @returns {string|null} 本文。対象が無ければ null
     */
    function collectVariantText(settings) {
        if (app.documents.length === 0) return null;
        var stories = [];
        if (settings.scope === "story") collectStories(app.selection, stories, {});
        else stories = app.activeDocument.stories.everyItem().getElements();
        if (stories.length === 0) return null;
        var textParts = [];
        for (var i = 0; i < stories.length; i++) {
            if (settings.excludeToc && stories[i].storyType === StoryTypes.TOC_STORY) continue;
            if (settings.scope === "document" && !settings.masterPages && isMasterStory(stories[i])) continue;
            addStoryTexts(stories[i], textParts);
        }
        return textParts.join("\r");
    }

    /**
     * 文字列をクリップボードにコピーする（Mac は AppleScript、Windows は VBScript 経由）
     * @param {string} copyText - コピーする文字列
     * @returns {boolean} コピーできれば true
     */
    function copyToClipboard(copyText) {
        var isMac = (File.fs === "Macintosh");
        var scriptSource = isMac
            ? 'set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"'
            : 'CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""').replace(/\r?\n/g, '" & vbCrLf & "').replace(/\t/g, '" & vbTab & "') + '"';
        /* スクリプト実行の許可がないと失敗する / Fails without permission to run scripts */
        try {
            app.doScript(scriptSource, isMac ? ScriptLanguage.APPLESCRIPT_LANGUAGE : ScriptLanguage.VISUAL_BASIC);
            return true;
        } catch (e) {
            alert(getLabel(LABELS.alert.copyFailed) + "\n" + e.message);
            return false;
        }
    }

    // =========================================
    // 該当箇所の一覧 / Match list
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
     * 見つかった文字の入れ物の文字数を返す（同じ入れ物は1回だけ数えて控える）
     * @param {MatchRange} matchRange - 見つかった位置
     * @param {Object} containerLengths - 入れ物ごとの文字数の控え
     * @returns {number} 文字数
     */
    function getContainerLength(matchRange, containerLengths) {
        if (!containerLengths.hasOwnProperty(matchRange.container)) containerLengths[matchRange.container] = matchRange.text.parent.characters.length;
        return containerLengths[matchRange.container];
    }

    /**
     * 見つかった文字を前後の文字列と合わせて返す
     * @param {MatchRange} matchRange - 見つかった位置
     * @param {number} contextLength - 前後の文字数
     * @param {Object} containerLengths - 入れ物ごとの文字数（同じ入れ物を何度も数えないための控え）
     * @param {boolean} useBrackets - 見つかった文字を【】で囲むか
     * @param {boolean} withinParagraph - 前後の文字を段落の中に限るか
     * @returns {string} 一覧に出す文字列
     */
    function buildContextText(matchRange, contextLength, containerLengths, useBrackets, withinParagraph) {
        var container = matchRange.text.parent;
        var containerLength = getContainerLength(matchRange, containerLengths);
        var beforeText = getRangeText(container, Math.max(0, matchRange.start - contextLength), matchRange.start - 1);
        var afterText = getRangeText(container, matchRange.end, Math.min(containerLength - 1, matchRange.end + contextLength - 1));
        if (withinParagraph) {
            beforeText = beforeText.substring(Math.max(beforeText.lastIndexOf("\r"), beforeText.lastIndexOf("\n")) + 1);
            afterText = afterText.split("\n")[0].split("\r")[0];
        }
        var matchText = toDisplayText(matchRange.matchedText);
        if (useBrackets) matchText = "【" + matchText + "】";
        return toDisplayText(beforeText) + matchText + toDisplayText(afterText);
    }

    /**
     * 前後の文字数を整数に丸め、0〜MAX_CONTEXT_LENGTH に収める
     * @param {number} contextLength - 前後の文字数
     * @returns {number} 丸めた文字数
     */
    function clampContextLength(contextLength) {
        if (isNaN(contextLength)) return DEFAULT_CONTEXT_LENGTH;
        return Math.max(0, Math.min(MAX_CONTEXT_LENGTH, Math.round(contextLength)));
    }

    /**
     * 表示倍率を ZOOM_STEP_PERCENT 単位に丸め、スライダーの範囲に収める
     * @param {number} zoomPercent - 表示倍率（％）
     * @returns {number} 丸めた表示倍率
     */
    function snapZoomPercent(zoomPercent) {
        if (isNaN(zoomPercent)) return DEFAULT_ZOOM_PERCENT;
        var snapped = Math.round(zoomPercent / ZOOM_STEP_PERCENT) * ZOOM_STEP_PERCENT;
        return Math.max(MIN_ZOOM_PERCENT, Math.min(MAX_ZOOM_PERCENT, snapped));
    }

    /**
     * 作業中のレイアウトウィンドウの表示倍率を変える（ストーリーエディターなどレイアウト以外のウィンドウでは何もしない）
     * @param {number} zoomPercent - 表示倍率（％）
     * @returns {void}
     */
    function setLayoutZoom(zoomPercent) {
        if (app.layoutWindows.length === 0) return;
        var activeWindow = app.activeWindow;
        if (activeWindow.constructor.name !== "LayoutWindow") return;
        activeWindow.zoomPercentage = zoomPercent;
    }

    // =========================================
    // 置き換え / Replacing
    // =========================================

    /**
     * 見つかった文字を置き換える（1回の取り消しで戻せるようにまとめる）。
     * 同じ入れ物の中は後ろから置き換え、前の置き換えで後ろの位置がずれないようにする。
     * 見つけたときから文字が変わっているもの・消えたものは飛ばす
     * @param {Array<MatchRange>} matchRanges - 置き換える位置
     * @param {string} replacementText - 置き換える文字
     * @param {string} undoName - 取り消しの名前
     * @returns {{replaced: number, skipped: number, replacedRanges: Array<MatchRange>}} 置き換えた件数・飛ばした件数・置き換えた位置
     */
    function replaceMatches(matchRanges, replacementText, undoName) {
        var replaceResult = { replaced: 0, skipped: 0, replacedRanges: [] };
        var rangesByContainer = groupRangesByContainer(matchRanges);
        var orderedRanges = [];
        for (var containerKey in rangesByContainer) {
            if (!rangesByContainer.hasOwnProperty(containerKey)) continue;
            var descendingRanges = sortRangesDescending(rangesByContainer[containerKey]);
            /* 重なる位置（「取引」と「取引き」が同じ場所など）は後ろ側を優先し、残りは飛ばす / for overlapping ranges keep the later one and skip the rest */
            var nextStart = Infinity;
            for (var i = 0; i < descendingRanges.length; i++) {
                if (descendingRanges[i].end > nextStart) { replaceResult.skipped++; continue; }
                orderedRanges.push(descendingRanges[i]);
                nextStart = descendingRanges[i].start;
            }
        }

        app.doScript(function () {
            for (var k = 0; k < orderedRanges.length; k++) {
                var matchRange = orderedRanges[k];
                try {
                    if (!matchRange.text.isValid || String(matchRange.text.contents) !== matchRange.matchedText) {
                        replaceResult.skipped++;
                        continue;
                    }
                    matchRange.text.contents = replacementText;
                    replaceResult.replaced++;
                    replaceResult.replacedRanges.push(matchRange);
                } catch (e) {
                    replaceResult.skipped++;
                }
            }
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, undoName);
        return replaceResult;
    }

    /**
     * 位置を開始位置の大きい順に並べる（比較関数つき sort() は ExtendScript で遅く並びも狂うので、挿入ソートで並べる）
     * @param {Array<MatchRange>} ranges - 位置
     * @returns {Array<MatchRange>} 大きい順に並べた新しい配列
     */
    function sortRangesDescending(ranges) {
        var sortedRanges = [];
        for (var i = 0; i < ranges.length; i++) {
            var insertAt = sortedRanges.length;
            while (insertAt > 0 && sortedRanges[insertAt - 1].start < ranges[i].start) insertAt--;
            sortedRanges.splice(insertAt, 0, ranges[i]);
        }
        return sortedRanges;
    }

    // =========================================
    // パレット / Palette
    // =========================================

    /**
     * 右のパネルの余った高さを一覧に足す。
     * 縦並びのパネルでは listbox の alignment "fill" が縦に効かず、余りがパネルの上や行の間に配られるため、
     * 一度レイアウトしてから余りを測り、一覧の高さに足して並べ直す
     * @param {Window} targetWindow - パレット
     * @param {Panel} matchesPanel - 右のパネル
     * @param {ListBox} matchList - 一覧
     * @returns {void}
     */
    function fitMatchListHeight(targetWindow, matchesPanel, matchList) {
        targetWindow.layout.layout(true);
        var panelHeight = matchesPanel.size.height;
        var usedHeight = PANEL_MARGINS[1] + PANEL_MARGINS[3] + matchesPanel.spacing * (matchesPanel.children.length - 1);
        for (var i = 0; i < matchesPanel.children.length; i++) usedHeight += matchesPanel.children[i].size.height;
        var extraHeight = panelHeight - usedHeight;
        if (extraHeight <= 0) return;
        setListHeight(matchList.size.height + extraHeight);
        targetWindow.layout.layout(true);
        /* パネルが求める高さには見出しの分が足されるので、伸ばしすぎてカラムごと高くなった分を戻す
           / the panel asks for extra height for its title, so take back what made the whole column taller */
        var overshoot = matchesPanel.size.height - panelHeight;
        if (overshoot > 0) {
            setListHeight(matchList.size.height - overshoot);
            targetWindow.layout.layout(true);
        }

        /**
         * 一覧の高さを変える（一度レイアウトしたあとは preferredSize だけでは効かないので size も入れる）
         * @param {number} listHeight - 高さ（px）
         * @returns {void}
         */
        function setListHeight(listHeight) {
            matchList.preferredSize.height = listHeight;
            matchList.size = [matchList.size.width, listHeight];
        }
    }


    /**
     * 入力欄の横に置く × のボタン（onDraw で描く）を足す。押すと onClick を呼ぶ
     * @param {Group} parentGroup - 足す先
     * @returns {Group} ボタン
     */
    function addClearButton(parentGroup) {
        var clearButton = parentGroup.add("group");
        clearButton.preferredSize = CLEAR_BUTTON_SIZE;
        clearButton.minimumSize = CLEAR_BUTTON_SIZE;
        clearButton.maximumSize = CLEAR_BUTTON_SIZE;
        clearButton.onDraw = function () {
            drawClearMark(this.graphics, CLEAR_BUTTON_SIZE[0], CLEAR_BUTTON_SIZE[1]);
        };
        /* フォーカスを入力欄へ移すので、押したときではなく離したときに呼ぶ（押したときに移しても離すと外れる）
           / fire on mouseup, since focus moved on mousedown is lost when the button is released */
        clearButton.addEventListener("mouseup", function () {
            if (typeof clearButton.onClick === "function") clearButton.onClick();
        });
        return clearButton;
    }

    /**
     * 丸の中に × を描く（入力欄を空にするボタンの絵）
     * @param {ScriptUIGraphics} buttonGraphics - 描く先の graphics
     * @param {number} buttonWidth - 幅
     * @param {number} buttonHeight - 高さ
     * @returns {void}
     */
    function drawClearMark(buttonGraphics, buttonWidth, buttonHeight) {
        var diameter = Math.min(buttonWidth, buttonHeight) - 2;
        var left = (buttonWidth - diameter) / 2;
        var top = (buttonHeight - diameter) / 2;
        buttonGraphics.newPath();
        buttonGraphics.ellipsePath(left, top, diameter, diameter);
        buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, CLEAR_BUTTON_CIRCLE_COLOR));
        /* × は丸の中央に、直径の 3 割の幅で / the × sits in the middle, 30% of the diameter on each side */
        var centerX = left + diameter / 2;
        var centerY = top + diameter / 2;
        var arm = diameter * 0.18;
        var markPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, CLEAR_BUTTON_MARK_COLOR, 1.5);
        buttonGraphics.newPath();
        buttonGraphics.moveTo(centerX - arm, centerY - arm);
        buttonGraphics.lineTo(centerX + arm, centerY + arm);
        buttonGraphics.moveTo(centerX + arm, centerY - arm);
        buttonGraphics.lineTo(centerX - arm, centerY + arm);
        buttonGraphics.strokePath(markPen);
    }

    /**
     * スライダーの行（項目名・スライダー・今の値）を足す
     * @param {Panel} parentPanel - 足す先
     * @param {Object} labelEntry - 項目名の { ja, en }
     * @param {Object} tooltipEntry - ツールチップの { ja, en }
     * @param {number[]} sliderRange - [初期値, 最小, 最大]
     * @param {string} valueLabel - 右に出す今の値
     * @param {number} bottomMargin - 行の下の余白（px）
     * @returns {{slider: Slider, valueText: StaticText}} スライダーと値の表示
     */
    function addSliderRow(parentPanel, labelEntry, tooltipEntry, sliderRange, valueLabel, bottomMargin) {
        var sliderRowGroup = parentPanel.add("group");
        setupRow(sliderRowGroup, "fill", 6);
        sliderRowGroup.margins = [0, 0, 0, bottomMargin];
        sliderRowGroup.add("statictext", undefined, labelText(labelEntry)).preferredSize.width = SLIDER_LABEL_WIDTH;
        var slider = sliderRowGroup.add("slider", undefined, sliderRange[0], sliderRange[1], sliderRange[2]);
        slider.alignment = ["fill", "center"];
        slider.helpTip = getLabel(tooltipEntry);
        var valueText = sliderRowGroup.add("statictext", undefined, valueLabel);
        valueText.preferredSize.width = SLIDER_VALUE_WIDTH;
        return { slider: slider, valueText: valueText };
    }

    /**
     * チェックボックスを足し、ツールチップと初期値を入れる
     * @param {Panel} parentPanel - 足す先
     * @param {Object} labelEntry - 文言の { ja, en }
     * @param {Object|null} tooltipEntry - ツールチップの { ja, en }（無ければ null）
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} チェックボックス
     */
    function addOptionCheckbox(parentPanel, labelEntry, tooltipEntry, initialValue) {
        var optionCheckbox = parentPanel.add("checkbox", undefined, getLabel(labelEntry));
        if (tooltipEntry) optionCheckbox.helpTip = getLabel(tooltipEntry);
        optionCheckbox.value = initialValue;
        return optionCheckbox;
    }

    /**
     * ［キーワード］パネルを作る（各行に統一先のラジオ・キーワード欄・×・件数、下に合計とオプション）
     * @param {Group} parentColumn - 足す先（左のカラム）
     * @param {HelperSettings} settings - 初期値
     * @returns {Object} パネルの部品（targetRadios / keywordInputs / clearButtons / countTexts / totalText / excludeContainedCheckbox）
     */
    function buildKeywordsPanel(parentColumn, settings) {
        var keywordsPanel = parentColumn.add("panel", undefined, getLabel(LABELS.panel.keywords));
        setupPanel(keywordsPanel, 6);
        var keywordsUi = { targetRadios: [], keywordInputs: [], clearButtons: [], countTexts: [] };
        for (var i = 0; i < KEYWORD_FIELD_COUNT; i++) {
            var keywordRowGroup = keywordsPanel.add("group");
            setupRow(keywordRowGroup, "fill", 6);
            var targetRadio = keywordRowGroup.add("radiobutton", undefined, "");
            targetRadio.helpTip = getLabel(LABELS.tooltip.unifyTarget);
            var keywordInput = keywordRowGroup.add("edittext", undefined, settings.keywords[i]);
            keywordInput.preferredSize.width = KEYWORD_FIELD_WIDTH;
            keywordInput.alignment = ["fill", "center"];
            keywordInput.helpTip = getLabel(LABELS.tooltip.keyword);
            var clearButton = addClearButton(keywordRowGroup);
            clearButton.helpTip = getLabel(LABELS.tooltip.clearKeyword);
            var countText = keywordRowGroup.add("statictext", undefined, getLabel(LABELS.status.initial));
            countText.preferredSize.width = COUNT_TEXT_WIDTH;
            countText.justify = "right";
            keywordsUi.targetRadios.push(targetRadio);
            keywordsUi.keywordInputs.push(keywordInput);
            keywordsUi.clearButtons.push(clearButton);
            keywordsUi.countTexts.push(countText);
        }

        /* 合計 / Total */
        var totalRowGroup = keywordsPanel.add("group");
        setupRow(totalRowGroup, "fill", 6);
        totalRowGroup.margins = [0, TOTAL_ROW_TOP_MARGIN, 0, TOTAL_ROW_BOTTOM_MARGIN];
        totalRowGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.total));
        keywordsUi.totalText = totalRowGroup.add("statictext", undefined, getLabel(LABELS.status.initial));
        keywordsUi.totalText.alignment = ["fill", "center"];
        keywordsUi.excludeContainedCheckbox = addOptionCheckbox(keywordsPanel, LABELS.checkbox.excludeContained, LABELS.tooltip.excludeContained, settings.excludeContained);
        return keywordsUi;
    }

    /**
     * ［対象］パネルを作る（ラジオは横並び）
     * @param {Group} parentColumn - 足す先（左のカラム）
     * @param {HelperSettings} settings - 初期値
     * @returns {{story: RadioButton, document: RadioButton}} 対象のラジオ
     */
    function buildScopePanel(parentColumn, settings) {
        var scopePanel = parentColumn.add("panel", undefined, getLabel(LABELS.panel.scope));
        setupPanel(scopePanel);
        scopePanel.orientation = "row";
        scopePanel.alignChildren = ["left", "center"];
        var scopeRadios = {
            story: scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.story)),
            document: scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.document))
        };
        scopeRadios.story.helpTip = getLabel(LABELS.tooltip.story);
        scopeRadios.document.helpTip = getLabel(LABELS.tooltip.document);
        scopeRadios[scopeRadios[settings.scope] ? settings.scope : "document"].value = true;
        return scopeRadios;
    }

    /**
     * ［オプション］パネルを作る
     * @param {Group} parentColumn - 足す先（左のカラム）
     * @param {HelperSettings} settings - 初期値
     * @returns {{caseSensitive: Checkbox, wholeWord: Checkbox, masterPages: Checkbox, excludeToc: Checkbox}} オプションのチェックボックス
     */
    function buildOptionsPanel(parentColumn, settings) {
        var optionsPanel = parentColumn.add("panel", undefined, getLabel(LABELS.panel.options));
        setupPanel(optionsPanel, 6);
        return {
            caseSensitive: addOptionCheckbox(optionsPanel, LABELS.checkbox.caseSensitive, null, settings.caseSensitive),
            wholeWord: addOptionCheckbox(optionsPanel, LABELS.checkbox.wholeWord, LABELS.tooltip.wholeWord, settings.wholeWord),
            masterPages: addOptionCheckbox(optionsPanel, LABELS.checkbox.masterPages, LABELS.tooltip.masterPages, settings.masterPages),
            excludeToc: addOptionCheckbox(optionsPanel, LABELS.checkbox.excludeToc, LABELS.tooltip.excludeToc, settings.excludeToc)
        };
    }

    /**
     * ［揺れ候補］パネルを作る（上に拾うボタンと絞り込み、一覧、下に［無視］［一覧をコピー］）
     * @param {Group} parentRow - 足す先（カラムの行）
     * @param {HelperSettings} settings - 初期値
     * @returns {Object} パネルの部品
     */
    function buildVariantsPanel(parentRow, settings) {
        var variantsPanel = parentRow.add("panel", undefined, getLabel(LABELS.panel.variants));
        setupPanel(variantsPanel, 6);
        variantsPanel.alignment = ["left", "fill"];
        var variantsUi = { panel: variantsPanel };

        variantsUi.btnCollect = variantsPanel.add("button", undefined, getLabel(LABELS.button.collectVariants));
        variantsUi.btnCollect.helpTip = getLabel(LABELS.tooltip.collectVariants);

        var filterRowGroup = variantsPanel.add("group");
        setupRow(filterRowGroup, "fill", 6);
        var kindLabels = [];
        for (var i = 0; i < VARIANT_KIND_KEYS.length; i++) kindLabels.push(getLabel(LABELS.variantKind[VARIANT_KIND_KEYS[i]]));
        variantsUi.filterList = filterRowGroup.add("dropdownlist", undefined, kindLabels);
        variantsUi.filterList.selection = Math.min(settings.variantFilter, kindLabels.length - 1);
        variantsUi.filterList.helpTip = getLabel(LABELS.tooltip.variantFilter);
        variantsUi.sortList = filterRowGroup.add("dropdownlist", undefined, [getLabel(LABELS.variantSort.frequent), getLabel(LABELS.variantSort.rare)]);
        variantsUi.sortList.selection = Math.min(settings.variantSort, 1);
        variantsUi.sortList.helpTip = getLabel(LABELS.tooltip.variantSort);

        variantsUi.groupList = variantsPanel.add("listbox", undefined, []);
        variantsUi.groupList.preferredSize = [VARIANT_LIST_WIDTH, VARIANT_LIST_MIN_HEIGHT];
        variantsUi.groupList.alignment = ["fill", "fill"];
        variantsUi.groupList.helpTip = getLabel(LABELS.tooltip.variantList);

        variantsUi.noteText = variantsPanel.add("statictext", undefined, " ");
        variantsUi.noteText.alignment = ["fill", "bottom"];

        var footerRowGroup = variantsPanel.add("group");
        setupRow(footerRowGroup, ["fill", "bottom"], 6);
        variantsUi.btnIgnore = footerRowGroup.add("button", undefined, getLabel(LABELS.button.ignore));
        variantsUi.btnIgnore.helpTip = getLabel(LABELS.tooltip.ignore);
        variantsUi.btnCopy = footerRowGroup.add("button", undefined, getLabel(LABELS.button.copyTable));
        variantsUi.btnCopy.helpTip = getLabel(LABELS.tooltip.copyTable);
        return variantsUi;
    }

    /**
     * ［該当箇所］パネルを作る（上にスライダーと表示のオプション、一覧、メモと［除外する］、いちばん下に1か所ずつの置換）
     * @param {Group} parentRow - 足す先（2カラムの行）
     * @param {HelperSettings} settings - 初期値
     * @returns {Object} パネルの部品
     */
    function buildMatchesPanel(parentRow, settings) {
        var matchesPanel = parentRow.add("panel", undefined, getLabel(LABELS.panel.matches));
        setupPanel(matchesPanel, 6);
        matchesPanel.alignment = ["fill", "fill"];
        var matchesUi = { panel: matchesPanel };

        var initialContextLength = clampContextLength(settings.contextLength);
        var contextRow = addSliderRow(matchesPanel, LABELS.fieldLabel.contextLength, LABELS.tooltip.contextLength,
            [initialContextLength, 0, MAX_CONTEXT_LENGTH], String(initialContextLength), CONTEXT_ROW_BOTTOM_MARGIN);
        matchesUi.contextLengthSlider = contextRow.slider;
        matchesUi.contextLengthText = contextRow.valueText;
        var initialZoomPercent = snapZoomPercent(settings.zoomPercent);
        var zoomRow = addSliderRow(matchesPanel, LABELS.fieldLabel.zoom, LABELS.tooltip.zoom,
            [initialZoomPercent, MIN_ZOOM_PERCENT, MAX_ZOOM_PERCENT], initialZoomPercent + "%", ZOOM_ROW_BOTTOM_MARGIN);
        matchesUi.zoomSlider = zoomRow.slider;
        matchesUi.zoomText = zoomRow.valueText;
        matchesUi.bracketMatchCheckbox = addOptionCheckbox(matchesPanel, LABELS.checkbox.bracketMatch, LABELS.tooltip.bracketMatch, settings.bracketMatch);
        matchesUi.withinParagraphCheckbox = addOptionCheckbox(matchesPanel, LABELS.checkbox.withinParagraph, LABELS.tooltip.withinParagraph, settings.withinParagraph);

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
        matchesUi.matchList = matchList;

        /* 一覧の下：メモと［除外する］ / Below the list: notes and Exclude */
        var listFooterGroup = matchesPanel.add("group");
        setupRow(listFooterGroup, ["fill", "bottom"], 6);
        /* 幅が狭いので2行まで折り返す / The width is narrow, so the note wraps to two lines */
        matchesUi.listNoteText = listFooterGroup.add("statictext", undefined, " ", { multiline: true });
        matchesUi.listNoteText.preferredSize = [MATCH_LIST_WIDTH - EXCLUDE_BUTTON_WIDTH - 6, LIST_NOTE_HEIGHT];
        matchesUi.listNoteText.alignment = ["fill", "center"];
        matchesUi.btnExclude = listFooterGroup.add("button", undefined, getLabel(LABELS.button.exclude));
        matchesUi.btnExclude.preferredSize.width = EXCLUDE_BUTTON_WIDTH;
        matchesUi.btnExclude.helpTip = getLabel(LABELS.tooltip.exclude);

        /* 1か所ずつ置換する（一覧で選んだ文字を直して［置換］）/ Change one match at a time (edit the picked text, then Change) */
        var changeRowGroup = matchesPanel.add("group");
        setupRow(changeRowGroup, ["fill", "bottom"], 6);
        changeRowGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.changeTo));
        matchesUi.changeToInput = changeRowGroup.add("edittext", undefined, "");
        matchesUi.changeToInput.preferredSize.width = CHANGE_TO_FIELD_WIDTH;
        matchesUi.changeToInput.alignment = ["fill", "center"];
        matchesUi.changeToInput.helpTip = getLabel(LABELS.tooltip.changeTo);
        matchesUi.btnChange = changeRowGroup.add("button", undefined, getLabel(LABELS.button.change));
        matchesUi.btnChange.preferredSize.width = CHANGE_BUTTON_WIDTH;
        matchesUi.btnChange.helpTip = getLabel(LABELS.tooltip.change);
        return matchesUi;
    }

    /**
     * パレットを組み立てる（左に揺れ候補、中央に設定、右に該当箇所の一覧と置き換え）
     * @param {HelperSettings} settings - 初期値
     * @returns {Window} パレット
     */
    function buildPalette(settings) {
        var helperPalette = new Window("palette", getLabel(LABELS.palette.title)); /* バージョンは左下に出す / the version is shown at the bottom left */
        setupWindow(helperPalette);

        /* 3カラム（高さをそろえる）/ Three columns of the same height */
        var columnsGroup = helperPalette.add("group");
        setupRow(columnsGroup, ["fill", "fill"], COLUMN_SPACING);
        columnsGroup.alignChildren = ["fill", "fill"];
        var variantsUi = buildVariantsPanel(columnsGroup, settings);
        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];
        leftColumn.spacing = WINDOW_SPACING;

        var keywordsUi = buildKeywordsPanel(leftColumn, settings);
        var targetRadios = keywordsUi.targetRadios;
        var keywordInputs = keywordsUi.keywordInputs;
        var clearButtons = keywordsUi.clearButtons;
        var countTexts = keywordsUi.countTexts;
        var totalText = keywordsUi.totalText;
        var excludeContainedCheckbox = keywordsUi.excludeContainedCheckbox;
        var scopeRadios = buildScopePanel(leftColumn, settings);
        var optionCheckboxes = buildOptionsPanel(leftColumn, settings);
        var matchesUi = buildMatchesPanel(columnsGroup, settings);
        var contextLengthSlider = matchesUi.contextLengthSlider;
        var contextLengthText = matchesUi.contextLengthText;
        var zoomSlider = matchesUi.zoomSlider;
        var zoomText = matchesUi.zoomText;
        var bracketMatchCheckbox = matchesUi.bracketMatchCheckbox;
        var withinParagraphCheckbox = matchesUi.withinParagraphCheckbox;
        var matchList = matchesUi.matchList;
        var listNoteText = matchesUi.listNoteText;
        var btnExclude = matchesUi.btnExclude;
        var changeToInput = matchesUi.changeToInput;
        var btnChange = matchesUi.btnChange;
        var lastActiveIndex = -1;

        /* ボタン：左に［選択文字を取り込む］［すべてを統一］、右に［閉じる］［検索］ / Buttons: Load Selection and Unify All on the left, Close and Find on the right */
        var buttonRow = addButtonRow(helperPalette);
        var btnLoadSelection = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.loadSelection));
        btnLoadSelection.helpTip = getLabel(LABELS.tooltip.loadSelection);
        var btnUnify = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.unify));
        btnUnify.helpTip = getLabel(LABELS.tooltip.unify);
        var btnClose = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.close));
        var btnFind = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.find));
        btnFind.helpTip = getLabel(LABELS.tooltip.find);
        alignRightOnlyButtonRow(buttonRow);
        addVersionFooter(helperPalette);

        /* 最後に検索した結果（前後の文字数などを変えたら、検索し直さずに一覧だけ作り直す）/ Last results; display-only changes rebuild just the list */
        var lastMatchesByKeyword = null;
        var lastKeywords = [];
        var lastScope = "document";
        /* 検索したあとにキーワードや設定を変えたら、検索し直すまで置き換えさせない / After keywords or settings change, block replacing until found again */
        var isResultStale = false;
        /* 除外した位置（「入れ物#開始位置」→ 見つけたときの文字）。検索し直しても残し、置き換えでずれた分は付け直す
           / Excluded matches ("container#start" -> matched text); kept across searches and shifted after replacing */
        var excludedMatches = {};
        /* 揺れ候補の組。一覧の行と同じ並びの組を visibleGroups に控える / Variant groups; visibleGroups lines up with the list rows */
        var variantDictionary = null;
        var variantGroups = [];
        var visibleGroups = [];
        var variantRules = settings.variantRules;
        var ignoredGroups = settings.ignoredGroups;
        var hasCollectedVariants = false;
        /* 一覧を作り直している間の onChange を無視する / Ignores onChange while the list is rebuilt */
        var isRenderingVariants = false;
        /* 語句欄に入っている組（統一したら統一先を覚える）/ The group in the fields; its target is remembered on unifying */
        var pickedGroupId = null;

        /**
         * 除外の控えのキーを返す
         * @param {MatchRange} matchRange - 見つかった位置
         * @returns {string} 「入れ物の指定子#開始位置」
         */
        function getMatchKey(matchRange) {
            return matchRange.container + "#" + matchRange.start;
        }

        /**
         * 除外したものかを返す（同じ位置でも文字が変わっていれば除外していないものとみなす）
         * @param {MatchRange} matchRange - 見つかった位置
         * @returns {boolean} 除外したものなら true
         */
        function isMatchExcluded(matchRange) {
            return excludedMatches[getMatchKey(matchRange)] === matchRange.matchedText;
        }

        /**
         * 除外する／しないを切り替える
         * @param {MatchRange} matchRange - 見つかった位置
         * @param {boolean} shouldExclude - 除外するなら true
         * @returns {void}
         */
        function setMatchExcluded(matchRange, shouldExclude) {
            if (shouldExclude) excludedMatches[getMatchKey(matchRange)] = matchRange.matchedText;
            else delete excludedMatches[getMatchKey(matchRange)];
        }

        /**
         * 置き換えで文字数が変わった分、同じ入れ物の後ろにある除外の位置をずらす
         * @param {Array<MatchRange>} replacedRanges - 置き換えた位置
         * @param {number} replacementLength - 置き換えた文字の長さ
         * @returns {void}
         */
        function shiftExcludedMatches(replacedRanges, replacementLength) {
            var shiftedMatches = {};
            for (var matchKey in excludedMatches) {
                if (!excludedMatches.hasOwnProperty(matchKey)) continue;
                var separatorIndex = matchKey.lastIndexOf("#");
                var containerKey = matchKey.substring(0, separatorIndex);
                var start = Number(matchKey.substring(separatorIndex + 1));
                var delta = 0;
                for (var j = 0; j < replacedRanges.length; j++) {
                    if (replacedRanges[j].container === containerKey && replacedRanges[j].end <= start) {
                        delta += replacementLength - (replacedRanges[j].end - replacedRanges[j].start);
                    }
                }
                shiftedMatches[containerKey + "#" + (start + delta)] = excludedMatches[matchKey];
            }
            excludedMatches = shiftedMatches;
        }

        /**
         * 前後の文字数の欄の値を返す
         * @returns {number} 前後の文字数（0〜MAX_CONTEXT_LENGTH の整数。スライダーの値を丸める）
         */
        function readContextLength() {
            return clampContextLength(contextLengthSlider.value);
        }

        /**
         * 表示倍率の欄の値を返す
         * @returns {number} 表示倍率（％。スライダーの値を ZOOM_STEP_PERCENT 単位に丸める）
         */
        function readZoomPercent() {
            return snapZoomPercent(zoomSlider.value);
        }

        /**
         * 統一先に選んだ欄の番号を返す
         * @returns {number} 欄の番号。選んでいなければ -1
         */
        function getTargetIndex() {
            for (var j = 0; j < targetRadios.length; j++) {
                if (targetRadios[j].value) return j;
            }
            return -1;
        }

        /**
         * 画面の値を設定にまとめる
         * @returns {HelperSettings} 設定
         */
        function readSettings() {
            var keywords = [];
            for (var j = 0; j < keywordInputs.length; j++) keywords.push(keywordInputs[j].text);
            return {
                keywords: keywords,
                scope: scopeRadios.story.value ? "story" : "document",
                caseSensitive: optionCheckboxes.caseSensitive.value,
                wholeWord: optionCheckboxes.wholeWord.value,
                masterPages: optionCheckboxes.masterPages.value,
                excludeToc: optionCheckboxes.excludeToc.value,
                excludeContained: excludeContainedCheckbox.value,
                contextLength: readContextLength(),
                bracketMatch: bracketMatchCheckbox.value,
                withinParagraph: withinParagraphCheckbox.value,
                variantRules: variantRules,
                ignoredGroups: ignoredGroups,
                variantFilter: variantsUi.filterList.selection ? variantsUi.filterList.selection.index : 0,
                variantSort: variantsUi.sortList.selection ? variantsUi.sortList.selection.index : 0,
                zoomPercent: readZoomPercent(),
                targetIndex: getTargetIndex()
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
         * 一覧の行が［すべてを統一］で置き換える対象かを返す（統一先と異なるキーワードの行）
         * @param {ListItem} listItem - 一覧の行
         * @returns {boolean} 置き換える対象なら true
         */
        function isUnifyRow(listItem) {
            var targetIndex = getTargetIndex();
            if (targetIndex < 0 || lastKeywords[targetIndex] === "") return false;
            return listItem.keywordIndex !== targetIndex && lastKeywords[listItem.keywordIndex] !== lastKeywords[targetIndex];
        }

        /**
         * ボタンなどを押せるかどうかをそろえる
         * @returns {void}
         */
        function updateEnabled() {
            optionCheckboxes.masterPages.enabled = scopeRadios.document.value;
            btnFind.enabled = hasAnyKeyword();
            var canReplace = !!lastMatchesByKeyword && !isResultStale;
            var targetIndex = getTargetIndex();
            btnUnify.enabled = canReplace && targetIndex >= 0 && lastKeywords[targetIndex] !== "";
            var selectedItem = matchList.selection;
            var hasSelectedMatch = !!(selectedItem && selectedItem.matchRange);
            btnChange.enabled = canReplace && hasSelectedMatch;
            btnExclude.enabled = hasSelectedMatch;
            variantsUi.btnIgnore.enabled = !!variantsUi.groupList.selection;
            variantsUi.btnCopy.enabled = visibleGroups.length > 0;
            btnExclude.text = (hasSelectedMatch && isMatchExcluded(selectedItem.matchRange))
                ? getLabel(LABELS.button.include)
                : getLabel(LABELS.button.exclude);
        }

        /**
         * 件数（除外した分を除く）を各欄の右と合計の行に出す
         * @returns {void}
         */
        function updateCountTexts() {
            if (!lastMatchesByKeyword) return;
            var totalCount = 0;
            var excludedCount = 0;
            for (var k = 0; k < countTexts.length; k++) {
                var matches = lastMatchesByKeyword[k] || [];
                var keptCount = 0;
                for (var matchIndex = 0; matchIndex < matches.length; matchIndex++) {
                    if (isMatchExcluded(matches[matchIndex])) excludedCount++;
                    else keptCount++;
                }
                totalCount += keptCount;
                countTexts[k].text = (lastKeywords[k] !== "") ? getLabel(LABELS.status.count, { count: keptCount }) : getLabel(LABELS.status.initial);
            }
            var scopeLabel = getLabel(LABELS.radio[lastScope]);
            totalText.text = (excludedCount > 0)
                ? getLabel(LABELS.status.totalWithExcluded, { count: totalCount, scope: scopeLabel, excluded: excludedCount })
                : getLabel(LABELS.status.total, { count: totalCount, scope: scopeLabel });
        }

        /**
         * 一覧の行の前後の文字列に「→ 統一先」を付ける（統一する行で、除外していないものだけ）
         * @param {ListItem} listItem - 一覧の行（baseContext に前後の文字列を持つ）
         * @returns {void}
         */
        function refreshListItem(listItem) {
            var isExcluded = isMatchExcluded(listItem.matchRange);
            listItem.checked = !isExcluded;
            var contextText = listItem.baseContext;
            if (!isExcluded && isUnifyRow(listItem)) contextText += " → " + lastKeywords[getTargetIndex()];
            listItem.subItems[0].text = contextText;
        }

        /**
         * 一覧で選んだ行の除外を切り替える
         * @returns {void}
         */
        function toggleSelectedExcluded() {
            var listItem = matchList.selection;
            if (!listItem || !listItem.matchRange) return;
            setMatchExcluded(listItem.matchRange, !isMatchExcluded(listItem.matchRange));
            refreshListItem(listItem);
            updateCountTexts();
            updateEnabled();
        }

        /**
         * 検索した結果が古くなったことにする（キーワード・対象・検索オプションを変えたとき）
         * @returns {void}
         */
        function markResultStale() {
            if (lastMatchesByKeyword && !isResultStale) {
                isResultStale = true;
                listNoteText.text = getLabel(LABELS.status.stale);
            }
            updateEnabled();
        }

        /**
         * 該当箇所の一覧を作り直す（キーワード欄の順、その中は見つかった順。先頭の MAX_LIST_ITEMS 件まで）。
         * チェックは「除外していない」印。統一先を選んでいれば、置き換える行に「→ 統一先」を付ける
         * @returns {void}
         */
        function renderMatchList() {
            matchList.removeAll();
            changeToInput.text = "";
            listNoteText.text = " ";
            if (!lastMatchesByKeyword) { updateEnabled(); return; }

            var contextLength = readContextLength();
            var containerLengths = {};
            var listedCount = 0;
            var shownCount = 0;
            for (var k = 0; k < lastMatchesByKeyword.length; k++) {
                var matches = lastMatchesByKeyword[k];
                for (var matchIndex = 0; matchIndex < matches.length; matchIndex++) {
                    var pageLabel, contextText;
                    /* パレット表示中に文字が消されたり、ドキュメントが閉じられたりする / the text or document may be gone while the palette is open */
                    try {
                        listedCount++;
                        if (shownCount >= MAX_LIST_ITEMS) continue;
                        pageLabel = getPageLabel(matches[matchIndex].text);
                        contextText = buildContextText(matches[matchIndex], contextLength, containerLengths, bracketMatchCheckbox.value, withinParagraphCheckbox.value);
                    } catch (e) {
                        if (shownCount >= MAX_LIST_ITEMS) continue;
                        pageLabel = "-";
                        contextText = getLabel(LABELS.status.unavailable);
                    }
                    var listItem = matchList.add("item", pageLabel);
                    listItem.matchRange = matches[matchIndex];
                    listItem.keywordIndex = k;
                    listItem.baseContext = contextText;
                    refreshListItem(listItem);
                    shownCount++;
                }
            }

            var notes = [];
            if (isResultStale) notes.push(getLabel(LABELS.status.stale));
            if (shownCount < listedCount) notes.push(getLabel(LABELS.status.truncated, { total: listedCount, shown: shownCount }));
            if (notes.length > 0) listNoteText.text = notes.join(" / ");
            updateEnabled();
        }

        /**
         * 検索して、件数を各欄の右と合計の行に、該当箇所を右の一覧に出す
         * @returns {void}
         */
        function findAndShow() {
            if (!hasAnyKeyword()) return;
            var currentSettings = readSettings();
            saveSettings(currentSettings);
            var countResult = runFind(currentSettings);
            lastMatchesByKeyword = countResult.matchesByKeyword;
            lastKeywords = currentSettings.keywords;
            lastScope = currentSettings.scope;
            isResultStale = false;
            if (lastMatchesByKeyword) {
                updateCountTexts();
            } else {
                for (var j = 0; j < countTexts.length; j++) countTexts[j].text = getLabel(LABELS.status.initial);
                totalText.text = countResult.message;
            }
            renderMatchList();
        }

        /**
         * 置き換えたあとに検索し直し、結果を一覧の下に出す
         * @param {{replaced: number, skipped: number, replacedRanges: Array<MatchRange>}} replaceResult - replaceMatches() の戻り値
         * @param {string} replacementText - 置き換えた文字
         * @param {string} doneMessage - 置き換えた件数の文言
         * @returns {void}
         */
        function refindAfterReplace(replaceResult, replacementText, doneMessage) {
            shiftExcludedMatches(replaceResult.replacedRanges, replacementText.length);
            findAndShow();
            var notes = [doneMessage];
            if (replaceResult.skipped > 0) notes.push(getLabel(LABELS.status.skipped, { count: replaceResult.skipped }));
            listNoteText.text = notes.join(" / ");
            if (hasCollectedVariants) collectVariants();
        }

        /**
         * 組の統一先にする表記を返す（前回の統一先、辞書の推奨、いちばん多い表記の順。要確認の組は選ばない）
         * @param {VariantGroup} group - 組
         * @returns {string|null} 表記。選ばなければ null
         */
        function getPreferredSpelling(group) {
            if (variantRules.hasOwnProperty(group.id)) return variantRules[group.id];
            if (group.recommended) return group.recommended;
            return (getVariantKind(group) === "review") ? null : group.spellings[0].text;
        }

        /**
         * 揺れ候補の一覧の1行を作る（統一先に ◉ を付ける）
         * @param {VariantGroup} group - 組
         * @returns {string} 行の文字列
         */
        function buildGroupRowText(group) {
            var preferred = getPreferredSpelling(group);
            var parts = [];
            for (var j = 0; j < group.spellings.length; j++) {
                var spelling = group.spellings[j];
                parts.push((spelling.text === preferred ? "◉" : "") + toDisplayText(spelling.text) + " " + spelling.count);
            }
            return parts.join(" / ");
        }

        /**
         * 揺れ候補の一覧を作り直す（無視している組を除き、種類で絞り、並び順に並べる）
         * @returns {void}
         */
        function renderVariantList() {
            var kindKey = VARIANT_KIND_KEYS[variantsUi.filterList.selection ? variantsUi.filterList.selection.index : 0];
            visibleGroups = [];
            for (var j = 0; j < variantGroups.length; j++) {
                if (ignoredGroups.hasOwnProperty(variantGroups[j].id)) continue;
                if (kindKey === "all" || getVariantKind(variantGroups[j]) === kindKey) visibleGroups.push(variantGroups[j]);
            }
            if (variantsUi.sortList.selection && variantsUi.sortList.selection.index === 1) visibleGroups.reverse();
            isRenderingVariants = true;
            variantsUi.groupList.removeAll();
            for (var k = 0; k < visibleGroups.length; k++) variantsUi.groupList.add("item", buildGroupRowText(visibleGroups[k]));
            isRenderingVariants = false;
            if (hasCollectedVariants) {
                variantsUi.noteText.text = (visibleGroups.length > 0) ? getLabel(LABELS.status.variants, { count: visibleGroups.length }) : getLabel(LABELS.status.noVariants);
            }
            updateEnabled();
        }

        /**
         * ［対象］の文字を読み取って揺れ候補を一覧にする
         * @returns {void}
         */
        function collectVariants() {
            var sourceText = collectVariantText(readSettings());
            if (sourceText === null) {
                variantsUi.noteText.text = getLabel(app.documents.length === 0 ? LABELS.status.noDocument : LABELS.status.noStory);
                return;
            }
            if (!variantDictionary) variantDictionary = loadVariantDictionary(VARIANT_DICTIONARY_LINES);
            variantGroups = findVariantGroups(sourceText, variantDictionary);
            hasCollectedVariants = true;
            renderVariantList();
        }

        /**
         * 組の表記をキーワード欄に入れて検索する。統一先の表記は使われていなくても1つ目の欄に入れる
         * @param {VariantGroup} group - 組
         * @returns {void}
         */
        function pickVariantGroup(group) {
            var preferred = getPreferredSpelling(group);
            var ordered = preferred ? [preferred] : [];
            var lowerTexts = {};
            var hasCaseVariants = false;
            for (var j = 0; j < group.spellings.length; j++) {
                if (group.spellings[j].text !== preferred) ordered.push(group.spellings[j].text);
            }
            for (var k = 0; k < keywordInputs.length; k++) {
                keywordInputs[k].text = (k < ordered.length) ? ordered[k] : "";
                targetRadios[k].value = !!preferred && k === 0;
                var lowerText = keywordInputs[k].text.toLowerCase();
                if (lowerText !== "" && lowerTexts.hasOwnProperty(lowerText)) hasCaseVariants = true;
                lowerTexts[lowerText] = true;
            }
            /* Apple と apple を分け、「サーバー」の中の「サーバ」と「仕事」の中の「事」を数えない
               / Tell Apple from apple, and do not count サーバ inside サーバー or 事 inside 仕事 */
            if (hasCaseVariants) optionCheckboxes.caseSensitive.value = true;
            excludeContainedCheckbox.value = true;
            optionCheckboxes.wholeWord.value = true;
            lastActiveIndex = -1;
            pickedGroupId = group.id;
            findAndShow();
        }

        /**
         * 表示している組を表（タブ区切り）にしてコピーする
         * @returns {void}
         */
        function copyVariantTable() {
            var lines = [[getLabel(LABELS.variantTable.group), getLabel(LABELS.variantTable.spelling), getLabel(LABELS.variantTable.count), getLabel(LABELS.variantTable.target)].join("\t")];
            for (var j = 0; j < visibleGroups.length; j++) {
                var preferred = getPreferredSpelling(visibleGroups[j]);
                for (var k = 0; k < visibleGroups[j].spellings.length; k++) {
                    var spelling = visibleGroups[j].spellings[k];
                    lines.push([j + 1, spelling.text, spelling.count, spelling.text === preferred ? "◉" : ""].join("\t"));
                }
            }
            if (copyToClipboard(lines.join("\n") + "\n")) variantsUi.noteText.text = getLabel(LABELS.status.copied, { count: visibleGroups.length });
        }

        /**
         * ［すべてを統一］：統一先と異なるキーワードの行のうち、除外していないもの（チェックが付いたもの）を統一先に置き換える
         * @returns {void}
         */
        function unifyMatches() {
            var targetIndex = getTargetIndex();
            if (targetIndex < 0 || !lastMatchesByKeyword || isResultStale) return;
            var targetKeyword = lastKeywords[targetIndex];
            var replaceTargets = [];
            for (var k = 0; k < lastMatchesByKeyword.length; k++) {
                if (k === targetIndex || lastKeywords[k] === targetKeyword) continue;
                var matches = lastMatchesByKeyword[k];
                for (var matchIndex = 0; matchIndex < matches.length; matchIndex++) {
                    if (!isMatchExcluded(matches[matchIndex])) replaceTargets.push(matches[matchIndex]);
                }
            }
            if (replaceTargets.length === 0) {
                listNoteText.text = getLabel(LABELS.status.nothingToReplace);
                return;
            }
            if (!confirm(getLabel(LABELS.alert.confirmUnify, { count: replaceTargets.length, target: targetKeyword }))) return;

            if (pickedGroupId) variantRules[pickedGroupId] = targetKeyword;
            var replaceResult = replaceMatches(replaceTargets, targetKeyword, getLabel(LABELS.undo.unify));
            refindAfterReplace(replaceResult, targetKeyword, getLabel(LABELS.status.unified, { count: replaceResult.replaced, target: targetKeyword }));
        }

        /**
         * ［置換］：一覧で選んだ1か所を、［置換文字列］の文字に置き換える
         * @returns {void}
         */
        function changeSelectedMatch() {
            var listItem = matchList.selection;
            if (!listItem || !listItem.matchRange || isResultStale) return;
            var matchRange = listItem.matchRange;
            var replacementText = changeToInput.text;
            if (replacementText === matchRange.matchedText) return;

            var keywordIndex = listItem.keywordIndex;
            var offsetInKeyword = listItem.index - findFirstRowOfKeyword(keywordIndex);
            var replaceResult = replaceMatches([matchRange], replacementText, getLabel(LABELS.undo.change));
            var doneMessage = (replaceResult.replaced > 0)
                ? getLabel(LABELS.status.changed, { before: matchRange.matchedText, after: replacementText })
                : getLabel(LABELS.status.nothingToReplace);
            refindAfterReplace(replaceResult, replacementText, doneMessage);

            /* 続けて直せるように、同じキーワードの次の行（置き換えた行があった位置）を選ぶ / select the next row of the same keyword so you can continue */
            var firstRow = findFirstRowOfKeyword(keywordIndex);
            if (firstRow < 0) return;
            var nextRow = Math.min(firstRow + offsetInKeyword, matchList.items.length - 1);
            if (matchList.items[nextRow].keywordIndex !== keywordIndex) nextRow--;
            if (nextRow >= firstRow) matchList.selection = nextRow;
        }

        /**
         * 一覧で、指定したキーワードの最初の行の番号を返す
         * @param {number} keywordIndex - キーワード欄の番号
         * @returns {number} 行の番号。無ければ -1
         */
        function findFirstRowOfKeyword(keywordIndex) {
            for (var j = 0; j < matchList.items.length; j++) {
                if (matchList.items[j].keywordIndex === keywordIndex) return j;
            }
            return -1;
        }

        /**
         * 選択文字を取り込む欄を決める（最後に入力した欄、なければ最初の空欄、空欄も無ければ1つ目）
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
         * キーワード欄と統一先のラジオのイベントを付ける
         * @param {number} fieldIndex - 何番目の欄か
         * @returns {void}
         */
        function bindKeywordRow(fieldIndex) {
            var keywordInput = keywordInputs[fieldIndex];
            keywordInput.onChanging = function () {
                lastActiveIndex = fieldIndex;
                pickedGroupId = null;
                markResultStale();
            };
            keywordInput.onActivate = function () { lastActiveIndex = fieldIndex; };
            /* × で欄を空にして、その欄に入力できるようにする / the × clears the field and puts the cursor in it */
            clearButtons[fieldIndex].onClick = function () {
                /* option キーを押しながらなら、すべての欄を空にする / With the Option key, clear every field */
                var clearsAll = ScriptUI.environment.keyboardState.altKey;
                if (keywordInput.text === "" && !clearsAll) return;
                for (var j = 0; j < keywordInputs.length; j++) {
                    if (clearsAll || j === fieldIndex) keywordInputs[j].text = "";
                }
                lastActiveIndex = clearsAll ? 0 : fieldIndex;
                pickedGroupId = null;
                markResultStale();
                keywordInputs[lastActiveIndex].active = true;
            };
            /* キーワード欄の return で検索する / Return in a keyword field finds */
            keywordInput.addEventListener("keydown", function (event) {
                if (event.keyName !== "Enter") return;
                event.preventDefault();
                findAndShow();
            });
            /* ラジオは行ごとに別のグループにあるので、自分以外を外す（同じ親の中でしか排他にならない）
               / The radios sit in separate row groups, so clear the others by hand (radios are exclusive only within one parent) */
            targetRadios[fieldIndex].onClick = function () {
                for (var j = 0; j < targetRadios.length; j++) targetRadios[j].value = (j === fieldIndex);
                renderMatchList();
            };
        }
        for (var fieldIndex = 0; fieldIndex < keywordInputs.length; fieldIndex++) bindKeywordRow(fieldIndex);
        var initialTargetIndex = (typeof settings.targetIndex === "number") ? settings.targetIndex : -1;
        if (initialTargetIndex >= 0 && initialTargetIndex < targetRadios.length) targetRadios[initialTargetIndex].value = true;

        /* 動かしている間は数字だけ変え、離したときに一覧を作り直す（一覧の作り直しは件数が多いと重い）
           / Update only the number while dragging and rebuild the list on release (rebuilding is slow with many rows) */
        contextLengthSlider.onChanging = function () {
            contextLengthText.text = String(readContextLength());
        };
        contextLengthSlider.onChange = function () {
            var contextLength = readContextLength();
            contextLengthSlider.value = contextLength;
            contextLengthText.text = String(contextLength);
            renderMatchList();
        };
        bracketMatchCheckbox.onClick = renderMatchList;
        withinParagraphCheckbox.onClick = renderMatchList;
        /* 動かしている間もつまみを10%刻みの位置に置き直す / keep the knob on the 10% steps while dragging */
        zoomSlider.onChanging = function () {
            var zoomPercent = readZoomPercent();
            if (zoomSlider.value !== zoomPercent) zoomSlider.value = zoomPercent;
            zoomText.text = zoomPercent + "%";
        };
        zoomSlider.onChange = zoomSlider.onChanging;

        /* 一覧で選んだ文字を選択して表示し、［置換文字列］に入れる / Select and show the picked text, and put it in Change to */
        matchList.onChange = function () {
            var listItem = matchList.selection;
            updateEnabled();
            if (!listItem || !listItem.matchRange) return;
            changeToInput.text = listItem.matchRange.matchedText;
            try {
                var foundText = listItem.matchRange.text;
                if (!foundText.isValid) return;
                foundText.select();
                /* 表示倍率を先に変えてから、その文字の位置へ移動する / change the zoom first, then move to the text */
                setLayoutZoom(readZoomPercent());
                foundText.showText();
            } catch (e) { /* 表示できない場所（閉じたドキュメントなど）は何もしない / nothing to show (e.g. a closed document) */ }
        };
        /* ダブルクリックでも除外を切り替える（Mac はチェック欄を直接クリックできない）
           / Double-click also toggles exclusion (the check box cannot be clicked directly on Mac) */
        matchList.onDoubleClick = toggleSelectedExcluded;
        btnExclude.onClick = toggleSelectedExcluded;

        /* 検索の条件が変わったら、検索し直すまで置き換えさせない / changing the search conditions blocks replacing until the next Find */
        var conditionControls = [scopeRadios.story, scopeRadios.document, excludeContainedCheckbox,
            optionCheckboxes.caseSensitive, optionCheckboxes.wholeWord, optionCheckboxes.masterPages, optionCheckboxes.excludeToc];
        for (var controlIndex = 0; controlIndex < conditionControls.length; controlIndex++) conditionControls[controlIndex].onClick = markResultStale;
        btnFind.onClick = findAndShow;
        btnUnify.onClick = unifyMatches;
        btnChange.onClick = changeSelectedMatch;
        btnLoadSelection.onClick = function () {
            var selectedKeyword = getSelectedKeyword(getCurrentSelection());
            if (selectedKeyword === "") return;
            getInsertTargetInput().text = selectedKeyword;
            markResultStale();
        };
        variantsUi.btnCollect.onClick = collectVariants;
        variantsUi.filterList.onChange = renderVariantList;
        variantsUi.sortList.onChange = renderVariantList;
        variantsUi.groupList.onChange = function () {
            updateEnabled();
            if (isRenderingVariants || !variantsUi.groupList.selection) return;
            pickVariantGroup(visibleGroups[variantsUi.groupList.selection.index]);
        };
        variantsUi.btnIgnore.onClick = function () {
            if (!variantsUi.groupList.selection) return;
            ignoredGroups[visibleGroups[variantsUi.groupList.selection.index].id] = true;
            renderVariantList();
        };
        variantsUi.btnCopy.onClick = copyVariantTable;
        btnClose.onClick = function () { helperPalette.close(); };

        /* パレットは Esc で閉じないので、［閉じる］と同じにする / palettes do not close on Esc, so treat it as Close */
        helperPalette.addEventListener("keydown", function (event) {
            if (event.keyName !== "Escape") return;
            event.preventDefault();
            helperPalette.close();
        });

        /* 閉じるときは設定を控えて参照を外すだけ（DOM には触れない）/ On close, only keep the settings and drop the reference (no DOM access) */
        helperPalette.onClose = function () {
            saveSettings(readSettings());
            $.global[PALETTE_GLOBAL_KEY] = null;
        };

        helperPalette.onShow = function () {
            getInsertTargetInput().active = true;
        };

        updateEnabled();
        fitMatchListHeight(helperPalette, matchesUi.panel, matchList);
        fitMatchListHeight(helperPalette, variantsUi.panel, variantsUi.groupList);
        return helperPalette;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択した文字を、まだ入っていなければ最初の空欄に入れる
     * @param {HelperSettings} settings - 設定（keywords を書き換える）
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

        var helperPalette = buildPalette(settings);
        /* パレットは show() から即座に戻るので、参照を $.global に残す / palettes return from show() at once, so keep a reference */
        $.global[PALETTE_GLOBAL_KEY] = helperPalette;
        helperPalette.show();
    }

    main();

})();
