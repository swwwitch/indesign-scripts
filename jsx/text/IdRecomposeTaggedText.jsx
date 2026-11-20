#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdRecomposeTaggedText
Meta-Default-language=ja
Name=流し込んだタグ付きテキストの記述を書式に変換
Name[en]=Convert placed tagged text markup into formatting
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.0.0
Release-Date=2026-11-21
Tags=text,テキスト
X-com.swwwitch-Script-Name=IdRecomposeTaggedText
X-com.swwwitch-Updated=2026-11-21
X-com.swwwitch-README-JA=https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdRecomposeTaggedText.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdRecomposeTaggedText.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/indesign-scripts/blob/main/jsx/text/IdRecomposeTaggedText.jsx
Description-BEGIN
[ja]
テキストフレームにそのまま流し込まれたタグ付きテキストの記述を、タグ付きテキストとして読み込み直して書式に変換します。何も選択していないときは、ドキュメント内でタグの記述を含むストーリーをすべて変換します。
[en]
Re-imports tagged text markup that was placed as plain text, turning the tags into formatting. With nothing selected, every story in the document that contains tag markup is converted.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読んで SCRIPT_* に入れる（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; it is read into SCRIPT_* at run time */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME           = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名 / script name */
var SCRIPT_VERSION        = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : ""; /* バージョン / version */
var SCRIPT_AUTHOR         = getScriptMetaValue("Author");                       /* 作者 / author */
var SCRIPT_RELEASED       = getScriptMetaValue("Release-Date");                 /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED        = getScriptMetaValue("X-com.swwwitch-Updated");       /* 更新日 / last updated */
var SCRIPT_DESCRIPTION_JA = getScriptMetaValue("Name");                         /* README の見出し / README title (Japanese) */
var SCRIPT_README_JA      = getScriptMetaValue("X-com.swwwitch-README-JA");     /* README（日本語） */
var SCRIPT_README_EN      = getScriptMetaValue("X-com.swwwitch-README-EN");     /* README (English) */
var SCRIPT_ARTICLE_URL    = getScriptMetaValue("X-com.swwwitch-Article-URL");   /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL     = getScriptMetaValue("X-com.swwwitch-Source-URL");    /* 最新版のソース / latest source */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @discussion
 * 発想：id_recomposeTaggedText_a.jsx（z-、2012）。段落スタイル・文字スタイルのタグだけを自前で解釈していたのに対し、
 * こちらは文字列をタグ付きテキストのファイルに書き出して InDesign の読み込み機能で読み直すため、すべてのタグが効く。
 * Idea from id_recomposeTaggedText_a.jsx (z-, 2012), which parsed only paragraph and character style tags itself.
 * This script writes the text to a tagged text file and lets InDesign's importer read it back, so every tag works.
 */

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 書き出すタグ付きテキストのヘッダー。ASCII 以外の文字は <0xXXXX> で書く（UNICODE のヘッダーでは <0x000A> の強制改行が効かない）
       Header of the written tagged text. Non-ASCII characters are written as <0xXXXX> (a UNICODE header ignores <0x000A> forced line breaks) */
    var TAGGED_TEXT_HEADER = "<ASCII-MAC>";

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
        alert: {
            noDocument:    { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noTaggedText:  { ja: "タグ付きテキストの記述が見つかりませんでした。", en: "No tagged text markup was found." },
            convertedCount: { ja: "変換したテキスト", en: "Converted texts" },
            skippedObjects: { ja: "表・アンカー付きオブジェクト・脚注などを含むため変換しなかったテキスト", en: "Skipped (contains tables, anchored objects, footnotes, etc.)" },
            failedCount:   { ja: "変換できなかったテキスト（ロックなど）", en: "Failed (locked, etc.)" }
        },
        undo: {
            recompose: { ja: "タグ付きテキストを書式に変換", en: "Recompose Tagged Text" }
        }
    };

    // =========================================
    // 判定 / Detection
    // =========================================

    /* タグの記述：<名前:値>・<0x0000>・<SJIS-MAC> などのヘッダー・値のない表のタグ
       Tag markup: <Name:value>, <0x0000>, headers such as <SJIS-MAC>, and value-less table tags */
    var TAG_PATTERN = /<(?:[A-Za-z][A-Za-z0-9]*:|0x[0-9A-Fa-f]{4}>|[A-Z][A-Z0-9]*-(?:MAC|WIN)>|(?:tEnd|rStart|rEnd|cEnd|coStart|TableEnd|RowStart|RowEnd|CellEnd|ColStart)>)/;

    /* 先頭のエンコーディング指定（<SJIS-MAC> など）。書き出すときに付け直す / Leading encoding header such as <SJIS-MAC>; replaced on write */
    var ENCODING_HEADER_PATTERN = /^\s*<[A-Z][A-Z0-9]*-(?:MAC|WIN)>[ \t]*[\r\n]?/;

    /* 読み直すと失われる文字（表・アンカー付きオブジェクト・脚注・ページ番号などの特殊文字）。タブ・改行・「ここまでインデント」・右インデントタブは除く
       Characters lost on re-import (tables, anchored objects, footnotes, page numbers, …); tabs, breaks, indent-to-here and right indent tabs excluded */
    var OBJECT_CHAR_PATTERN = /[\u0000-\u0006\u000B\u000C\u000E-\u001F\uFFFC]/;

    /* 定義のタグ（<DefineParaStyle:…>・<Version:…> など）の書き出し / Start of a definition tag such as <DefineParaStyle:…> or <Version:…> */
    var DEFINITION_TAG_PATTERN = /^<(Define[A-Za-z]*|Version|FeatureSet|ColorTable):/;

    /**
     * 範囲のプロパティが配列で返ったときは先頭を取る（itemByRange で作った範囲は配列で返す）
     * @param {*} propertyValue - プロパティの値
     * @returns {*} 先頭の値、または値そのもの
     */
    function firstValue(propertyValue) {
        return (propertyValue instanceof Array) ? propertyValue[0] : propertyValue;
    }

    /**
     * 文字列にタグ付きテキストの記述があるか判定する
     * @param {string} sourceText - 調べる文字列
     * @returns {boolean} タグの記述があれば true
     */
    function hasTagMarkup(sourceText) {
        return TAG_PATTERN.test(sourceText);
    }

    // =========================================
    // 対象の収集 / Target collection
    // =========================================

    /**
     * @typedef {Object} RecomposeTarget
     * @property {Object} container - 書き換える範囲を持つ入れ物（Story または Cell）
     * @property {Text} [textRange] - 選択したテキスト範囲（入れ物全体のときは無し）
     */

    /**
     * 入れ物全体を対象として追加する（同じストーリー・セルは1度だけ）
     * @param {Array<RecomposeTarget>} targets - 対象の配列
     * @param {Object} seenIds - 追加済みの入れ物の ID
     * @param {Object} container - Story または Cell
     * @returns {void}
     */
    function addContainerTarget(targets, seenIds, container) {
        var containerKey = container.constructor.name + ":" + container.id;
        if (seenIds[containerKey]) return;
        seenIds[containerKey] = true;
        targets.push({ container: container });
    }

    /**
     * 選択されたオブジェクト1つから対象を集める（グループは中へたどる）
     * @param {Object} selectedItem - 選択されたオブジェクト
     * @param {Array<RecomposeTarget>} targets - 対象の配列
     * @param {Object} seenIds - 追加済みの入れ物の ID
     * @returns {void}
     */
    function collectFromItem(selectedItem, targets, seenIds) {
        var itemType = selectedItem.constructor.name;

        if (itemType === "TextFrame") {
            addContainerTarget(targets, seenIds, selectedItem.parentStory);
            return;
        }
        if (itemType === "Cell") {
            var selectedCells = selectedItem.cells.everyItem().getElements();
            for (var i = 0; i < selectedCells.length; i++) addContainerTarget(targets, seenIds, selectedCells[i]);
            return;
        }
        if (itemType === "Table") {
            var tableCells = selectedItem.cells.everyItem().getElements();
            for (var j = 0; j < tableCells.length; j++) addContainerTarget(targets, seenIds, tableCells[j]);
            return;
        }
        if (selectedItem.hasOwnProperty("insertionPoints") && selectedItem.hasOwnProperty("parentStory")) {
            /* カーソルだけなら入れ物（ストーリーまたはセル）全体、文字を選んでいればその範囲 / Whole container (story or cell) for a caret, the range for selected text */
            var container = firstValue(selectedItem.insertionPoints[0].parent);
            if (itemType === "InsertionPoint" || selectedItem.characters.length === 0) {
                addContainerTarget(targets, seenIds, container);
            } else {
                targets.push({ container: container, textRange: selectedItem.texts[0] });
            }
            return;
        }
        if (selectedItem.hasOwnProperty("pageItems")) {
            var childItems = selectedItem.pageItems.everyItem().getElements();
            for (var k = 0; k < childItems.length; k++) collectFromItem(childItems[k], targets, seenIds);
        }
    }

    /**
     * 選択から対象を集める。選択が無ければ、タグの記述を含むストーリーをすべて対象にする
     * @param {Document} targetDocument - 対象のドキュメント
     * @returns {Array<RecomposeTarget>} 対象の配列
     */
    function collectTargets(targetDocument) {
        var targets = [];
        var seenIds = {};
        var selectedItems = targetDocument.selection;

        if (selectedItems.length === 0) {
            var stories = targetDocument.stories.everyItem().getElements();
            for (var i = 0; i < stories.length; i++) {
                if (hasTagMarkup(stories[i].contents)) addContainerTarget(targets, seenIds, stories[i]);
            }
            return targets;
        }
        for (var j = 0; j < selectedItems.length; j++) collectFromItem(selectedItems[j], targets, seenIds);
        return targets;
    }

    // =========================================
    // タグ付きテキストの読み直し / Tagged text re-import
    // =========================================

    /**
     * 対象の今の文字列を返す
     * @param {RecomposeTarget} target - 対象
     * @returns {string} 文字列
     */
    function getTargetText(target) {
        var sourceText = target.textRange ? target.textRange.contents : target.container.texts[0].contents;
        return String(sourceText);
    }

    /**
     * 段落スタイルを ParaStyle タグに書く名前に変える（グループは「:」で区切り、名前の中の「\」「:」「<」「>」はエスケープ）
     * @param {ParagraphStyle} paragraphStyle - 段落スタイル
     * @returns {string} タグに書く名前（［段落スタイルなし］は空、［基本段落］は NormalParagraphStyle）
     */
    function paragraphStyleTagName(paragraphStyle) {
        var nameParts = [];
        var styleNode = paragraphStyle;
        while (styleNode.constructor.name !== "Document") {
            nameParts.unshift(styleNode.name.replace(/\\/g, "\\\\").replace(/:/g, "\\:"));
            styleNode = styleNode.parent;
        }
        var rootStyles = styleNode.paragraphStyles;
        if (paragraphStyle.id === rootStyles[0].id) return "";
        if (paragraphStyle.id === rootStyles[1].id) return "NormalParagraphStyle";
        return nameParts.join(":").replace(/([\\:<>])/g, "\\$1");
    }

    /**
     * 読み込む文字の先頭に付けるタグを返す。
     * 書き換える位置の段落スタイルを指定し、タグの無い段落が［段落スタイルなし］にならないようにする。
     * 先頭に段落のタグが無いと、すぐ後の文字のタグ（<cSize:30> など）も効かない。
     * 本文が ParaStyle で始まるときは付けない（ParaStyle が続くと空の段落ができる）。先頭に定義があればその後ろに付ける
     * @param {RecomposeTarget} target - 対象（書き換える前）
     * @returns {string} ParaStyle タグ
     */
    function getLeadingTags(target) {
        var startPoint = target.textRange ? target.textRange.insertionPoints[0] : target.container.insertionPoints[0];
        return "<ParaStyle:" + paragraphStyleTagName(firstValue(startPoint.appliedParagraphStyle)) + ">";
    }

    /**
     * position の「<」で始まるタグの閉じ「>」の位置を返す（入れ子の <…> と「\」のエスケープを考慮）。
     * ExtendScript の正規表現は入れ子の量指定を誤評価するので、1文字ずつ数える
     * @param {string} text - 文字列
     * @param {number} position - タグの「<」の位置
     * @returns {number} 閉じ「>」の位置（閉じていなければ -1）
     */
    function findTagEnd(text, position) {
        var depth = 0;
        for (var i = position; i < text.length; i++) {
            var character = text.charAt(i);
            if (character === "\\") {
                i++;
            } else if (character === "<") {
                depth++;
            } else if (character === ">") {
                depth--;
                if (depth === 0) return i;
            }
        }
        return -1;
    }

    /**
     * 先頭に続く定義のタグと、その直後の改行までの長さを返す。
     * 段落のタグは定義の後ろに付ける（前に付けると空の段落ができる）
     * @param {string} bodyText - 本文
     * @returns {number} 定義の部分の長さ（定義が無ければ 0）
     */
    function measureLeadingDefinitions(bodyText) {
        var position = 0;
        var definitionEnd = 0;
        while (true) {
            while (position < bodyText.length && /\s/.test(bodyText.charAt(position))) position++;
            if (!DEFINITION_TAG_PATTERN.test(bodyText.substr(position, 40))) break;
            var tagEnd = findTagEnd(bodyText, position);
            if (tagEnd < 0) break;
            position = tagEnd + 1;
            definitionEnd = position;
        }
        if (definitionEnd === 0) return 0;
        while (definitionEnd < bodyText.length && /[ \t]/.test(bodyText.charAt(definitionEnd))) definitionEnd++;
        if (/[\r\n]/.test(bodyText.charAt(definitionEnd))) definitionEnd++;
        return definitionEnd;
    }

    /**
     * 文字より前に続くタグの中に ParaStyle があるか判定する
     * @param {string} contentText - 本文（定義を除く）
     * @returns {boolean} ParaStyle があれば true
     */
    function startsWithParaStyle(contentText) {
        var position = 0;
        while (contentText.charAt(position) === "<") {
            if (contentText.substr(position, 11) === "<ParaStyle:") return true;
            var tagEnd = findTagEnd(contentText, position);
            if (tagEnd < 0) return false;
            position = tagEnd + 1;
        }
        return false;
    }

    /**
     * InDesign の文字列をタグ付きテキストのファイルの中身に変える
     * @param {string} sourceText - 段落改行が CR、強制改行が LF の文字列
     * @param {string} leadingTags - 本文の先頭に付けるタグ
     * @returns {string} ヘッダー付き、段落改行は LF、ASCII 以外と制御文字（強制改行を含む）は <0xXXXX> の文字列
     */
    function buildTaggedText(sourceText, leadingTags) {
        var bodyText = sourceText.replace(ENCODING_HEADER_PATTERN, "");
        var definitionLength = measureLeadingDefinitions(bodyText);
        var contentText = bodyText.substr(definitionLength);
        if (!startsWithParaStyle(contentText)) contentText = leadingTags + contentText;
        bodyText = bodyText.substr(0, definitionLength) + contentText;
        var taggedParts = [];
        for (var i = 0; i < bodyText.length; i++) {
            var charCode = bodyText.charCodeAt(i);
            if (charCode === 13) taggedParts.push("\n");
            else if (charCode === 9 || (charCode >= 32 && charCode <= 126)) taggedParts.push(bodyText.charAt(i));
            else taggedParts.push("<0x" + ("000" + charCode.toString(16).toUpperCase()).slice(-4) + ">");
        }
        return TAGGED_TEXT_HEADER + "\n" + taggedParts.join("");
    }

    /**
     * タグ付きテキストを一時ファイルに書く（ASCII・LF。CR だとタグ付きテキストと認識されない）
     * @param {string} taggedText - ファイルの中身
     * @param {number} fileIndex - ファイル名の通し番号
     * @returns {File} 書いたファイル
     */
    function writeTaggedTextFile(taggedText, fileIndex) {
        var taggedFile = new File(Folder.temp.fsName + "/" + SCRIPT_NAME + "_" + new Date().getTime() + "_" + fileIndex + ".txt");
        taggedFile.encoding = "ASCII";
        taggedFile.lineFeed = "Unix";
        taggedFile.open("w");
        taggedFile.write(taggedText);
        taggedFile.close();
        return taggedFile;
    }

    /**
     * 対象の文字を消し、同じ位置にタグ付きテキストのファイルを読み込む。
     * 先に読み込んでから消すと、範囲が読み込んだ文字まで伸びて一緒に消えたり、先頭の段落スタイルが元のまま残ったりするので、消してから読み込む。
     * 読み込みに失敗したときは元の文字列（書式なし）を戻す
     * @param {RecomposeTarget} target - 対象
     * @param {File} taggedFile - タグ付きテキストのファイル
     * @param {string} sourceText - 元の文字列
     * @returns {void}
     */
    function replaceWithTaggedFile(target, taggedFile, sourceText) {
        var container = target.container;
        var insertIndex = 0;
        if (target.textRange) {
            insertIndex = firstValue(target.textRange.insertionPoints[0].index);
            target.textRange.remove();
        } else {
            container.texts[0].remove();
        }
        try {
            container.insertionPoints[insertIndex].place(taggedFile);
        } catch (e) {
            container.insertionPoints[insertIndex].contents = sourceText;
            throw e;
        }
    }

    /**
     * 対象を1つ変換する
     * @param {RecomposeTarget} target - 対象
     * @param {number} fileIndex - 一時ファイルの通し番号
     * @returns {string} 結果（"converted" / "notTagged" / "skipped" / "failed"）
     */
    function recomposeTarget(target, fileIndex) {
        var sourceText = getTargetText(target);
        if (!hasTagMarkup(sourceText)) return "notTagged";
        if (OBJECT_CHAR_PATTERN.test(sourceText)) return "skipped";

        var taggedFile = writeTaggedTextFile(buildTaggedText(sourceText, getLeadingTags(target)), fileIndex);
        try {
            replaceWithTaggedFile(target, taggedFile, sourceText);
            return "converted";
        } catch (e) {
            return "failed";
        } finally {
            taggedFile.remove();
        }
    }

    /**
     * 読み込みの設定を変換向けにし、元の値を返す（引用符を勝手に変えない・書式を残す・ダイアログボックスを出さない）。
     * 読み込みでダイアログボックスが出ると、非表示のドキュメントでは閉じられずに止まる
     * @returns {Object} 元の設定
     */
    function prepareImportPreferences() {
        var importPreferences = app.taggedTextImportPreferences;
        var savedPreferences = {
            useTypographersQuotes: importPreferences.useTypographersQuotes,
            removeTextFormatting:  importPreferences.removeTextFormatting,
            userInteractionLevel:  app.scriptPreferences.userInteractionLevel
        };
        importPreferences.useTypographersQuotes = false;
        importPreferences.removeTextFormatting = false;
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
        return savedPreferences;
    }

    /**
     * 読み込みの設定を元に戻す
     * @param {Object} savedPreferences - prepareImportPreferences が返した値
     * @returns {void}
     */
    function restoreImportPreferences(savedPreferences) {
        var importPreferences = app.taggedTextImportPreferences;
        importPreferences.useTypographersQuotes = savedPreferences.useTypographersQuotes;
        importPreferences.removeTextFormatting = savedPreferences.removeTextFormatting;
        app.scriptPreferences.userInteractionLevel = savedPreferences.userInteractionLevel;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 対象をすべて変換し、件数を返す
     * @param {Array<RecomposeTarget>} targets - 対象の配列
     * @returns {Object} 結果ごとの件数（converted / notTagged / skipped / failed）
     */
    function recomposeTargets(targets) {
        var resultCounts = { converted: 0, notTagged: 0, skipped: 0, failed: 0 };
        var savedPreferences = prepareImportPreferences();
        try {
            /* 後ろから処理し、同じ入れ物の前の範囲の位置をずらさない / Work backwards so earlier ranges in the same container keep their positions */
            for (var i = targets.length - 1; i >= 0; i--) {
                resultCounts[recomposeTarget(targets[i], i)]++;
            }
        } finally {
            restoreImportPreferences(savedPreferences);
        }
        return resultCounts;
    }

    /**
     * 結果を知らせる（すべて変換できたときは何も出さない）
     * @param {Object} resultCounts - recomposeTargets の戻り値
     * @returns {void}
     */
    function reportResult(resultCounts) {
        if (resultCounts.converted === 0 && resultCounts.skipped === 0 && resultCounts.failed === 0) {
            alert(getLabel(LABELS.alert.noTaggedText));
            return;
        }
        if (resultCounts.skipped === 0 && resultCounts.failed === 0) return;
        var reportLines = [labelValueText(LABELS.alert.convertedCount, resultCounts.converted)];
        if (resultCounts.skipped > 0) reportLines.push(labelValueText(LABELS.alert.skippedObjects, resultCounts.skipped));
        if (resultCounts.failed > 0) reportLines.push(labelValueText(LABELS.alert.failedCount, resultCounts.failed));
        alert(reportLines.join("\n"));
    }

    /**
     * 選択（無ければドキュメント全体）のタグ付きテキストの記述を書式に変換する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var targets = collectTargets(app.activeDocument);
        var resultCounts = { converted: 0, notTagged: 0, skipped: 0, failed: 0 };
        if (targets.length > 0) {
            app.doScript(function () {
                resultCounts = recomposeTargets(targets);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undo.recompose));
        }
        reportResult(resultCounts);
    }

    main();

})();
