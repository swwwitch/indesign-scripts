#target indesign

/*

### 概要

MS Word から取り込んだ段落スタイル名（Heading 1 / Normal / Quote など）を、対応する HTML 要素名（h1 / p / blockquote）へ一括でリネームします。
スタイルグループの中も再帰的にたどり、リネーム先の名前が同じ階層に既にある場合は連番を付けて重複を避けます。

詳細は README を参照してください。

### Overview

Renames paragraph styles imported from MS Word (Heading 1 / Normal / Quote and the like) to the matching HTML element names (h1 / p / blockquote) in one pass.
Style groups are traversed recursively, and a numeric suffix is appended when the target name is already taken in the same group.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdConvertWordStyleNamesToHtml"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-20";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdConvertWordStyleNamesToHtml.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdConvertWordStyleNamesToHtml.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /*
     * Word のスタイル名 → HTML 要素名の対応表 / Word style name to HTML element name
     * キーは正規化後（小文字・半角スペース）の名前で比較する / Keys are compared after normalization
     * "Heading 1"〜"Heading 6" は下の正規表現で拾うため、ここには持たない / Heading 1-6 are matched by regex below
     */
    var STYLE_NAME_MAP = {
        /* 本文 / Body text */
        "normal": "p",
        "default paragraph style": "p",
        "body text": "p",
        "body text 2": "p",
        "no spacing": "p",

        /* 見出し / Headings */
        "title": "h1",
        "subtitle": "h2",
        "toc heading": "h1",

        /* 引用・リスト・キャプション / Quotes, lists and captions */
        "quote": "blockquote",
        "intense quote": "blockquote",
        "list paragraph": "li",
        "list bullet": "li",
        "list number": "li",
        "caption": "figcaption",

        /* ヘッダー・フッター / Header and footer */
        "header": "header",
        "footer": "footer",

        /* 文字装飾由来の名前 / Inline formatting names */
        "hyperlink": "a",
        "emphasis": "em",
        "strong": "strong"
    };

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
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." }
        },
        undo: {
            convertStyleNames: { ja: "Wordスタイル名をHTML要素名に変換", en: "Convert Word style names to HTML element names" }
        },
        report: {
            converted: { ja: "■変換: {n}件", en: "■Converted: {n}" },
            skipped: { ja: "■スキップ: {n}件", en: "■Skipped: {n}" },
            failed: { ja: "■エラー: {n}件", en: "■Errors: {n}" }
        },
        reason: {
            noMatch: { ja: "（対応なし）", en: " (no match)" },
            noChange: { ja: "（変更不要）", en: " (no change needed)" }
        }
    };

    /**
     * ドット区切りキーでラベルを取得する
     * @param {string} labelKey 例: "alert.noDocument"
     * @returns {string} 現在の言語のラベル文字列
     */
    function getLabel(labelKey) {
        var keyParts = labelKey.split(".");
        var node = LABELS;

        for (var i = 0; i < keyParts.length; i++) {
            if (!node || !node.hasOwnProperty(keyParts[i])) {
                throw new Error("Missing label path: " + labelKey);
            }
            node = node[keyParts[i]];
        }

        return node[currentLang] || node.en || labelKey;
    }

    // =========================================
    // 名前の変換 / Name conversion
    // =========================================

    /**
     * スタイル名を比較用に正規化する（小文字化・全角スペースを半角へ・連続空白の圧縮・前後の空白除去）
     * @param {string} styleName 元のスタイル名
     * @returns {string} 正規化した名前
     */
    function normalizeStyleName(styleName) {
        return styleName
            .toLowerCase()
            .replace(/　/g, " ")
            .replace(/\s+/g, " ")
            .replace(/^\s+|\s+$/g, "");
    }

    /**
     * Word のスタイル名に対応する HTML 要素名を返す
     * @param {string} originalName 元のスタイル名
     * @returns {string|null} HTML 要素名。対応がない場合は null
     */
    function toHtmlElementName(originalName) {
        var normalizedName = normalizeStyleName(originalName);

        if (STYLE_NAME_MAP.hasOwnProperty(normalizedName)) {
            return STYLE_NAME_MAP[normalizedName];
        }

        /* "heading 1"〜"heading 6"（スペースなしも許容） / Heading 1-6, with or without a space */
        var headingMatch = normalizedName.match(/^heading\s*([1-6])$/);
        return headingMatch ? "h" + headingMatch[1] : null;
    }

    /**
     * 同じ階層に同名の段落スタイルがあるか調べる
     * @param {Document|ParagraphStyleGroup} container ドキュメントまたは段落スタイルグループ
     * @param {string} styleName 調べる名前
     * @param {string} excludeName 自分自身として除外する名前
     * @returns {boolean} 同名のスタイルがあれば true
     */
    function isStyleNameTaken(container, styleName, excludeName) {
        if (styleName === excludeName) return false;
        return container.paragraphStyles.itemByName(styleName).isValid;
    }

    /**
     * 同じ階層で衝突しない名前を作る（衝突する場合は "_1" "_2" と連番を付ける）
     * @param {Document|ParagraphStyleGroup} container ドキュメントまたは段落スタイルグループ
     * @param {string} baseName 希望する名前
     * @param {string} excludeName 自分自身として除外する名前
     * @returns {string} 使用できる名前
     */
    function makeUniqueStyleName(container, baseName, excludeName) {
        var candidateName = baseName;
        var suffixNumber = 1;

        while (isStyleNameTaken(container, candidateName, excludeName)) {
            candidateName = baseName + "_" + suffixNumber;
            suffixNumber++;
        }

        return candidateName;
    }

    // =========================================
    // スタイルのリネーム / Renaming styles
    // =========================================

    /**
     * @typedef {object} RenameResult
     * @property {Array<string>} renamed リネームしたスタイルの記録
     * @property {Array<string>} skipped 対象外としたスタイルの記録
     * @property {Array<string>} failed リネームに失敗したスタイルの記録
     */

    /**
     * 段落スタイル1つをリネームする
     * @param {Document|ParagraphStyleGroup} container 親コンテナー
     * @param {ParagraphStyle} paragraphStyle 対象の段落スタイル
     * @param {RenameResult} renameResult 結果の記録先
     * @returns {void}
     */
    function renameOneStyle(container, paragraphStyle, renameResult) {
        var originalName = paragraphStyle.name;

        /* "[基本段落]" / "[No Paragraph Style]" など角かっこ始まりは対象外 / Skip bracketed built-in styles */
        if (originalName.indexOf("[") === 0) return;

        var htmlElementName = toHtmlElementName(originalName);

        if (htmlElementName === null) {
            renameResult.skipped.push(originalName + getLabel("reason.noMatch"));
            return;
        }

        if (htmlElementName === originalName) {
            renameResult.skipped.push(originalName + getLabel("reason.noChange"));
            return;
        }

        var newName = makeUniqueStyleName(container, htmlElementName, originalName);

        try {
            paragraphStyle.name = newName;
            renameResult.renamed.push(originalName + " → " + newName);
        } catch (e) {
            /* ロックされたスタイルなど、InDesign 側がリネームを拒むことがある / InDesign may refuse the rename */
            renameResult.failed.push(originalName + " : " + e.message);
        }
    }

    /**
     * コンテナー内の段落スタイルを再帰的にリネームする
     * @param {Document|ParagraphStyleGroup} container ドキュメントまたは段落スタイルグループ
     * @param {RenameResult} renameResult 結果の記録先
     * @returns {void}
     */
    function renameStylesInContainer(container, renameResult) {
        /* サブグループを先に処理 / Process subgroups first */
        var styleGroups = container.paragraphStyleGroups.everyItem().getElements();
        for (var i = 0; i < styleGroups.length; i++) {
            renameStylesInContainer(styleGroups[i], renameResult);
        }

        /* リネーム中に並び順が変わっても影響しないよう、先に配列へ取り出す / Snapshot before renaming */
        var stylesInContainer = container.paragraphStyles.everyItem().getElements();
        for (var j = 0; j < stylesInContainer.length; j++) {
            renameOneStyle(container, stylesInContainer[j], renameResult);
        }
    }

    // =========================================
    // 結果レポート / Result report
    // =========================================

    /**
     * 見出し1行と明細をまとめた1セクション分の文字列を作る
     * @param {string} labelKey 見出しラベルのドット区切りキー
     * @param {Array<string>} entryLines 明細行
     * @returns {string} 見出しと明細を連結した文字列
     */
    function buildReportSection(labelKey, entryLines) {
        var headingLine = getLabel(labelKey).replace("{n}", entryLines.length);
        return entryLines.length > 0 ? headingLine + "\n" + entryLines.join("\n") : headingLine;
    }

    /**
     * 実行結果をアラート用の文字列にまとめる
     * @param {RenameResult} renameResult 集計結果
     * @returns {string} レポート文字列
     */
    function buildReport(renameResult) {
        var reportSections = [buildReportSection("report.converted", renameResult.renamed)];

        if (renameResult.skipped.length > 0) {
            reportSections.push(buildReportSection("report.skipped", renameResult.skipped));
        }
        if (renameResult.failed.length > 0) {
            reportSections.push(buildReportSection("report.failed", renameResult.failed));
        }

        return reportSections.join("\n\n");
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメント内の段落スタイル名を HTML 要素名へ変換する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var activeDoc = app.activeDocument;
        var renameResult = { renamed: [], skipped: [], failed: [] };

        /* Undo 可能な1操作としてまとめる / Group into a single undo step */
        app.doScript(
            function () {
                renameStylesInContainer(activeDoc, renameResult);
            },
            ScriptLanguage.JAVASCRIPT,
            undefined,
            UndoModes.ENTIRE_SCRIPT,
            getLabel("undo.convertStyleNames")
        );

        alert(buildReport(renameResult));
    }

    main();

})();
