#target indesign

/*

### 概要

段落スタイル・文字スタイル・オブジェクトスタイル・スウォッチのうち、ドキュメント内で使われていないものを削除します。

詳細は README を参照してください。

### Overview

Deletes paragraph styles, character styles, object styles, and swatches that are not used in the document.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdDeleteUnused";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-24";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */

    /* ボタン / Buttons */
    var BUTTON_ROW_TOP_MARGIN = 10;          /* ボタン列の上余白 / top margin above the button row */

    // =========================================
    // ラベル定義 / Labels
    // =========================================

    /**
     * UI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentUILang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentUILang();

    var LABELS = {
        dialog: {
            title: { ja: "未使用項目の削除", en: "Delete Unused Items" }
        },
        panel: {
            scope:   { ja: "対象", en: "Target" },
            targets: { ja: "削除する項目", en: "Items to Delete" }
        },
        radio: {
            activeDocument: { ja: "このドキュメント", en: "This document" },
            allDocuments:   { ja: "すべてのドキュメント", en: "All documents" }
        },
        checkbox: {
            paragraphStyle: { ja: "段落スタイル", en: "Paragraph styles" },
            characterStyle: { ja: "文字スタイル", en: "Character styles" },
            objectStyle:    { ja: "オブジェクトスタイル", en: "Object styles" },
            swatch:         { ja: "スウォッチ", en: "Swatches" }
        },
        tooltip: {
            paragraphStyle: {
                ja: "テキストにも、他のスタイルの「基準」「次のスタイル」、オブジェクトスタイル、セルスタイル、目次スタイル、脚注の設定にも使われていない段落スタイルを削除します。［ ］で囲まれた既定のスタイルは残します。",
                en: "Deletes paragraph styles not used in text, as Based On / Next Style of other styles, or in object, cell, TOC styles and footnote options. Default styles in [ ] are kept."
            },
            characterStyle: {
                ja: "テキストにも、他のスタイルの「基準」、先頭文字スタイル・正規表現スタイル・行スタイル、箇条書き、ドロップキャップ、目次スタイル、脚注の設定にも使われていない文字スタイルを削除します。［なし］は残します。",
                en: "Deletes character styles not used in text, as Based On, in nested / GREP / line styles, bullets and numbering, drop caps, TOC styles, or footnote options. [None] is kept."
            },
            objectStyle: {
                ja: "どのオブジェクトにも、他のスタイルの「基準」にも、新規オブジェクトの既定にも使われていないオブジェクトスタイルを削除します。［ ］で囲まれた既定のスタイルは残します。",
                en: "Deletes object styles not applied to any object, not used as Based On, and not set as the default for new objects. Default styles in [ ] are kept."
            },
            swatch: {
                ja: "使われていないスウォッチを削除します。グラデーションの分岐点や濃淡の元になっているカラー、名前のないカラー、削除できない既定のスウォッチは残します。",
                en: "Deletes unused swatches. Colors used in gradient stops or as the base of tints, unnamed colors, and default swatches that cannot be deleted are kept."
            },
            optionClick: {
                ja: "option（Alt）+クリック：この項目だけをON。もう一度押すとすべてをON",
                en: "Option (Alt)-click: check only this item. Do it again to check all"
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            result:     { ja: "削除しました。", en: "Deleted." },
            docCount:   { ja: "%1個のドキュメントを処理しました。", en: "Processed %1 documents." },
            countLine:  { ja: "%1：%2個", en: "%1: %2" }
        },
        undoName: { ja: "未使用項目の削除", en: "Delete Unused Items" }
    };

    /**
     * ラベル定義から現在のUI言語の文字列を取り出す
     * @param {{ja: string, en: string}} labelSet - 言語別のラベル定義
     * @returns {string} 現在のUI言語の文字列
     */
    function getLabel(labelSet) {
        return labelSet[uiLang] || labelSet.en;
    }

    /**
     * ラベル内のプレースホルダー（%1, %2 …）を値で置き換える
     * @param {string} template - プレースホルダーを含む文字列
     * @param {Array<string>} values - 差し込む値
     * @returns {string} 置き換え後の文字列
     */
    function formatLabel(template, values) {
        var text = template;
        for (var i = 0; i < values.length; i++) {
            text = text.split("%" + (i + 1)).join(String(values[i]));
        }
        return text;
    }

    // =========================================
    // 削除対象の定義 / Target definitions
    // =========================================

    /* 削除対象の種類（ダイアログの並び順） / Target kinds, in dialog order */
    var TARGET_KEYS = ["paragraphStyle", "characterStyle", "objectStyle", "swatch"];

    /* 初期状態でONにする種類 / Kinds checked by default */
    var DEFAULT_CHECKED_TARGETS = { swatch: true };

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ウィンドウの共通設定を適用する
     * @param {Window} win - 対象ウィンドウ
     * @returns {void}
     */
    function setupWindow(win) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = WINDOW_SPACING;
    }

    /**
     * パネルの共通設定を適用する
     * @param {Panel} panel - 対象パネル
     * @returns {void}
     */
    function setupPanel(panel) {
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.margins = PANEL_MARGINS;
        panel.spacing = PANEL_SPACING;
    }

    /**
     * 対象ドキュメントのラジオボタンを並べる（初期値は「このドキュメント」）
     * @param {Window} parent - 追加先のウィンドウ
     * @returns {RadioButton} 「すべてのドキュメント」のラジオボタン
     */
    function addScopeRadios(parent) {
        var scopePanel = parent.add("panel", undefined, getLabel(LABELS.panel.scope));
        setupPanel(scopePanel);

        var activeDocumentRadio = scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.activeDocument));
        var allDocumentsRadio = scopePanel.add("radiobutton", undefined, getLabel(LABELS.radio.allDocuments));
        activeDocumentRadio.value = true;
        return allDocumentsRadio;
    }

    /**
     * 削除する項目のチェックボックスを並べる（初期値は DEFAULT_CHECKED_TARGETS のものだけON）
     * @param {Window} parent - 追加先のウィンドウ
     * @returns {Object<string, Checkbox>} 種類ごとのチェックボックス
     */
    function addTargetCheckboxes(parent) {
        var targetPanel = parent.add("panel", undefined, getLabel(LABELS.panel.targets));
        setupPanel(targetPanel);

        var targetCheckboxes = {};
        for (var i = 0; i < TARGET_KEYS.length; i++) {
            var targetKey = TARGET_KEYS[i];
            var targetCheckbox = targetPanel.add("checkbox", undefined, getLabel(LABELS.checkbox[targetKey]));
            targetCheckbox.value = DEFAULT_CHECKED_TARGETS[targetKey] === true;
            targetCheckbox.helpTip = getLabel(LABELS.tooltip[targetKey]) + "\n\n" + getLabel(LABELS.tooltip.optionClick);
            targetCheckboxes[targetKey] = targetCheckbox;
        }
        return targetCheckboxes;
    }

    /**
     * option（Alt）+クリックで、押した項目だけをONにする
     * すでにそれだけがONの状態で押したときは、すべてをONに戻す
     * onClick の時点で押した項目の値は反転済みなので、反転前の状態から判定する
     * @param {Object<string, Checkbox>} targetCheckboxes - 種類ごとのチェックボックス
     * @param {string} clickedKey - 押された項目の種類
     * @returns {void}
     */
    function applyOptionClick(targetCheckboxes, clickedKey) {
        var wasChecked = !targetCheckboxes[clickedKey].value;
        var othersUnchecked = true;
        for (var i = 0; i < TARGET_KEYS.length; i++) {
            if (TARGET_KEYS[i] !== clickedKey && targetCheckboxes[TARGET_KEYS[i]].value) othersUnchecked = false;
        }
        var checkAll = wasChecked && othersUnchecked;
        for (var j = 0; j < TARGET_KEYS.length; j++) {
            targetCheckboxes[TARGET_KEYS[j]].value = checkAll || TARGET_KEYS[j] === clickedKey;
        }
    }

    /**
     * キャンセル／OK のボタン列を追加する
     * @param {Window} parent - 追加先のウィンドウ
     * @returns {Button} OK ボタン
     */
    function addButtonRow(parent) {
        // メイングループ（横並び） / Main group (horizontal layout)
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        // スペーサー（伸縮）/ Spacer (stretchable)
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        // 右側グループ / Right-side button group
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        return btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
    }

    /**
     * ダイアログを表示して、対象ドキュメントと削除する項目を選ばせる
     * @returns {{allDocuments: boolean, targets: Object<string, boolean>}|null} 選んだ内容。キャンセル時は null
     */
    function showTargetDialog() {
        var targetDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(targetDialog);

        var allDocumentsRadio = addScopeRadios(targetDialog);
        var targetCheckboxes = addTargetCheckboxes(targetDialog);
        var btnOK = addButtonRow(targetDialog);

        /* 1つもONでなければ OK を押せなくする / Disable OK when nothing is checked */
        function updateOKButton() {
            var hasTarget = false;
            for (var i = 0; i < TARGET_KEYS.length; i++) {
                if (targetCheckboxes[TARGET_KEYS[i]].value) hasTarget = true;
            }
            btnOK.enabled = hasTarget;
        }
        /**
         * チェックボックスのクリック処理を作る（option 併用の判定と OK ボタンの更新）
         * @param {string} targetKey - 対象の種類
         * @returns {Function} onClick に渡す関数
         */
        function createCheckboxClickHandler(targetKey) {
            return function () {
                if (ScriptUI.environment.keyboardState.altKey) applyOptionClick(targetCheckboxes, targetKey);
                updateOKButton();
            };
        }
        for (var i = 0; i < TARGET_KEYS.length; i++) {
            targetCheckboxes[TARGET_KEYS[i]].onClick = createCheckboxClickHandler(TARGET_KEYS[i]);
        }

        if (targetDialog.show() !== 1) return null;

        var selectedTargets = {};
        for (var j = 0; j < TARGET_KEYS.length; j++) {
            selectedTargets[TARGET_KEYS[j]] = targetCheckboxes[TARGET_KEYS[j]].value;
        }
        return { allDocuments: allDocumentsRadio.value, targets: selectedTargets };
    }

    // =========================================
    // 参照の収集 / Reference collection
    // =========================================

    /**
     * 値が指定クラスのスタイルなら、参照済みとして記録する
     * 未設定のときは文字列や NothingEnum が返るので、クラスで見分ける
     * @param {Object<string, boolean>} referenceMap - スタイルIDをキーにした参照済みの記録
     * @param {*} value - スタイルのプロパティ値
     * @param {Function} styleClass - ParagraphStyle などのクラス
     * @returns {void}
     */
    function markReferenced(referenceMap, value, styleClass) {
        if (value instanceof styleClass) referenceMap[value.id] = true;
    }

    /**
     * ルートスタイル（［段落スタイルなし］や［なし］）を除いたスタイルの一覧を返す
     * ルートスタイルは basedOn などを読むだけで「ルートスタイルに対する無効な要求」の例外になる
     * @param {Array<ParagraphStyle|CharacterStyle|ObjectStyle|CellStyle>} allStyles - allParagraphStyles などの一覧
     * @param {ParagraphStyle|CharacterStyle|ObjectStyle} rootStyle - コレクションの先頭にあるルートスタイル
     * @returns {Array<ParagraphStyle|CharacterStyle|ObjectStyle>} ルートスタイルを除いた一覧
     */
    function excludeRootStyle(allStyles, rootStyle) {
        var styles = [];
        for (var i = 0; i < allStyles.length; i++) {
            if (allStyles[i].id !== rootStyle.id) styles.push(allStyles[i]);
        }
        return styles;
    }

    /**
     * 先頭文字スタイル・正規表現スタイル・行スタイルの文字スタイルを参照済みにする
     * @param {Object<string, boolean>} characterRefs - 文字スタイルの参照済みの記録
     * @param {ParagraphStyle} paragraphStyle - 調べる段落スタイル
     * @returns {void}
     */
    function markNestedCharacterStyles(characterRefs, paragraphStyle) {
        var nestedCollections = [paragraphStyle.nestedStyles, paragraphStyle.nestedGrepStyles, paragraphStyle.nestedLineStyles];
        for (var i = 0; i < nestedCollections.length; i++) {
            for (var j = 0; j < nestedCollections[i].length; j++) {
                markReferenced(characterRefs, nestedCollections[i][j].appliedCharacterStyle, CharacterStyle);
            }
        }
    }

    /**
     * 段落スタイル同士・段落スタイルから文字スタイルへの参照を集める
     * @param {Document} doc - 対象ドキュメント
     * @param {{paragraph: Object, character: Object, object: Object}} refs - 参照済みの記録
     * @returns {void}
     */
    function collectParagraphStyleRefs(doc, refs) {
        var paragraphStyles = excludeRootStyle(doc.allParagraphStyles, doc.paragraphStyles[0]);
        for (var i = 0; i < paragraphStyles.length; i++) {
            var paragraphStyle = paragraphStyles[i];
            markReferenced(refs.paragraph, paragraphStyle.basedOn, ParagraphStyle);
            /* 自分自身を「次のスタイル」にしているのは参照に数えない / Self as Next Style does not count */
            if (paragraphStyle.nextStyle instanceof ParagraphStyle && paragraphStyle.nextStyle.id !== paragraphStyle.id) {
                refs.paragraph[paragraphStyle.nextStyle.id] = true;
            }
            markReferenced(refs.character, paragraphStyle.bulletsCharacterStyle, CharacterStyle);
            markReferenced(refs.character, paragraphStyle.numberingCharacterStyle, CharacterStyle);
            markReferenced(refs.character, paragraphStyle.dropCapStyle, CharacterStyle);
            markNestedCharacterStyles(refs.character, paragraphStyle);
        }

        var characterStyles = excludeRootStyle(doc.allCharacterStyles, doc.characterStyles[0]);
        for (var j = 0; j < characterStyles.length; j++) {
            markReferenced(refs.character, characterStyles[j].basedOn, CharacterStyle);
        }
    }

    /**
     * オブジェクトスタイルの基準と、そこから段落スタイルへの参照、
     * 新規オブジェクトの既定、ページアイテムへの適用を集める
     * @param {Document} doc - 対象ドキュメント
     * @param {{paragraph: Object, character: Object, object: Object}} refs - 参照済みの記録
     * @returns {void}
     */
    function collectObjectStyleRefs(doc, refs) {
        var objectStyles = excludeRootStyle(doc.allObjectStyles, doc.objectStyles[0]);
        for (var i = 0; i < objectStyles.length; i++) {
            markReferenced(refs.object, objectStyles[i].basedOn, ObjectStyle);
            markReferenced(refs.paragraph, objectStyles[i].appliedParagraphStyle, ParagraphStyle);
        }

        var itemDefaults = doc.pageItemDefaults;
        markReferenced(refs.object, itemDefaults.appliedTextObjectStyle, ObjectStyle);
        markReferenced(refs.object, itemDefaults.appliedGraphicObjectStyle, ObjectStyle);
        markReferenced(refs.object, itemDefaults.appliedGridObjectStyle, ObjectStyle);

        var pageItems = doc.allPageItems;
        for (var j = 0; j < pageItems.length; j++) {
            markReferenced(refs.object, pageItems[j].appliedObjectStyle, ObjectStyle);
        }
    }

    /**
     * セルスタイル・目次スタイル・脚注・テキストの既定からの参照を集める
     * @param {Document} doc - 対象ドキュメント
     * @param {{paragraph: Object, character: Object, object: Object}} refs - 参照済みの記録
     * @returns {void}
     */
    function collectOtherStyleRefs(doc, refs) {
        var cellStyles = excludeRootStyle(doc.allCellStyles, doc.cellStyles[0]);
        for (var i = 0; i < cellStyles.length; i++) {
            markReferenced(refs.paragraph, cellStyles[i].appliedParagraphStyle, ParagraphStyle);
        }

        /* 目次の項目は、対象スタイルを名前で持つ / TOC entries hold their source style by name */
        var tocEntryNames = {};
        for (var j = 0; j < doc.tocStyles.length; j++) {
            var tocStyle = doc.tocStyles[j];
            markReferenced(refs.paragraph, tocStyle.titleStyle, ParagraphStyle);
            for (var k = 0; k < tocStyle.tocStyleEntries.length; k++) {
                var tocEntry = tocStyle.tocStyleEntries[k];
                tocEntryNames[tocEntry.name] = true;
                markReferenced(refs.paragraph, tocEntry.formatStyle, ParagraphStyle);
                markReferenced(refs.character, tocEntry.pageNumberStyle, CharacterStyle);
                markReferenced(refs.character, tocEntry.separatorStyle, CharacterStyle);
            }
        }
        var paragraphStyles = doc.allParagraphStyles;
        for (var m = 0; m < paragraphStyles.length; m++) {
            if (tocEntryNames[paragraphStyles[m].name]) refs.paragraph[paragraphStyles[m].id] = true;
        }

        markReferenced(refs.paragraph, doc.footnoteOptions.footnoteTextStyle, ParagraphStyle);
        markReferenced(refs.character, doc.footnoteOptions.footnoteMarkerStyle, CharacterStyle);
        markReferenced(refs.paragraph, doc.textDefaults.appliedParagraphStyle, ParagraphStyle);
        markReferenced(refs.character, doc.textDefaults.appliedCharacterStyle, CharacterStyle);
    }

    /**
     * スタイル同士やドキュメント設定からの参照をまとめて集める
     * テキストへの適用は検索で調べるので、ここには含めない
     * @param {Document} doc - 対象ドキュメント
     * @returns {{paragraph: Object, character: Object, object: Object}} 種類ごとの参照済みスタイルID
     */
    function collectStyleRefs(doc) {
        var refs = { paragraph: {}, character: {}, object: {} };
        collectParagraphStyleRefs(doc, refs);
        collectObjectStyleRefs(doc, refs);
        collectOtherStyleRefs(doc, refs);
        return refs;
    }

    // =========================================
    // 削除 / Deletion
    // =========================================

    /**
     * スタイルがテキストに適用されているかを検索で調べる
     * @param {Document} doc - 対象ドキュメント
     * @param {string} findProperty - "appliedParagraphStyle" または "appliedCharacterStyle"
     * @param {ParagraphStyle|CharacterStyle} style - 調べるスタイル
     * @returns {boolean} 適用されていれば true
     */
    function isAppliedToText(doc, findProperty, style) {
        app.findTextPreferences = NothingEnum.NOTHING;
        app.findTextPreferences[findProperty] = style;
        var isApplied = doc.findText().length > 0;
        app.findTextPreferences = NothingEnum.NOTHING;
        return isApplied;
    }

    /**
     * 参照されていないスタイルを削除する
     * ［ ］で囲まれた既定のスタイルは削除できないので対象外
     * @param {Array<ParagraphStyle|CharacterStyle|ObjectStyle>} styles - 候補のスタイル
     * @param {Object<string, boolean>} referenceMap - 参照済みのスタイルID
     * @param {string|null} findProperty - テキスト検索に使うプロパティ名。オブジェクトスタイルは null
     * @param {Document} doc - 対象ドキュメント
     * @returns {number} 削除した数
     */
    function removeUnreferencedStyles(styles, referenceMap, findProperty, doc) {
        var removedCount = 0;
        for (var i = styles.length - 1; i >= 0; i--) {
            var style = styles[i];
            if (style.name.charAt(0) === "[" || referenceMap[style.id]) continue;
            if (findProperty && isAppliedToText(doc, findProperty, style)) continue;
            style.remove();
            removedCount++;
        }
        return removedCount;
    }

    /**
     * 選ばれた種類のスタイルを、参照がなくなるまで繰り返し削除する
     * 子スタイルを消すと、その基準だった親が未使用になることがあるため
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, boolean>} selectedTargets - 種類ごとのON/OFF
     * @param {Object<string, number>} removedCounts - 種類ごとの削除数（加算していく）
     * @returns {void}
     */
    function removeUnusedStyles(doc, selectedTargets, removedCounts) {
        var passRemovedCount;
        do {
            var refs = collectStyleRefs(doc);
            var countsThisPass = {
                paragraphStyle: selectedTargets.paragraphStyle ? removeUnreferencedStyles(doc.allParagraphStyles, refs.paragraph, "appliedParagraphStyle", doc) : 0,
                characterStyle: selectedTargets.characterStyle ? removeUnreferencedStyles(doc.allCharacterStyles, refs.character, "appliedCharacterStyle", doc) : 0,
                objectStyle:    selectedTargets.objectStyle ? removeUnreferencedStyles(doc.allObjectStyles, refs.object, null, doc) : 0
            };
            passRemovedCount = 0;
            for (var targetKey in countsThisPass) {
                removedCounts[targetKey] += countsThisPass[targetKey];
                passRemovedCount += countsThisPass[targetKey];
            }
        } while (passRemovedCount > 0);
    }

    /**
     * グラデーションの分岐点と濃淡の元になっているカラーのIDを集める
     * unusedSwatches にはこれらが含まれ、消すとグラデーションや濃淡が変わってしまう
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object<string, boolean>} 使われているカラーのID
     */
    function collectSwatchRefs(doc) {
        var swatchRefs = {};
        for (var i = 0; i < doc.gradients.length; i++) {
            var gradientStops = doc.gradients[i].gradientStops;
            for (var j = 0; j < gradientStops.length; j++) {
                markReferenced(swatchRefs, gradientStops[j].stopColor, Swatch);
            }
        }
        for (var k = 0; k < doc.tints.length; k++) {
            markReferenced(swatchRefs, doc.tints[k].baseColor, Swatch);
        }
        return swatchRefs;
    }

    /**
     * 使用していないスウォッチを削除する
     * 名前が空のスウォッチ（名前なしのカラー）と、削除できない既定のスウォッチは残す
     * @param {Document} doc - 対象ドキュメント
     * @returns {number} 削除した数
     */
    function removeUnusedSwatches(doc) {
        var swatchRefs = collectSwatchRefs(doc);
        var unusedSwatchList = doc.unusedSwatches;
        var removedCount = 0;
        for (var i = unusedSwatchList.length - 1; i >= 0; i--) {
            var swatch = unusedSwatchList[i];
            if (swatch.name === "" || swatchRefs[swatch.id]) continue;
            /* Swatch には削除可否のプロパティが無いので、既定のスウォッチは remove() の例外で見分ける
               / Swatch has no "removable" property, so default swatches are detected by remove() throwing */
            try {
                swatch.remove();
                removedCount++;
            } catch (e) {}
        }
        return removedCount;
    }

    /**
     * テキスト検索の範囲を、非表示・ロック・マスター・脚注まで広げる
     * @returns {Object<string, boolean>} 変更前の設定
     */
    function widenFindScope() {
        var findOptions = app.findChangeTextOptions;
        var scopeKeys = ["includeHiddenLayers", "includeLockedLayersForFind", "includeLockedStoriesForFind", "includeMasterPages", "includeFootnotes"];
        var savedScope = {};
        for (var i = 0; i < scopeKeys.length; i++) {
            savedScope[scopeKeys[i]] = findOptions[scopeKeys[i]];
            findOptions[scopeKeys[i]] = true;
        }
        return savedScope;
    }

    /**
     * 選ばれた種類の未使用項目を削除する
     * スタイルを先に消し、それで使われなくなったスウォッチも拾う
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, boolean>} selectedTargets - 種類ごとのON/OFF
     * @returns {Object<string, number>} 種類ごとの削除数
     */
    function removeUnusedItems(doc, selectedTargets) {
        var removedCounts = { paragraphStyle: 0, characterStyle: 0, objectStyle: 0, swatch: 0 };

        var savedScope = widenFindScope();
        try {
            removeUnusedStyles(doc, selectedTargets, removedCounts);
        } finally {
            /* 検索範囲はユーザーの設定なので、失敗しても戻す / Always restore the user's find scope */
            app.findChangeTextOptions.properties = savedScope;
        }

        if (selectedTargets.swatch) removedCounts.swatch = removeUnusedSwatches(doc);
        return removedCounts;
    }

    /**
     * 削除数を、選んだ種類だけ並べて表示する
     * @param {Object<string, boolean>} selectedTargets - 種類ごとのON/OFF
     * @param {Object<string, number>} removedCounts - 種類ごとの削除数（全ドキュメントの合計）
     * @param {number} documentCount - 処理したドキュメント数
     * @returns {void}
     */
    function showResult(selectedTargets, removedCounts, documentCount) {
        var resultLines = [getLabel(LABELS.alert.result)];
        if (documentCount > 1) resultLines.push(formatLabel(getLabel(LABELS.alert.docCount), [documentCount]));
        resultLines.push("");
        for (var i = 0; i < TARGET_KEYS.length; i++) {
            var targetKey = TARGET_KEYS[i];
            if (!selectedTargets[targetKey]) continue;
            resultLines.push(formatLabel(getLabel(LABELS.alert.countLine), [getLabel(LABELS.checkbox[targetKey]), removedCounts[targetKey]]));
        }
        alert(resultLines.join("\n"), getLabel(LABELS.dialog.title));
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * 1つのドキュメントの未使用項目を、1回の取り消し単位で削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {Object<string, boolean>} selectedTargets - 種類ごとのON/OFF
     * @returns {Object<string, number>} 種類ごとの削除数
     */
    function removeUnusedItemsWithUndo(doc, selectedTargets) {
        var removedCounts;
        app.doScript(function () {
            removedCounts = removeUnusedItems(doc, selectedTargets);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undoName));
        return removedCounts;
    }

    /**
     * ダイアログで選んだドキュメントと種類の未使用項目を削除する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument), getLabel(LABELS.dialog.title));
            return;
        }

        var dialogResult = showTargetDialog();
        if (!dialogResult) return;

        var targetDocs = dialogResult.allDocuments ? app.documents.everyItem().getElements() : [app.activeDocument];
        var totalCounts = { paragraphStyle: 0, characterStyle: 0, objectStyle: 0, swatch: 0 };
        for (var i = 0; i < targetDocs.length; i++) {
            var docCounts = removeUnusedItemsWithUndo(targetDocs[i], dialogResult.targets);
            for (var targetKey in totalCounts) totalCounts[targetKey] += docCounts[targetKey];
        }

        showResult(dialogResult.targets, totalCounts, targetDocs.length);
    }

    main();

})();
