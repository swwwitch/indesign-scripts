#target "InDesign"

var SCRIPT_VERSION = "v2.6";

/*
 * 選択した文字列からフレーム作成 / Create Frame from Selected Text
 *
 * 概要:
 * 選択したテキストのサイズをもとに、同サイズのグラフィックフレームを作成します。
 * テキスト選択時はインライン（アンカー付き）またはページ上への配置を選択でき、
 * 挿入行の段落スタイル、オブジェクトスタイル、テキストの回り込みを指定できます。
 * テキスト未選択時は、ページ上にグラフィックフレームを作成できます。
 *
 * Summary:
 * Creates a graphic frame based on the size of the selected text.
 * When text is selected, you can choose either inline (anchored) insertion or placement on the page,
 * and specify the paragraph style for the inserted line, object style, and text wrap options.
 * When no text is selected, the script can create a graphic frame on the page.
 *
 * 更新日: 2026-03-17
 * Updated: 2026-03-17
 *
 * 紹介記事 / Article
 * https://note.com/dtp_tranist/n/ndd1b7c5246a3
 * 
 * 更新履歴:
 * - v1.0
 * - v2.6 行数モードの概算tooltip追加、InsertionPointガード強化、単位変換整理
 *
 * Version history:
 * - v1.0
 * - v2.6 Added approximate-height tooltip for line count mode, strengthened insertion point guards, and cleaned up unit conversion
 */

// オリジナルアイデア
// DTP Script note さん
// https://note.com/yosi2631/n/ned2dbc1cb79d

var LABELS = {
    dialogTitle: {
        ja: "選択した文字列からフレーム作成",
        en: "Create Frame from Selected Text"
    },
    noDoc: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    selectText: {
        ja: "テキストを選択してください。",
        en: "Please select text."
    },
    selectTextBeforeRun: {
        ja: "テキスト項目を1つだけ選択してから実行してください。",
        en: "Select only one text item before running the script."
    },
    boundsError: {
        ja: "選択したテキストの座標を取得できませんでした。",
        en: "Could not get the bounds of the selected text."
    },
    parentPageError: {
        ja: "配置先のページを取得できませんでした。",
        en: "Could not determine the destination page."
    },
    parentTextFrameError: {
        ja: "挿入位置の親テキストフレームを取得できませんでした。",
        en: "Could not get the parent text frame for the insertion point."
    },
    methodPanel: { ja: "配置方法", en: "Placement" },
    graphicFrame: { ja: "グラフィックフレーム", en: "Graphic Frame" },
    inlineFrame: { ja: "インライン（アンカー付き）", en: "Inline (Anchored)" },
    frameSizePanel: { ja: "フレームサイズ", en: "Frame Size" },
    widthPanel: { ja: "幅", en: "Width" },
    widthText: { ja: "選択したテキスト", en: "Selected Text" },
    widthColumn: { ja: "カラム幅", en: "Column Width" },
    widthFrame: { ja: "親フレーム", en: "Parent Frame" },
    textSettingsPanel: { ja: "テキスト設定", en: "Text Settings" },
    paraStyle: { ja: "挿入行の段落スタイル:", en: "Paragraph Style for Inserted Line:" },
    objectSettingsPanel: { ja: "オブジェクト設定", en: "Object Settings" },
    objStyle: { ja: "オブジェクトスタイル:", en: "Object Style:" },
    wrap: { ja: "テキストの回り込み:", en: "Text Wrap:" },
    wrapNone: { ja: "なし", en: "None" },
    wrapBoundingBox: { ja: "境界線ボックスで回り込む", en: "Wrap Around Bounding Box" },
    wrapContour: { ja: "オブジェクトのシェイプで回り込む", en: "Wrap Around Object Shape" },
    wrapJumpObject: { ja: "オブジェクトを挟んで回り込む", en: "Jump Object" },
    wrapNextColumn: { ja: "次の段へテキストを送る", en: "Jump to Next Column" },
    autoLeading: { ja: "行送り：自動", en: "Leading: Auto" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" },
    undoName: { ja: "フレーム作成", en: "Create Frame" },
    dialogTitleNoSelection: {
        ja: "グラフィックフレーム作成",
        en: "Create Graphic Frame"
    },
    widthMargin: { ja: "ページのマージン", en: "Page Margins" },
    heightPanel: { ja: "高さ", en: "Height" },
    heightLines: { ja: "行数", en: "Line Count" },
    heightLinesTip: {
        ja: "行数モードの高さは概算です。1行目の文字サイズ＋残り行数×行送りで計算します。",
        en: "Line Count height is approximate. It is calculated as first-line point size plus leading for the remaining lines."
    },
    heightSize: { ja: "サイズ指定", en: "Specify Size" },
    unitLines: { ja: "行", en: "lines" },
    unitMm: { ja: "mm", en: "mm" }
};

function L(key) {
    return LABELS[key][lang];
}

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

(function () {
    if (app.documents.length === 0) {
        alert(L("noDoc"));
        return;
    }

    var selection = null;
    var isInsertionPoint = false;
    var noSelection = false;

    if (app.selection.length === 0) {
        // 何も選択していない場合
        noSelection = true;
    } else if (app.selection.length !== 1) {
        alert(L("selectTextBeforeRun"));
        return;
    } else {
        selection = app.selection[0];
        if (selection.constructor && selection.constructor.name === "InsertionPoint") {
            isInsertionPoint = true;
        } else if (!isTextSelection(selection)) {
            alert(L("selectText"));
            return;
        }
    }

    var doc = app.activeDocument;

    // 選択テキスト範囲の外接矩形を取得（InsertionPoint・未選択時はスキップ）
    var selectionBounds = null;
    var hasTextBounds = false;
    if (!isInsertionPoint && !noSelection) {
        selectionBounds = getTextSelectionBounds(selection);
        if (!selectionBounds) {
            alert(L("boundsError"));
            return;
        }
        hasTextBounds = true;
    }

    var width = hasTextBounds ? (selectionBounds[3] - selectionBounds[1]) : 0;
    var height = hasTextBounds ? (selectionBounds[2] - selectionBounds[0]) : 0;

    // オブジェクトスタイル一覧を取得
    var objStyleNames = [];
    for (var s = 0; s < doc.objectStyles.length; s++) {
        objStyleNames.push(doc.objectStyles[s].name);
    }

    // 段落スタイル一覧を取得
    var paraStyleNames = [];
    var defaultParaStyleIndex = 0;
    for (var p = 0; p < doc.paragraphStyles.length; p++) {
        var pName = doc.paragraphStyles[p].name;
        paraStyleNames.push(pName);
        if (pName === "p.img" || pName.indexOf(".img") !== -1) {
            defaultParaStyleIndex = p;
        }
    }

    // 回り込みの選択肢
    var wrapLabels = [
        L("wrapNone"),
        L("wrapBoundingBox"),
        L("wrapContour"),
        L("wrapJumpObject"),
        L("wrapNextColumn")
    ];
    var wrapModes = [
        TextWrapModes.NONE,
        TextWrapModes.BOUNDING_BOX_TEXT_WRAP,
        TextWrapModes.CONTOUR,
        TextWrapModes.JUMP_OBJECT_TEXT_WRAP,
        TextWrapModes.NEXT_COLUMN_TEXT_WRAP
    ];

    // 「なし」オブジェクトスタイルのインデックスを特定
    var noneObjStyleIndex = 0;
    for (var n = 0; n < objStyleNames.length; n++) {
        if (objStyleNames[n] === "[なし]" || objStyleNames[n] === "[None]") {
            noneObjStyleIndex = n;
            break;
        }
    }

    // ScriptUIダイアログ
    var dlgTitle = (isInsertionPoint || noSelection) ? L("dialogTitleNoSelection") : L("dialogTitle");
    var dlg = new Window("dialog", dlgTitle + " " + SCRIPT_VERSION);
    dlg.alignChildren = ["left", "top"];

    // 追加方法パネル
    var methodPanel = dlg.add("panel", undefined, L("methodPanel"));
    methodPanel.alignChildren = ["left", "top"];
    methodPanel.margins = [15, 20, 15, 10];
    var radioGraphic = methodPanel.add("radiobutton", undefined, L("graphicFrame"));
    var radioInline = methodPanel.add("radiobutton", undefined, L("inlineFrame"));
    if (isInsertionPoint || noSelection) {
        radioGraphic.value = true;
        // 未選択時はインラインを無効化（テキスト挿入位置がない）
        if (noSelection) radioInline.enabled = false;
    } else {
        radioInline.value = true;
    }

    // 2カラムレイアウト
    var mainColumnsGroup = dlg.add("group");
    mainColumnsGroup.orientation = "row";
    mainColumnsGroup.alignChildren = ["fill", "top"];

    // 左カラム: フレームサイズ
    var leftColumnGroup = mainColumnsGroup.add("group");
    leftColumnGroup.orientation = "column";
    leftColumnGroup.alignChildren = ["fill", "top"];

    var frameSizePanel = leftColumnGroup.add("panel", undefined, L("frameSizePanel"));
    frameSizePanel.alignChildren = ["fill", "top"];
    frameSizePanel.margins = [15, 20, 15, 10];

    // 幅
    var widthPanel = frameSizePanel.add("panel", undefined, L("widthPanel"));
    widthPanel.alignChildren = ["left", "top"];
    widthPanel.margins = [15, 20, 15, 10];
    // Keep explicit reference for UI symmetry with the other width options.
    // This variable is also used by updateState() for enable/disable control
    // and fallback selection handling, while keeping the three radio options
    // structurally consistent and easier to extend in future updates.
    var radioWidthSelectedText = widthPanel.add("radiobutton", undefined, L("widthText"));
    var radioWidthColumn = widthPanel.add("radiobutton", undefined, L("widthColumn"));
    var radioWidthParentFrame = widthPanel.add("radiobutton", undefined, L("widthFrame"));
    var radioWidthMargin = widthPanel.add("radiobutton", undefined, L("widthMargin"));

    // テキスト未選択時: 「ページのマージン」を自動選択、「選択した文字」をディム
    if (isInsertionPoint || noSelection) {
        radioWidthMargin.value = true;
        radioWidthSelectedText.enabled = false;
        if (noSelection) {
            radioWidthColumn.enabled = false;
            radioWidthParentFrame.enabled = false;
        }
    } else {
        radioWidthColumn.value = true;
    }

    // 高さ（テキスト未選択時のみ表示）
    var heightInput = null;
    var radioHeightLines = null;
    var radioHeightSize = null;
    if (!hasTextBounds) {
        var heightPanel = frameSizePanel.add("panel", undefined, L("heightPanel"));
        heightPanel.alignChildren = ["left", "top"];
        heightPanel.margins = [15, 20, 15, 10];

        radioHeightLines = heightPanel.add("radiobutton", undefined, L("heightLines"));
        radioHeightLines.helpTip = L("heightLinesTip");
        radioHeightSize = heightPanel.add("radiobutton", undefined, L("heightSize"));

        // 入力欄＋単位ラベル（共有）
        var heightInputGroup = heightPanel.add("group");
        heightInputGroup.alignment = ["left", "top"];
        heightInput = heightInputGroup.add("edittext", undefined, noSelection ? "40" : "1");
        if (!noSelection) heightInput.helpTip = L("heightLinesTip");
        heightInput.characters = 8;
        var heightUnitLabel = heightInputGroup.add("statictext", undefined, noSelection ? L("unitMm") : L("unitLines"));

        // 未選択時は行数モード無効（行送り参照不可）
        if (noSelection) {
            radioHeightSize.value = true;
            radioHeightLines.enabled = false;
        } else {
            radioHeightLines.value = true;
        }

        // ラジオボタン切替で入力値と単位を切替
        radioHeightLines.onClick = function () {
            heightInput.text = "1";
            heightUnitLabel.text = L("unitLines");
        };
        radioHeightSize.onClick = function () {
            heightInput.text = "40";
            heightUnitLabel.text = L("unitMm");
        };
    }

    // 右カラム: 段落スタイル、行送り、オブジェクトスタイル、回り込み
    var rightColumnGroup = mainColumnsGroup.add("group");
    rightColumnGroup.orientation = "column";
    rightColumnGroup.alignChildren = ["fill", "top"];

    var textSettingsPanel = rightColumnGroup.add("panel", undefined, L("textSettingsPanel"));
    textSettingsPanel.alignChildren = ["fill", "top"];
    textSettingsPanel.margins = [15, 20, 15, 10];

    var paraStyleLabel = textSettingsPanel.add("statictext", undefined, L("paraStyle"));
    var paraStyleDropdown = textSettingsPanel.add("dropdownlist", undefined, paraStyleNames);
    paraStyleDropdown.selection = defaultParaStyleIndex;

    var autoLeadingCheck = textSettingsPanel.add("checkbox", undefined, L("autoLeading"));
    autoLeadingCheck.value = true;

    var objectSettingsPanel = rightColumnGroup.add("panel", undefined, L("objectSettingsPanel"));
    objectSettingsPanel.alignment = ["fill", "top"];
    objectSettingsPanel.alignChildren = ["fill", "top"];
    objectSettingsPanel.margins = [15, 20, 15, 10];

    objectSettingsPanel.add("statictext", undefined, L("objStyle"));
    var objStyleDropdown = objectSettingsPanel.add("dropdownlist", undefined, objStyleNames);
    objStyleDropdown.selection = noneObjStyleIndex;

    var wrapLabel = objectSettingsPanel.add("statictext", undefined, L("wrap"));
    var wrapDropdown = objectSettingsPanel.add("dropdownlist", undefined, wrapLabels);
    wrapDropdown.selection = 0;

    // ディム表示の更新
    function updateState() {
        var isInline = radioInline.value;
        var isNoneStyle = (objStyleDropdown.selection.index === noneObjStyleIndex);

        // グラフィックフレーム → 挿入行の段落スタイルをディム
        paraStyleLabel.enabled = isInline;
        paraStyleDropdown.enabled = isInline;
        autoLeadingCheck.enabled = isInline;

        // 新しい有効・無効ロジックとラジオボタン選択解除
        var disableSelectedText = !hasTextBounds;
        var disableColumn = noSelection;
        var disableParentFrame = isInline || noSelection;
        var disableMargin = isInline;

        radioWidthSelectedText.enabled = !disableSelectedText;
        radioWidthColumn.enabled = !disableColumn;
        radioWidthParentFrame.enabled = !disableParentFrame;
        radioWidthMargin.enabled = !disableMargin;

        if ((radioWidthSelectedText.value && disableSelectedText) ||
            (radioWidthColumn.value && disableColumn) ||
            (radioWidthParentFrame.value && disableParentFrame) ||
            (radioWidthMargin.value && disableMargin)) {
            if (!disableSelectedText) {
                radioWidthSelectedText.value = true;
            } else if (!disableColumn) {
                radioWidthColumn.value = true;
            } else if (!disableParentFrame) {
                radioWidthParentFrame.value = true;
            } else if (!disableMargin) {
                radioWidthMargin.value = true;
            }
        }

        // インライン → テキストの回り込みをディム
        // グラフィックフレームでもオブジェクトスタイルが「なし」以外ならディム
        wrapLabel.enabled = !isInline && isNoneStyle;
        wrapDropdown.enabled = !isInline && isNoneStyle;
    }
    updateState();
    radioGraphic.onClick = updateState;
    radioInline.onClick = updateState;
    objStyleDropdown.onChange = updateState;

    var btnGroup = dlg.add("group");
    btnGroup.alignment = ["center", "top"];
    btnGroup.margins = [0, 10, 0, 0];
    btnGroup.add("button", undefined, L("cancel"), { name: "cancel" });
    btnGroup.add("button", undefined, L("ok"), { name: "ok" });

    if (dlg.show() !== 1) return;

    var isInline = radioInline.value;
    var selectedObjStyleIndex = objStyleDropdown.selection.index;
    var selectedWrap = wrapModes[wrapDropdown.selection.index];
    var selectedParaStyleIndex = paraStyleDropdown.selection.index;
    var useAutoLeading = autoLeadingCheck.value;

    // フレーム高さを決定
    var frameHeight = height;
    if (radioHeightLines && radioHeightSize) {
        if (radioHeightLines.value) {
            // 行数モード: 1行目の文字サイズ + (残り行数 × 行送り値) で計算
            var lineCount = parseFloat(heightInput.text) || 1;
            if (lineCount < 1) lineCount = 1;
            try {
                var textFrame = isInsertionPoint ? getInsertionPointTextFrame(selection) : selection.parentTextFrames[0];
                if (!textFrame) throw new Error("No parent text frame");
                var story = textFrame.parentStory;
                var ipIndex = isInsertionPoint ? selection.index : selection.characters[0].index;
                var refChar = (ipIndex < story.characters.length)
                    ? story.characters[ipIndex]
                    : story.characters[story.characters.length - 1];
                var pointSizeVal = refChar.pointSize;
                var leadingVal = refChar.leading;
                if (leadingVal === Leading.AUTO) {
                    leadingVal = pointSizeVal * (refChar.autoLeading / 100);
                }
                frameHeight = pointSizeVal + Math.max(0, lineCount - 1) * leadingVal;
            } catch (_) {
                frameHeight = 14 + Math.max(0, lineCount - 1) * 14; // フォールバック
            }
        } else {
            // サイズ指定モード: mm入力をポイントへ変換
            var sizeVal = parseFloat(heightInput.text) || 40;
            frameHeight = sizeVal * 2.834645669;
        }
    }

    // フレーム幅を決定
    var frameWidth = width;
    if (radioWidthColumn.value) {
        try {
            var textFrame = selection.parentTextFrames[0];
            var colWidth = textFrame.textFramePreferences.textColumnFixedWidth;
            if (colWidth <= 0) {
                var textFrameBounds = textFrame.geometricBounds;
                var tfWidth = textFrameBounds[3] - textFrameBounds[1];
                var colCount = textFrame.textFramePreferences.textColumnCount;
                var gutter = textFrame.textFramePreferences.textColumnGutter;
                var insetLeft = textFrame.textFramePreferences.insetSpacing[1];
                var insetRight = textFrame.textFramePreferences.insetSpacing[3];
                colWidth = (tfWidth - insetLeft - insetRight - gutter * (colCount - 1)) / colCount;
            }
            frameWidth = colWidth;
        } catch (_) { }
    } else if (radioWidthParentFrame.value) {
        try {
            var textFrame = selection.parentTextFrames[0];
            var textFrameBounds = textFrame.geometricBounds;
            frameWidth = textFrameBounds[3] - textFrameBounds[1];
        } catch (_) { }
    } else if (radioWidthMargin.value) {
        try {
            var parentPage;
            if (noSelection) {
                parentPage = app.activeWindow.activePage;
            } else if (isInsertionPoint) {
                var insertionPointTextFrame = getInsertionPointTextFrame(selection);
                parentPage = insertionPointTextFrame ? insertionPointTextFrame.parentPage : null;
            } else {
                parentPage = getParentPage(selection);
            }
            if (parentPage) {
                var pageBounds = parentPage.bounds;
                var mp = parentPage.marginPreferences;
                frameWidth = (pageBounds[3] - pageBounds[1]) - mp.left - mp.right;
            }
        } catch (_) { }
    }

    app.doScript(function () {
        if (isInline) {
            // インライン（アンカー付き）
            var textFrame = getInsertionPointTextFrame(selection);
            if (!textFrame) {
                alert(L("parentTextFrameError"));
                return;
            }
            var story = textFrame.parentStory;
            var charIndex = isInsertionPoint ? selection.index : selection.characters[0].index;

            // 元のテキストの段落スタイルを保持
            var originalParaStyle = null;
            var originalPara = null;
            if (!isInsertionPoint) {
                originalParaStyle = selection.characters[0].appliedParagraphStyle;
                originalPara = selection.characters[0].paragraphs[0];
            }

            // 直前の文字が改行かチェック
            var needReturn = true;
            if (charIndex > 0) {
                var prevContent = story.characters[charIndex - 1].contents;
                if (prevContent === "\r") {
                    needReturn = false;
                }
            } else {
                needReturn = false; // テキスト先頭の場合は不要
            }

            if (needReturn) {
                story.insertionPoints[charIndex].contents = "\r";
                charIndex = charIndex + 1;
            }

            // 改行直後のインサーションポイントにアンカー付きフレームを作成
            var anchorIP = story.insertionPoints[charIndex];
            var frameRect = anchorIP.rectangles.add();
            frameRect.geometricBounds = [0, 0, frameHeight, frameWidth];
            frameRect.contentType = ContentType.GRAPHIC_TYPE;
            frameRect.anchoredObjectSettings.anchoredPosition = AnchorPosition.ANCHORED;

            // オブジェクトスタイルを適用
            frameRect.appliedObjectStyle = doc.objectStyles[selectedObjStyleIndex];

            // アンカー付きフレームの後に改行を挿入
            // フレーム挿入で文字が1つ増えているため、charIndex + 1 の位置に改行
            story.insertionPoints[charIndex + 1].contents = "\r";

            // アンカーが属する段落に段落スタイルを適用
            var newPara = anchorIP.paragraphs[0];
            newPara.appliedParagraphStyle = doc.paragraphStyles[selectedParaStyleIndex];

            // 行送りを自動に設定
            if (useAutoLeading) {
                newPara.autoLeading = 100;
                newPara.leading = Leading.AUTO;
            }

            // 元のテキストの段落スタイルを復元
            if (originalPara && originalParaStyle) {
                try {
                    originalPara.appliedParagraphStyle = originalParaStyle;
                } catch (_) { }
            }
        } else {
            // グラフィックフレーム: ページ上に配置
            var parentPage;
            if (noSelection) {
                parentPage = app.activeWindow.activePage;
            } else if (isInsertionPoint) {
                var insertionPointTextFrame = getInsertionPointTextFrame(selection);
                parentPage = insertionPointTextFrame ? insertionPointTextFrame.parentPage : null;
            } else {
                parentPage = getParentPage(selection);
            }
            if (!parentPage) {
                alert(L("parentPageError"));
                return;
            }

            var frameTop, frameLeft, frameRight;
            if (hasTextBounds) {
                frameTop = selectionBounds[0];
                frameLeft = selectionBounds[1];
            } else {
                // InsertionPoint: マージン左上を基準に配置
                var mp = parentPage.marginPreferences;
                var pageBounds = parentPage.bounds;
                frameTop = pageBounds[0] + mp.top;
                frameLeft = pageBounds[1] + mp.left;
            }
            frameRight = frameLeft + frameWidth;
            var frameRect = parentPage.rectangles.add();
            frameRect.geometricBounds = [frameTop, frameLeft, frameTop + frameHeight, frameRight];
            frameRect.contentType = ContentType.GRAPHIC_TYPE;

            // オブジェクトスタイルを適用
            frameRect.appliedObjectStyle = doc.objectStyles[selectedObjStyleIndex];

            // 回り込み（オブジェクトスタイルが「なし」の場合のみ）
            if (selectedObjStyleIndex === noneObjStyleIndex) {
                frameRect.textWrapPreferences.textWrapMode = selectedWrap;
            }
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, L("undoName"));

    function getInsertionPointTextFrame(insertionPointObj) {
        try {
            if (insertionPointObj && insertionPointObj.parentTextFrames && insertionPointObj.parentTextFrames.length > 0) {
                return insertionPointObj.parentTextFrames[0];
            }
        } catch (_) { }
        return null;
    }

    function getParentPage(textObj) {
        try {
            if (textObj.parentTextFrames.length > 0) {
                return textObj.parentTextFrames[0].parentPage;
            }
        } catch (_) { }
        return null;
    }

    function isTextSelection(obj) {
        if (!obj) return false;

        try {
            if (obj.hasOwnProperty("baseline")) {
                return true;
            }
        } catch (_) { }

        try {
            if (obj.hasOwnProperty("characters") && obj.characters && obj.characters.length > 0) {
                return true;
            }
        } catch (_) { }

        try {
            if (obj.constructor && obj.constructor.name) {
                var typeName = String(obj.constructor.name);
                if (
                    typeName === "Text" ||
                    typeName === "Word" ||
                    typeName === "Line" ||
                    typeName === "Paragraph" ||
                    typeName === "TextStyleRange" ||
                    typeName === "Character"
                ) {
                    return true;
                }
            }
        } catch (_) { }

        return false;
    }

    function getTextSelectionBounds(textObj) {
        return getTextSelectionBoundsSingleOutline(textObj) || getTextSelectionBoundsPerCharacter(textObj);
    }

    function getTextSelectionBoundsSingleOutline(textObj) {
        try {
            if (!textObj || !textObj.characters || textObj.characters.length === 0) return null;

            var outlineItems = textObj.createOutlines(false);
            if (!outlineItems || outlineItems.length === 0) return null;

            return getAndRemoveOutlineBounds(outlineItems);
        } catch (_) {
            return null;
        }
    }

    function getTextSelectionBoundsPerCharacter(textObj) {
        try {
            var chars = textObj.characters;
            if (!chars || chars.length === 0) return null;

            var mergedTop = null;
            var mergedLeft = null;
            var mergedBottom = null;
            var mergedRight = null;
            var i, ch, charBounds;

            for (i = 0; i < chars.length; i++) {
                ch = chars[i];

                try {
                    if (ch.contents === "\r" || ch.contents === "\n" || ch.contents === "\u0003") {
                        continue;
                    }
                } catch (_) { }

                charBounds = getCharacterOutlineBounds(ch);
                if (!charBounds) continue;

                if (mergedTop === null || charBounds[0] < mergedTop) mergedTop = charBounds[0];
                if (mergedLeft === null || charBounds[1] < mergedLeft) mergedLeft = charBounds[1];
                if (mergedBottom === null || charBounds[2] > mergedBottom) mergedBottom = charBounds[2];
                if (mergedRight === null || charBounds[3] > mergedRight) mergedRight = charBounds[3];
            }

            if (mergedTop === null) {
                try {
                    return textObj.parentTextFrames[0].geometricBounds;
                } catch (_) {
                    return null;
                }
            }

            return [mergedTop, mergedLeft, mergedBottom, mergedRight];
        } catch (_) {
            return null;
        }
    }

    function getCharacterOutlineBounds(characterObj) {
        try {
            var outlineItems = characterObj.createOutlines(false);
            if (!outlineItems || outlineItems.length === 0) return null;
            return getAndRemoveOutlineBounds(outlineItems);
        } catch (_) {
            return null;
        }
    }

    function getAndRemoveOutlineBounds(outlineItems) {
        var top = null;
        var left = null;
        var bottom = null;
        var right = null;
        var i, itemBounds;

        try {
            for (i = 0; i < outlineItems.length; i++) {
                itemBounds = outlineItems[i].geometricBounds;
                if (top === null || itemBounds[0] < top) top = itemBounds[0];
                if (left === null || itemBounds[1] < left) left = itemBounds[1];
                if (bottom === null || itemBounds[2] > bottom) bottom = itemBounds[2];
                if (right === null || itemBounds[3] > right) right = itemBounds[3];
            }
        } finally {
            for (i = outlineItems.length - 1; i >= 0; i--) {
                try { outlineItems[i].remove(); } catch (_) { }
            }
        }

        if (top === null) return null;
        return [top, left, bottom, right];
    }

})();
