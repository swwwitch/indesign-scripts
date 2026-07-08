#target indesign
/*

### 概要

選択したテキストの「現在の行送り（絶対値）」とフォントサイズから行送り％を段落ごとに逆算し、
それを自動行送り量（％）として設定したうえで、行送りを「自動」に切り替えるスクリプトです。
あわせて行送りの基準位置を「仮想ボディの上/右」に設定します。
ダイアログやパレットは表示せず、実行するとその場で選択中のテキストへ適用します。

対応する選択：
- テキスト編集モードの範囲選択（触れている段落全体が対象）
- 選択ツールでのテキストフレーム選択（グループ内のネストしたフレームも再帰的に処理）

すでに自動行送りの段落は逆算できないためスキップします。
処理後は選択を選び直して文字パネルの表示を更新します。

### Overview

For each paragraph, back-calculates the leading percentage from the selection's current
(absolute) leading and font size, applies it as the paragraph's auto-leading amount, switches
leading to Auto, and sets the leading model to the top/right of the virtual body. No dialog is
shown; it applies to the selection in place.

Supported selections:
- Range selections in text-edit mode (whole touched paragraphs)
- Text frames picked with the Selection tool (recurses into nested frames inside groups)

Paragraphs already set to Auto leading are skipped (nothing to back-calculate). After applying,
the selection is re-selected so the Character panel refreshes its displayed values.

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

(function () {

    var isJa = ($.locale.indexOf("ja") === 0);
    /* 言語に応じた文字列 / Pick a string for the current UI language */
    function t(ja, en) { return isJa ? ja : en; }

    /* テキストとして段落を取り出せる選択オブジェクトの型 / Selection types that expose paragraphs */
    var TEXT_SELECTION_TYPES = {
        Character: true, Word: true, TextStyleRange: true, Paragraph: true,
        Line: true, TextColumn: true, Text: true, InsertionPoint: true, Story: true
    };

    /* 型名を安全に取得（InDesign は typename ではなく constructor.name）
       Safely resolve the constructor name (InDesign has no typename; use constructor.name) */
    function getConstructorName(obj) {
        try { return obj && obj.constructor ? obj.constructor.name : ""; } catch (e) { return ""; }
    }

    /* テキストオブジェクトの段落を対象配列へ追加 / Push a text object's paragraphs into the target list */
    function addParagraphs(textObject, paragraphTargets) {
        try {
            var paragraphs = textObject.paragraphs;
            for (var i = 0; i < paragraphs.length; i++) paragraphTargets.push(paragraphs[i]);
        } catch (e) { }
    }

    /* 選択アイテム1つから対象段落を収集（テキスト選択／テキストフレーム／グループ再帰）
       Collect target paragraphs from one selected item (text selection / text frame / recurse groups) */
    function collectFromItem(item, paragraphTargets) {
        if (!item) return;
        var typeName = getConstructorName(item);

        if (TEXT_SELECTION_TYPES[typeName]) {
            // テキスト編集モード：選択が触れている段落全体を対象 / Text-edit mode: touched paragraphs
            addParagraphs(item, paragraphTargets);

        } else if (typeName === "TextFrame") {
            // 選択ツールでフレーム選択：そのフレームの本文を対象 / Frame selected: its own text
            try {
                if (item.texts.length > 0) addParagraphs(item.texts[0], paragraphTargets);
            } catch (e) { }

        } else if (typeName === "Group") {
            // グループは内部（ネスト含む）の全テキストフレームを対象 / Group: every nested text frame
            try {
                var innerItems = item.allPageItems;
                for (var i = 0; i < innerItems.length; i++) {
                    if (getConstructorName(innerItems[i]) === "TextFrame" && innerItems[i].texts.length > 0) {
                        addParagraphs(innerItems[i].texts[0], paragraphTargets);
                    }
                }
            } catch (e) { }

        } else {
            // 長方形などテキストを保持し得る図形フレーム / Rectangle etc. that may hold text
            try {
                if (item.texts && item.texts.length > 0 && item.texts[0].contents.length > 0) {
                    addParagraphs(item.texts[0], paragraphTargets);
                }
            } catch (e) { }
        }
    }

    /* 1段落へ、その段落の現在の行送り（絶対値）÷サイズから % を逆算して自動行送りを適用
       Apply auto-leading to one paragraph by back-calculating the % from its absolute leading ÷ size */
    function applyAutoLeadingToParagraph(paragraph) {
        try {
            if (paragraph.characters.length === 0) return;
            var firstChar = paragraph.characters[0];
            var sizePt = firstChar.pointSize;
            var leadingValue = firstChar.leading; // 数値 or Leading.AUTO / Number or Leading.AUTO

            // すでに自動行送りなら計算不能なのでスキップ / Skip if already Auto (nothing to back-calculate)
            if (leadingValue === Leading.AUTO) return;
            var leadingPt = Number(leadingValue);
            if (isNaN(sizePt) || sizePt <= 0 || isNaN(leadingPt) || leadingPt <= 0) return;

            var percent = Math.round((leadingPt / sizePt) * 100 * 10) / 10;

            paragraph.autoLeading = percent;   // 自動行送り量（％）/ Auto-leading amount (%)
            paragraph.leading = Leading.AUTO;   // 行送りを自動に切り替え / Switch leading to Auto
            // 基準は仮想ボディの上/右に固定 / Fix the leading basis to the top/right of the virtual body
            paragraph.leadingModel = LeadingModel.LEADING_MODEL_AKI_BELOW;
        } catch (e) { }
    }

    function main() {
        if (app.documents.length === 0) {
            alert(t("ドキュメントを開いてください。", "Please open a document."));
            return;
        }

        var selectionItems = app.selection;
        if (!selectionItems || selectionItems.length === 0) {
            alert(t("テキストが選択されていません。", "No text is selected."));
            return;
        }

        var paragraphTargets = []; // 対象の段落 / Target paragraphs
        for (var i = 0; i < selectionItems.length; i++) collectFromItem(selectionItems[i], paragraphTargets);

        if (paragraphTargets.length === 0) {
            alert(t("テキストが選択されていません。", "No text is selected."));
            return;
        }

        for (var p = 0; p < paragraphTargets.length; p++) applyAutoLeadingToParagraph(paragraphTargets[p]);

        // 文字パネル等の表示を更新するため、いったん選択を解除してから同じ選択を選び直す
        // Deselect then re-select the same objects so the Character panel refreshes its cached values
        try {
            app.select(NothingEnum.NOTHING);
            app.select(selectionItems);
        } catch (e) { }
    }

    main();

})();
