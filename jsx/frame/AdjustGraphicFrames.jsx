#target indesign

/*
 * AdjustGraphicFrames.jsx
 *
 * テキストにアンカーされたグラフィックフレームを集め、フレーム幅・フレームサイズ・画像の縮尺率をまとめて調整します。
 * 詳細は README を参照してください。
 */

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AdjustGraphicFrames";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-06-02";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/AdjustGraphicFrames.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/AdjustGraphicFrames.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// ==============================
// UIレイアウトの共通設定 / Shared UI layout
// ==============================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

/**
 * ウィンドウの共通設定を適用する
 * @param {Window} win 対象ウィンドウ
 * @param {number} [spacing] 要素間隔。省略時は WINDOW_SPACING
 * @returns {void}
 */
function setupWindow(win, spacing) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルの共通設定を適用する
 * @param {Panel} panel 対象パネル
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * 行グループの共通設定を適用する（ボタン列など）
 * @param {Group} group 対象グループ
 * @param {string} [alignment] 配置。省略時は "left"
 * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
 * @returns {void}
 */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 縮尺率の切り捨て単位の選択肢（％）と既定値 / Round-down step choices (%) and the default */
    var ROUND_PRECISION_OPTIONS = [1, 5, 10];
    var DEFAULT_ROUND_PRECISION = 5;

    /* 既定で「切り捨てる」を ON にする / Turn on "Round Down" by default */
    var DEFAULT_ROUND_SCALE = true;

    /* 切り捨て対象を元解像度 72/96/144 ppi の画像に限定する（既定 ON）
       / Limit round-down to images whose actual resolution is 72, 96 or 144 ppi (on by default) */
    var DEFAULT_ROUND_ONLY_72_96_144 = true;
    var ROUND_TARGET_PPI = [72, 96, 144];

    /* 縦横比の崩れ（横スケール≠縦スケール）を同率とみなす許容差（％）
       / Tolerance (%) for treating horizontal and vertical scales as already uniform */
    var ASPECT_RATIO_TOLERANCE = 0.1;

    /* 既定で「フレームを内容に合わせる」を ON にする / Turn on "Fit Frame to Content" by default */
    var DEFAULT_FIT_FRAME_TO_CONTENT = true;

    /* 既定で「インライン画像を文字サイズに合わせる」を ON にする / Turn on inline-image matching by default */
    var DEFAULT_ADJUST_INLINE_IMAGE_IN_TEXT = true;

    /* 既定で「合わせ先より大きい場合のみ」を ON にする / Turn on "Only when wider than target" by default */
    var DEFAULT_WIDTH_ONLY_IF_LARGER = true;

    /* 既定で切り捨て後の「再フィット」を ON にする / Re-fit after rounding down by default */
    var DEFAULT_REFIT_AFTER_ROUND = true;

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
    var currentLanguage = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "アンカー付きオブジェクトのフレーム調整", en: "Adjust Anchored Object Frames" }
        },
        panel: {
            target: { ja: "対象", en: "Target" },
            width: { ja: "フレーム幅", en: "Frame Width" },
            frameSize: { ja: "フレームサイズ", en: "Frame Size" },
            scale: { ja: "縮尺率", en: "Scale" }
        },
        target: {
            document: { ja: "ドキュメント", en: "Document" },
            story: { ja: "ストーリー", en: "Story" },
            selection: { ja: "選択範囲", en: "Selection" }
        },
        fit: {
            frameToContent: { ja: "フレームを内容に合わせる", en: "Fit Frame to Content" },
            adjustInlineImageInText: { ja: "インライン画像を文字サイズに合わせる", en: "Match Inline Images to Text Size" }
        },
        width: {
            keep: { ja: "変更しない", en: "Don't Change" },
            fitToParent: { ja: "親フレームに合わせる", en: "Fit to Parent Frame" },
            onlyIfLarger: { ja: "親フレーム／マージンより大きい場合のみ", en: "Only When Wider Than Target" },
            fitToMargin: { ja: "マージンに合わせる", en: "Fit to Margins" }
        },
        scale: {
            round: { ja: "縮尺率を切り捨てる", en: "Round Scale Down" },
            precision: { ja: "単位", en: "Step" },
            refit: { ja: "調整後にフレームを内容へ合わせる", en: "Fit Frame to Content After Adjusting" },
            only7296144: { ja: "スクショのみ（72{slash}96{slash}144 ppi の画像）", en: "Screenshots only (72{slash}96{slash}144 ppi images)" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tip: {
            targetDocument: { ja: "ドキュメント内の、テキストにアンカーされたすべてのグラフィックフレーム", en: "All text-anchored graphic frames in the document" },
            targetStory: { ja: "選択中のテキストフレーム／テキスト範囲／挿入点が属するストーリーにアンカーされたグラフィックフレーム", en: "Graphic frames anchored in the story of the selected text frame / text range / insertion point" },
            targetSelection: { ja: "選択中の、テキストにアンカーされたグラフィックフレーム（グループは展開。テキスト選択時はそのストーリーを対象）", en: "Selected text-anchored graphic frames (groups expanded; a text selection targets its story)" },
            fitFrameToContent: { ja: "フレームの大きさを、配置された内容のバウンディングボックスに合わせます。", en: "Resize the frame to match the placed content's bounds." },
            adjustInlineImageInText: { ja: "真のインライン配置の画像の高さを、同じ段落内の文字サイズに合わせます（行揃え／カスタム配置は対象外）。", en: "Match true inline image height to the text size in the same paragraph (above-line / custom positions excluded)." },
            widthKeep: { ja: "フレームの幅は変更しません。", en: "Leave the frame width unchanged." },
            fitToParent: { ja: "アンカーされたフレームの幅を、親テキストフレームの内寸に合わせます。", en: "Match the anchored frame width to the parent text frame's inner width." },
            onlyIfLarger: { ja: "オンにすると、合わせ先（親フレーム／マージン）より広いフレームだけを調整します（縮小のみ）。オフの場合は、小さいフレームも合わせ先の幅に広げます。", en: "When on, only frames wider than the target (parent frame / margins) are adjusted (shrink only). When off, smaller frames are also expanded to the target width." },
            fitToMargin: { ja: "フレームをページのマージン（版面）幅に合わせて配置します。", en: "Fit frames to the page margin live-area width." },
            round: { ja: "画像の拡大縮小率を、指定した単位で切り捨てます。", en: "Round the image scale down to the selected step." },
            precision: { ja: "切り捨ての刻み幅です。小さいほど元の倍率に近づきます。", en: "Round-down step. Smaller values stay closer to the original scale." },
            refit: { ja: "縮尺率を切り捨てた後、フレームを内容に合わせ直します。", en: "After rounding the scale down, fit the frame to the content again." },
            only7296144: { ja: "元解像度が 72／96／144 ppi の画像だけを切り捨て対象にします。", en: "Limit round-down to images whose actual resolution is 72, 96, or 144 ppi." }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noTextFrame: { ja: "おっと!!!\nテキストフレーム、テキスト範囲、挿入点、またはアンカーオブジェクトマーカーを選択してください。", en: "Oops!!!\nPlease select a text frame, text range, insertion point, or anchored object marker." },
            noSelection: { ja: "フレームを選択してください。", en: "Please select one or more frames." },
            noFrames: { ja: "対象となるフレームが見つかりませんでした。", en: "No target frames were found." },
            done: { ja: "{count} 個のフレームを調整しました。", en: "Adjusted {count} frame(s)." },
            doneNone: { ja: "調整対象のフレームはありませんでした。", en: "No frames needed adjustment." }
        }
    };

    /**
     * ドット区切りキーでラベルを取得する（{slash} は言語別のスラッシュに置換）
     * @param {string} key 例: "dialog.title"
     * @returns {string} 現在の言語のラベル文字列。見つからない場合はキーをそのまま返す
     */
    function getLabel(key) {
        var pathParts = key.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathParts.length; i++) {
            if (labelNode === undefined || labelNode === null) return key;
            labelNode = labelNode[pathParts[i]];
        }
        if (labelNode === undefined || labelNode === null) return key;

        /* 現在の言語が無ければ英語→キーの順にフォールバック / Fall back to English, then the key */
        var labelString = labelNode[currentLanguage];
        if (labelString === undefined || labelString === null) labelString = labelNode.en;
        if (labelString === undefined || labelString === null) return key;

        var slash = (currentLanguage === "ja") ? "／" : "/";
        return ("" + labelString).replace(/\{slash\}/g, slash);
    }

    /**
     * ラベル内の {count} を件数で置き換える
     * @param {string} key 置換対象のラベルキー
     * @param {number} count 埋め込む件数
     * @returns {string} 置換後の文字列
     */
    function labelWithCount(key, count) {
        return getLabel(key).replace(/\{count\}/g, count);
    }


    main();

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * 検証・ダイアログ・収集・適用の流れを制御する
     * @returns {void}
     */
    function main() {
        /* ドキュメントの有無を確認 / Require an open document */
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var targetDocument = app.activeDocument;

        /* 選択中のテキストフレームを取得（初期選択の判定用）/ Get the selected text frame (for default state) */
        var selectedTextFrame = getSelectedTextFrame();

        /* ダイアログで設定を取得 / Get settings from dialog */
        var settings = showOptionsDialog(selectedTextFrame !== null);
        if (settings === null) return; // キャンセル / Cancelled

        /* 対象に応じてフレームを収集 / Collect frames by target */
        var targetFrames = collectFrames(targetDocument, settings.target, selectedTextFrame);
        if (targetFrames === null) return; // 選択不足などで中断 / Aborted (e.g. nothing selected)
        if (targetFrames.length === 0) {
            alert(getLabel("alert.noFrames"));
            return;
        }

        /* 取り消しをひとまとめに / Apply in a single undo step */
        var processedCount = 0;
        app.doScript(
            function () { processedCount = applyToFrames(targetFrames, settings); },
            ScriptLanguage.JAVASCRIPT,
            undefined,
            UndoModes.ENTIRE_SCRIPT,
            getLabel("dialog.title")
        );

        /* 完了メッセージ（0 件は専用文言）/ Completion message (dedicated text for zero) */
        alert(processedCount > 0 ? labelWithCount("alert.done", processedCount) : getLabel("alert.doneNone"));
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * グループの共通設定を適用する（orientation は呼び出し側で指定）
     * @param {Group} group 対象グループ
     * @param {string} [orientation] 並び方向。省略時は "column"
     * @param {number} [spacing] 要素間隔。省略時は PANEL_SPACING
     * @returns {void}
     */
    function setupGroup(group, orientation, spacing) {
        group.orientation = orientation || "column";
        group.alignChildren = ["fill", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 調整内容を指定するダイアログを表示する
     * @param {boolean} hasTextFrameSelection テキストフレームが選択されているか
     * @returns {object|null} 設定内容。キャンセル時は null
     */
    function showOptionsDialog(hasTextFrameSelection) {
        var dialog = new Window("dialog", getLabel("dialog.title") + "  " + SCRIPT_VERSION);
        setupWindow(dialog, 10);

        /* 対象はカラム貫通（全幅）/ Target spans the full width */
        var targetControls = buildTargetPanel(dialog, hasTextFrameSelection);

        /* 1 カラム：フレームの幅 → フレームの大きさ → 縮尺率 / One column: Width → Frame size → Scale */
        var widthControls = buildWidthPanel(dialog);
        var frameToContentControls = buildFrameToContentPanel(dialog);
        var scaleControls = buildScalePanel(dialog);

        /* ボタン / Buttons (Mac: Cancel → OK) */
        var dialogButtons = dialog.add("group");
        dialogButtons.alignment = "right";
        dialogButtons.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        dialogButtons.add("button", undefined, "OK", { name: "ok" });

        if (dialog.show() !== 1) return null;

        return collectSettings(targetControls, frameToContentControls, widthControls, scaleControls);
    }

    /**
     * 対象範囲のパネルを組み立てる
     * @param {Window} dialog 対象のダイアログ
     * @param {boolean} hasTextFrameSelection テキストフレームが選択されているか
     * @returns {object} パネル内のコントロール
     */
    function buildTargetPanel(parent, hasTextFrameSelection) {
        var panel = parent.add("panel", undefined, getLabel("panel.target"));
        setupPanel(panel);

        /* ラジオを横並び＋左右中央に / Radios in a row, centered horizontally */
        var radioGroup = panel.add("group");
        setupGroup(radioGroup, "row");
        radioGroup.alignChildren = ["center", "center"];
        radioGroup.alignment = ["center", "top"];

        var documentRadio = radioGroup.add("radiobutton", undefined, getLabel("target.document"));
        var storyRadio = radioGroup.add("radiobutton", undefined, getLabel("target.story"));
        var selectionRadio = radioGroup.add("radiobutton", undefined, getLabel("target.selection"));
        documentRadio.helpTip = getLabel("tip.targetDocument");
        storyRadio.helpTip = getLabel("tip.targetStory");
        selectionRadio.helpTip = getLabel("tip.targetSelection");

        /* 初期選択は常にドキュメント / Default is always Document */
        documentRadio.value = true;
        /* ストーリーはテキストフレーム選択時のみ選択可能 / Story is selectable only when a text frame is selected */
        storyRadio.enabled = hasTextFrameSelection;
        /* 選択範囲は何か選択しているときのみ選択可能 / Selection is selectable only when something is selected */
        selectionRadio.enabled = (app.selection.length > 0);

        return { documentRadio: documentRadio, storyRadio: storyRadio, selectionRadio: selectionRadio };
    }


    /**
     * フレームサイズのパネルを組み立てる
     * @param {Window} dialog 対象のダイアログ
     * @returns {object} パネル内のコントロール
     */
    function buildFrameToContentPanel(parent) {
        var panel = parent.add("panel", undefined, getLabel("panel.frameSize"));
        setupPanel(panel);

        var frameToContentCheckbox = panel.add("checkbox", undefined, getLabel("fit.frameToContent"));
        frameToContentCheckbox.value = DEFAULT_FIT_FRAME_TO_CONTENT;
        frameToContentCheckbox.helpTip = getLabel("tip.fitFrameToContent");

        var adjustInlineImageInTextCheckbox = panel.add("checkbox", undefined, getLabel("fit.adjustInlineImageInText"));
        adjustInlineImageInTextCheckbox.value = DEFAULT_ADJUST_INLINE_IMAGE_IN_TEXT;
        adjustInlineImageInTextCheckbox.helpTip = getLabel("tip.adjustInlineImageInText");

        return {
            frameToContentCheckbox: frameToContentCheckbox,
            adjustInlineImageInTextCheckbox: adjustInlineImageInTextCheckbox
        };
    }

    /**
     * フレーム幅のパネルを組み立てる
     * @param {Window} dialog 対象のダイアログ
     * @returns {object} パネル内のコントロール
     */
    function buildWidthPanel(parent) {
        var panel = parent.add("panel", undefined, getLabel("panel.width"));
        setupPanel(panel);

        var keepRadio = panel.add("radiobutton", undefined, getLabel("width.keep"));
        var fitToParentRadio = panel.add("radiobutton", undefined, getLabel("width.fitToParent"));
        var fitToMarginRadio = panel.add("radiobutton", undefined, getLabel("width.fitToMargin"));
        var onlyIfLargerCheckbox = panel.add("checkbox", undefined, getLabel("width.onlyIfLarger"));
        fitToParentRadio.value = true;
        onlyIfLargerCheckbox.value = DEFAULT_WIDTH_ONLY_IF_LARGER;
        keepRadio.helpTip = getLabel("tip.widthKeep");
        fitToParentRadio.helpTip = getLabel("tip.fitToParent");
        fitToMarginRadio.helpTip = getLabel("tip.fitToMargin");
        onlyIfLargerCheckbox.helpTip = getLabel("tip.onlyIfLarger");

        /**
         * 幅の指定に応じて「合わせ先より大きい場合のみ」の有効／無効を切り替える
         * @returns {void}
         */
        function syncOnlyIfLargerEnabled() {
            onlyIfLargerCheckbox.enabled = fitToParentRadio.value || fitToMarginRadio.value;
        }
        keepRadio.onClick = syncOnlyIfLargerEnabled;
        fitToParentRadio.onClick = syncOnlyIfLargerEnabled;
        fitToMarginRadio.onClick = syncOnlyIfLargerEnabled;
        syncOnlyIfLargerEnabled();

        return {
            keepRadio: keepRadio,
            fitToParentRadio: fitToParentRadio,
            fitToMarginRadio: fitToMarginRadio,
            onlyIfLargerCheckbox: onlyIfLargerCheckbox
        };
    }

    /**
     * 縮尺率のパネルを組み立てる
     * @param {Window} dialog 対象のダイアログ
     * @returns {object} パネル内のコントロール
     */
    function buildScalePanel(parent) {
        var panel = parent.add("panel", undefined, getLabel("panel.scale"));
        setupPanel(panel);

        var roundCheckbox = panel.add("checkbox", undefined, getLabel("scale.round"));
        roundCheckbox.value = DEFAULT_ROUND_SCALE;
        roundCheckbox.helpTip = getLabel("tip.round");

        var precisionGroup = panel.add("group");
        /* コロンは日本語は全角、英語は半角 / Colon: full-width JA, half-width EN */
        var precisionLabel = precisionGroup.add("statictext", undefined, getLabel("scale.precision") + (currentLanguage === "ja" ? "：" : ":"));
        precisionLabel.helpTip = getLabel("tip.precision");

        var precisionRadios = [];
        for (var i = 0; i < ROUND_PRECISION_OPTIONS.length; i++) {
            var radio = precisionGroup.add("radiobutton", undefined, ROUND_PRECISION_OPTIONS[i] + "%");
            if (ROUND_PRECISION_OPTIONS[i] === DEFAULT_ROUND_PRECISION) radio.value = true;
            radio.helpTip = getLabel("tip.precision");
            precisionRadios.push(radio);
        }

        var refitCheckbox = panel.add("checkbox", undefined, getLabel("scale.refit"));
        refitCheckbox.value = DEFAULT_REFIT_AFTER_ROUND;
        refitCheckbox.helpTip = getLabel("tip.refit");

        var only7296144Checkbox = panel.add("checkbox", undefined, getLabel("scale.only7296144"));
        only7296144Checkbox.value = DEFAULT_ROUND_ONLY_72_96_144;
        only7296144Checkbox.helpTip = getLabel("tip.only7296144");

        /**
         * 切り捨ての ON/OFF に応じて関連項目を切り替える
         * @returns {void}
         */
        function syncRoundEnabled() {
            precisionGroup.enabled = roundCheckbox.value;
            refitCheckbox.enabled = roundCheckbox.value;
            only7296144Checkbox.enabled = roundCheckbox.value;
        }
        roundCheckbox.onClick = syncRoundEnabled;
        syncRoundEnabled();

        return {
            roundCheckbox: roundCheckbox,
            precisionRadios: precisionRadios,
            refitCheckbox: refitCheckbox,
            only7296144Checkbox: only7296144Checkbox
        };
    }

    /**
     * ダイアログの入力内容を設定オブジェクトにまとめる
     * @returns {object} 適用に使う設定
     */
    function collectSettings(targetControls, frameToContentControls, widthControls, scaleControls) {
        /* 対象 / Target */
        var target = "document";
        if (targetControls.storyRadio.value) target = "story";
        else if (targetControls.selectionRadio.value) target = "selection";


        /* 丸め精度 / Rounding precision */
        var roundPrecision = DEFAULT_ROUND_PRECISION;
        for (var j = 0; j < scaleControls.precisionRadios.length; j++) {
            if (scaleControls.precisionRadios[j].value) {
                roundPrecision = ROUND_PRECISION_OPTIONS[j];
                break;
            }
        }

        /* フレーム幅の調整モード / Frame width mode */
        var widthMode = "keep";
        if (widthControls.fitToParentRadio.value) widthMode = "parent";
        else if (widthControls.fitToMarginRadio.value) widthMode = "margin";

        return {
            target: target,
            fitFrameToContent: frameToContentControls.frameToContentCheckbox.value,
            adjustInlineImageInText: frameToContentControls.adjustInlineImageInTextCheckbox.value,
            widthMode: widthMode,
            widthOnlyIfLarger: widthControls.onlyIfLargerCheckbox.value,
            roundScale: scaleControls.roundCheckbox.value,
            roundPrecision: roundPrecision,
            refitAfterRound: scaleControls.refitCheckbox.value,
            roundOnly7296144: scaleControls.only7296144Checkbox.value
        };
    }

    /**
     * 選択からテキストフレームを取り出す
     * @returns {TextFrame|null} テキストフレーム。取得できない場合は null
     */
    function getSelectedTextFrame() {
        if (app.selection.length === 0) return null;
        var selectedItem = app.selection[0];
        try {
            if (selectedItem.constructor.name === "TextFrame") return selectedItem;
            if (selectedItem.parentTextFrames && selectedItem.parentTextFrames.length > 0) return selectedItem.parentTextFrames[0];
            if (selectedItem.parent && selectedItem.parent.parentTextFrames && selectedItem.parent.parentTextFrames.length > 0) return selectedItem.parent.parentTextFrames[0];
            /* ^~a$（アンカーオブジェクトマーカーのみ選択）にも対応
               Support selections consisting only of an anchored-object marker */
            if (selectedItem.constructor.name === "Character" && selectedItem.contents === "\uFFFC") {
                if (selectedItem.parentTextFrames && selectedItem.parentTextFrames.length > 0) {
                    return selectedItem.parentTextFrames[0];
                }
            }
        } catch (e) { }
        return null;
    }

    // =========================================
    // フレーム収集 / Frame collection
    // =========================================

    /**
     * 対象範囲に応じて調整するフレームを集める
     * @param {Document} targetDocument 対象ドキュメント
     * @param {string} target 対象範囲を表す識別子
     * @param {TextFrame} selectedTextFrame 選択中のテキストフレーム
     * @returns {Array<PageItem>|null} 対象フレーム。中断時は null
     */
    function collectFrames(targetDocument, target, selectedTextFrame) {
        if (target === "story") {
            if (selectedTextFrame === null) {
                alert(getLabel("alert.noTextFrame"));
                return null;
            }
            return collectFramesFromStory(targetDocument, selectedTextFrame.parentStory);
        }
        if (target === "selection") {
            if (app.selection.length === 0) {
                alert(getLabel("alert.noSelection"));
                return null;
            }
            return collectFramesFromSelection(app.selection);
        }
        return collectFramesFromDocument(targetDocument);
    }

    /**
     * 編集できるフレームかどうかを判定する
     * @param {PageItem} frame 対象のフレーム
     * @returns {boolean} 編集できるなら true
     */
    function isEditableFrame(item) {
        if (!isFrameItem(item)) return false;
        /* テキストにアンカーされたフレームのみ対象（独立配置は対象外）/ Only frames anchored into text (free-floating frames are excluded) */
        if (getOwningStory(item) === null) return false;
        try {
            /* マスターページ上の項目は対象外 / Skip items on master spreads */
            var page = item.parentPage;
            if (page !== null && page.parent.constructor.name === "MasterSpread") return false;
            /* 非表示・ロックは対象外 / Skip hidden or locked items */
            if (item.visible === false) return false;
            if (item.locked === true) return false;
            /* ロックレイヤー・非表示レイヤーは対象外 / Skip items on locked or hidden layers */
            var itemLayer = item.itemLayer;
            if (itemLayer !== undefined && (itemLayer.locked === true || itemLayer.visible === false)) return false;
        } catch (e) {
            return false;
        }
        return true;
    }

    /**
     * グラフィックフレームとして扱える種別かを判定する
     * @param {PageItem} pageItem 対象のオブジェクト
     * @returns {boolean} 対象なら true
     */
    function isFrameItem(item) {
        var typeName = item.constructor.name;
        return typeName === "Rectangle" || typeName === "Oval" || typeName === "Polygon";
    }

    /**
     * オブジェクトが属するストーリーを取得する
     * @param {PageItem} pageItem 対象のオブジェクト
     * @returns {Story|null} ストーリー。取得できない場合は null
     */
    function getOwningStory(item) {
        try {
            /* storyOffset から親ストーリーを取得 / Read parent story from storyOffset */
            var anchorPoint = item.storyOffset;
            if (anchorPoint && anchorPoint.isValid) {
                if (anchorPoint.parentStory && anchorPoint.parentStory.isValid) return anchorPoint.parentStory;
                if (anchorPoint.parent && anchorPoint.parent.parentStory && anchorPoint.parent.parentStory.isValid) return anchorPoint.parent.parentStory;
            }
        } catch (e) { }

        try {
            /* インラインアンカーでは parent が Character になることがある / Inline anchored items may have a Character parent */
            var parent = item.parent;
            if (parent && parent.parentStory && parent.parentStory.isValid) return parent.parentStory;
            if (parent && parent.parent && parent.parent.parentStory && parent.parent.parentStory.isValid) return parent.parent.parentStory;
        } catch (err) { }

        return null;
    }

    /**
     * ストーリーにアンカーされたフレームを集める
     * @param {Story} story 対象のストーリー
     * @returns {Array<PageItem>} 対象フレーム
     */
    function collectFramesFromStory(targetDocument, story) {
        var allItems = targetDocument.allPageItems;
        var collectedFrames = [];
        for (var i = 0; i < allItems.length; i++) {
            if (!isEditableFrame(allItems[i])) continue;
            var owningStory = getOwningStory(allItems[i]);
            if (owningStory !== null && owningStory.id === story.id) collectedFrames.push(allItems[i]);
        }
        return collectedFrames;
    }

    /**
     * ドキュメント全体からアンカーされたフレームを集める
     * @param {Document} targetDocument 対象ドキュメント
     * @returns {Array<PageItem>} 対象フレーム
     */
    function collectFramesFromDocument(targetDocument) {
        var allItems = targetDocument.allPageItems;
        var collectedFrames = [];
        for (var i = 0; i < allItems.length; i++) {
            if (isEditableFrame(allItems[i])) collectedFrames.push(allItems[i]);
        }
        return collectedFrames;
    }

    /**
     * 選択からアンカーされたフレームを集める
     * @returns {Array<PageItem>|null} 対象フレーム。中断時は null
     */
    function collectFramesFromSelection(selection) {
        var collectedFrames = [];
        var seenIds = {};
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            var typeName = item.constructor.name;

            if (typeName === "TextFrame") {
                appendFrames(collectedFrames, seenIds, collectFramesFromStory(app.activeDocument, item.parentStory));
            } else if (typeName === "Character" && item.contents === "\uFFFC") {
                appendFrames(collectedFrames, seenIds, collectFramesFromStory(app.activeDocument, item.parentStory));
            } else if (typeName === "Group") {
                appendFrames(collectedFrames, seenIds, collectFramesFromGroup(item));
            } else if (isEditableFrame(item)) {
                appendFrame(collectedFrames, seenIds, item);
            } else {
                /* テキスト範囲・挿入点などのテキスト選択は、そのストーリーを対象に
                   Text range / insertion point → collect from its story */
                var selectionStory = getSelectionStory(item);
                if (selectionStory !== null) {
                    appendFrames(collectedFrames, seenIds, collectFramesFromStory(app.activeDocument, selectionStory));
                }
            }
        }
        return collectedFrames;
    }

    /**
     * 選択からストーリーを取り出す
     * @returns {Story|null} ストーリー。取得できない場合は null
     */
    function getSelectionStory(item) {
        try {
            if (item.parentStory && item.parentStory.isValid) return item.parentStory;
        } catch (e) { }
        return null;
    }

    /**
     * グループ内のフレームを再帰的に集める
     * @param {Group} group 対象のグループ
     * @param {Array<PageItem>} frames 収集先の配列
     * @returns {void}
     */
    function collectFramesFromGroup(group) {
        var collectedFrames = [];
        var groupItems = group.allPageItems;
        for (var i = 0; i < groupItems.length; i++) {
            if (isEditableFrame(groupItems[i])) collectedFrames.push(groupItems[i]);
        }
        return collectedFrames;
    }

    /**
     * 重複を避けながらフレームの配列を追加する
     * @param {Array<PageItem>} frames 収集先の配列
     * @param {Array<PageItem>} newFrames 追加するフレーム
     * @returns {void}
     */
    function appendFrames(targetFrames, seenIds, sourceFrames) {
        for (var i = 0; i < sourceFrames.length; i++) {
            appendFrame(targetFrames, seenIds, sourceFrames[i]);
        }
    }

    /**
     * 重複を避けながらフレームを 1 つ追加する
     * @param {Array<PageItem>} frames 収集先の配列
     * @param {PageItem} frame 追加するフレーム
     * @returns {void}
     */
    function appendFrame(targetFrames, seenIds, frame) {
        var key = "" + frame.id;
        if (seenIds[key]) return;
        seenIds[key] = true;
        targetFrames.push(frame);
    }

    // =========================================
    // 実行 / Apply
    // =========================================

    /**
     * 収集したフレームへ設定を適用する
     * @param {Array<PageItem>} frames 対象フレーム
     * @param {object} settings 適用する設定
     * @returns {number} 実際に調整した件数
     */
    function applyToFrames(targetFrames, settings) {
        var processedCount = 0;
        for (var i = 0; i < targetFrames.length; i++) {
            var frame = targetFrames[i];
            /* 変化検出のため処理前の状態を記録 / Snapshot the state before processing, to detect real changes */
            var beforeSignature = frameSignature(frame);
            applyToFrame(frame, settings);
            /* インライン画像を前後の文字サイズに合わせる / Match inline graphics to surrounding text size */
            if (settings.adjustInlineImageInText) matchInlineHeightToText(frame);
            /* フレームか内容が実際に変化したものだけを数える / Count only frames whose geometry or content actually changed */
            if (frameSignature(frame) !== beforeSignature) processedCount++;
        }
        return processedCount;
    }

    /**
     * 変更の有無を比べるためのフレームの状態を文字列化する
     * @param {PageItem} frame 対象のフレーム
     * @returns {string} 状態を表す文字列
     */
    function frameSignature(frame) {
        var parts = [];
        try {
            var bounds = frame.geometricBounds; // [y1, x1, y2, x2]
            parts.push(roundForCompare(bounds[0]), roundForCompare(bounds[1]), roundForCompare(bounds[2]), roundForCompare(bounds[3]));
        } catch (e) {
            parts.push("nb");
        }
        var graphic = getSingleImage(frame);
        if (graphic !== null) {
            try {
                parts.push(roundForCompare(graphic.horizontalScale), roundForCompare(graphic.verticalScale));
            } catch (eScale) { parts.push("ns"); }
            try {
                /* 画像自身の位置（フィットで動くため変化検出に含める）/ The image's own position (moves on fit, so include it) */
                var graphicBounds = graphic.geometricBounds; // [y1, x1, y2, x2]
                parts.push(roundForCompare(graphicBounds[0]), roundForCompare(graphicBounds[1]), roundForCompare(graphicBounds[2]), roundForCompare(graphicBounds[3]));
            } catch (eBounds) { parts.push("ng"); }
        }
        return parts.join(",");
    }

    /**
     * 比較用に数値を丸める
     * @param {number} value 対象の数値
     * @returns {number} 丸めた数値
     */
    function roundForCompare(value) {
        if (typeof value !== "number" || !isFinite(value)) return "x";
        return Math.round(value * 10000) / 10000;
    }

    /**
     * 1 つのフレームへ設定を適用する
     * @param {PageItem} frame 対象のフレーム
     * @param {object} settings 適用する設定
     * @returns {void}
     */
    function applyToFrame(frame, settings) {
        /* 縦横比の崩れた画像を縦横同率に正す（設定に依らず常に実行）/ Always correct distorted (non-uniform) image scaling */
        correctImageAspectRatio(frame);

        /* 前後に文字がある真のインライン画像は、ここでは何もしない（縮尺率切り捨て・内容フィット・フレーム幅調整は行わない）。
           高さ合わせは applyToFrames 側が「インライン画像を文字サイズに合わせる」ON のときだけ実行する
           True inline images with surrounding text are left untouched here (no round-down, fit-to-content, or width change).
           Height matching is handled by the caller only when "Match Inline Images to Text Size" is on */
        if (getInlineSurroundingPointSize(frame) > 0) return;

        /* 縮尺率を切り捨てる / Round the image scale down */
        if (settings.roundScale) {
            roundImageScale(frame, settings.roundPrecision, settings.refitAfterRound, settings.roundOnly7296144);
        }

        /* フレームを内容に合わせる / Fit frame to content */
        if (settings.fitFrameToContent) {
            safeFit(frame, FitOptions.frameToContent);
        }

        /* フレームの幅の調整 / Adjust frame width */
        if (settings.widthMode === "parent") {
            fitWidthToParentFrame(frame, settings.widthOnlyIfLarger);
        } else if (settings.widthMode === "margin") {
            fitWidthToMargin(frame, settings.widthOnlyIfLarger);
        }
    }

    /**
     * フィット処理を例外を握りつぶして実行する
     * @param {PageItem} frame 対象のフレーム
     * @param {FitOptions} fitOption フィット方法
     * @returns {void}
     */
    function safeFit(frame, fitOption) {
        try {
            frame.fit(fitOption);
            return true;
        } catch (e) {
            return false; // ロック・空フレームなど / Locked, empty, etc.
        }
    }

    /**
     * フレーム内の画像が 1 点だけならそれを返す
     * @param {PageItem} frame 対象のフレーム
     * @returns {Graphic|null} 画像。該当しない場合は null
     */
    function getSingleImage(frame) {
        try {
            if (frame.allGraphics && frame.allGraphics.length === 1) {
                var graphic = frame.allGraphics[0];
                var typeName = graphic.constructor.name;
                if (typeName === "Image" || typeName === "PDF" || typeName === "EPS") return graphic;
            }
        } catch (e) { }
        return null;
    }

    /**
     * 指定した刻み幅で切り捨てる
     * @param {number} value 対象の数値
     * @param {number} step 刻み幅
     * @returns {number} 切り捨てた数値
     */
    function floorToStep(value, step) {
        return Math.floor(value / step) * step;
    }

    /**
     * 切り捨て対象の解像度かどうかを判定する
     * @param {Graphic} image 対象の画像
     * @returns {boolean} 対象なら true
     */
    function isTargetResolution(graphic) {
        try {
            var actualPpi = graphic.actualPpi; // [horizontal, vertical]
            if (!actualPpi || actualPpi.length < 2) return false;
            for (var i = 0; i < ROUND_TARGET_PPI.length; i++) {
                if (Math.round(actualPpi[0]) === ROUND_TARGET_PPI[i] ||
                    Math.round(actualPpi[1]) === ROUND_TARGET_PPI[i]) {
                    return true;
                }
            }
        } catch (e) { }
        return false;
    }

    /**
     * 画像の縮尺率を指定単位で切り捨てる
     * @param {PageItem} frame 対象のフレーム
     * @param {object} settings 適用する設定
     * @returns {boolean} 変更したら true
     */
    function roundImageScale(frame, precision, refit, onlyTargetPpi) {
        var graphic = getSingleImage(frame);
        if (graphic === null) return false;
        /* 解像度フィルタ：72/96/144 ppi 以外はスキップ / Resolution filter: skip non-72/96/144 ppi images */
        if (onlyTargetPpi && !isTargetResolution(graphic)) return false;
        try {
            /* 横スケールを基準に切り捨て、縦にも同じ値を適用（縦横同率に揃える）/ Floor horizontal scale, apply the same value to vertical */
            var flooredScale = floorToStep(graphic.horizontalScale, precision);
            /* 0 など不正値になる場合は最小ステップに留める（エラー 11268 回避）/ Clamp to one step if flooring hits 0, etc. (avoids error 11268) */
            if (flooredScale < precision) flooredScale = precision;
            if (!isValidScale(flooredScale)) return false;

            graphic.horizontalScale = flooredScale;
            graphic.verticalScale = flooredScale;
            if (refit) safeFit(frame, FitOptions.frameToContent); // 再フィット / Re-fit
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 崩れた縦横比を横スケール基準に揃える
     * @param {PageItem} frame 対象のフレーム
     * @returns {boolean} 変更したら true
     */
    function correctImageAspectRatio(frame) {
        var graphic = getSingleImage(frame);
        if (graphic === null) return false;
        try {
            var hScale = graphic.horizontalScale;
            var vScale = graphic.verticalScale;
            if (!isValidScale(hScale) || !isValidScale(vScale)) return false;
            /* 既に同率（許容差内）なら触らない / Already uniform within tolerance */
            if (Math.abs(hScale - vScale) <= ASPECT_RATIO_TOLERANCE) return false;
            /* 横スケールを基準に縦を合わせる / Match vertical scale to the horizontal scale */
            graphic.verticalScale = hScale;
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * アンカー元の親テキストフレームを取得する
     * @param {PageItem} frame 対象のフレーム
     * @returns {TextFrame|null} 親テキストフレーム。取得できない場合は null
     */
    function getParentTextFrame(frame) {
        try {
            var anchorCharacter = getAnchorCharacter(frame);
            if (anchorCharacter !== null && anchorCharacter.parentTextFrames && anchorCharacter.parentTextFrames.length > 0) {
                return anchorCharacter.parentTextFrames[0];
            }
        } catch (e) { }

        try {
            var parent = frame.parent;
            if (parent && parent.parentTextFrames && parent.parentTextFrames.length > 0) {
                return parent.parentTextFrames[0];
            }
            if (parent && parent.parent && parent.parent.parentTextFrames && parent.parent.parentTextFrames.length > 0) {
                return parent.parent.parentTextFrames[0];
            }
        } catch (err) { }

        return null;
    }

    /**
     * フレーム幅を親テキストフレームの内寸に合わせる
     * @param {PageItem} frame 対象のフレーム
     * @param {object} settings 適用する設定
     * @returns {boolean} 変更したら true
     */
    function fitWidthToParentFrame(frame, onlyIfLarger) {
        try {
            var parentTextFrame = getParentTextFrame(frame);
            if (parentTextFrame === null) return false;

            var parentBounds = parentTextFrame.geometricBounds; // [y1, x1, y2, x2]
            var parentWidth = parentBounds[3] - parentBounds[1];

            /* テキストフレームの左右インセットを控除（取得不可なら 0）/ Subtract left/right text insets (0 if unavailable) */
            var insets = getTextFrameInsets(parentTextFrame);
            var contentWidth = parentWidth - insets.left - insets.right;
            if (contentWidth <= 0) return false; // 不正な幅は触らない / Skip invalid widths

            var bounds = frame.geometricBounds; // [y1, x1, y2, x2]
            var currentWidth = bounds[3] - bounds[1];
            /* 親の内寸以下のフレームはそのまま / Leave frames that already fit within the parent */
            if (onlyIfLarger && currentWidth <= contentWidth) return false;

            /* 左端を保持して親の内寸幅へ。広い→収める / 狭い→いっぱいに流し込む（マージンと同じ挙動）
               Keep the left edge, snap to the parent width. Wider→fit inside / narrower→fill (same as margin) */
            applyWidthFit(frame, bounds[1], bounds[1] + contentWidth, currentWidth);
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 指定した幅にフレームを合わせ、内容を収め直す
     * @param {PageItem} frame 対象のフレーム
     * @param {number} targetWidth 目標の幅
     * @param {object} settings 適用する設定
     * @returns {boolean} 変更したら true
     */
    function applyWidthFit(frame, left, right, originalWidth) {
        var bounds = frame.geometricBounds; // [y1, x1, y2, x2]
        frame.geometricBounds = [bounds[0], left, bounds[2], right];

        var targetWidth = right - left;
        if (originalWidth > targetWidth) {
            safeFit(frame, FitOptions.proportionally);   // 広い：内容を内側に収める / Wider: fit inside
        } else {
            safeFit(frame, FitOptions.fillProportionally); // 狭い：いっぱいに流し込む / Narrower: fill
        }
        safeFit(frame, FitOptions.frameToContent);
    }

    /**
     * テキストフレームの内側余白を取得する
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @returns {object} 各辺の余白
     */
    function getTextFrameInsets(textFrame) {
        var result = { left: 0, right: 0 };
        try {
            var insetPrefs = textFrame.textFramePreferences;
            var spacing = insetPrefs.insetSpacing; // 数値 or [top, left, bottom, right]
            if (spacing instanceof Array) {
                result.left = spacing[1];
                result.right = spacing[3];
            } else if (typeof spacing === "number") {
                result.left = spacing;
                result.right = spacing;
            }
        } catch (e) { }
        return result;
    }

    /**
     * フレーム幅をページのマージン幅に合わせる
     * @param {PageItem} frame 対象のフレーム
     * @param {object} settings 適用する設定
     * @returns {boolean} 変更したら true
     */
    function fitWidthToMargin(frame, onlyIfLarger) {
        try {
            var page = frame.parentPage;
            if (page === null) return false;

            /* マージン量はページ単位で取得（見開き左右でも正しい）/ Read margins per page (correct on facing pages) */
            var margins = page.marginPreferences;
            var pageBounds = page.bounds; // [y1, x1, y2, x2]
            var left = pageBounds[1] + margins.left;
            var right = pageBounds[3] - margins.right;
            var marginWidth = right - left;
            if (marginWidth <= 0) return false; // 不正な版面幅は触らない / Skip invalid live-area widths

            /* 元のフレーム幅を記録してから版面幅へ / Remember the original width, then snap to the live area */
            var bounds = frame.geometricBounds; // [y1, x1, y2, x2]
            var originalWidth = bounds[3] - bounds[1];
            /* 版面幅以下のフレームはそのまま（縮小のみ）/ Leave frames within the live area (shrink only) */
            if (onlyIfLarger && originalWidth <= marginWidth) return false;

            /* 版面幅へ。広い→収める / 狭い→いっぱいに流し込む / Snap to the live area. Wider→fit / narrower→fill */
            applyWidthFit(frame, left, right, originalWidth);
            return true;
        } catch (e) {
            return false;
        }
    }


    /**
     * インライン画像の高さを同じ段落の文字サイズに合わせる
     * @param {PageItem} frame 対象のフレーム
     * @returns {boolean} 変更したら true
     */
    function matchInlineHeightToText(frame) {
        try {
            /* アンカー文字を取得（インラインなら frame.parent が Character）/ Get the anchor character (frame.parent is a Character when inline) */
            var anchorCharacter = getAnchorCharacter(frame);
            if (anchorCharacter === null) return false; // インラインでなければ対象外 / Not inline

            var paragraph = anchorCharacter.paragraphs[0];
            var targetPoint = getSurroundingTextPointSize(paragraph);
            if (targetPoint <= 0) return false; // 前後に文字が無い（画像のみの段落）/ No surrounding text

            var graphic = getSingleImage(frame);
            if (graphic === null) return false; // 単一画像でなければ対象外 / Only single-image frames

            /* まずフレームを画像実寸に合わせ、現在の高さ（pt）を取得 / Fit to content first, then read the height in points */
            safeFit(frame, FitOptions.frameToContent);
            var currentPoint = getFrameHeightPoint(frame);
            if (!isFinite(currentPoint) || currentPoint <= 0) return false;

            /* 目標高さになるよう画像を比例拡大縮小 / Scale the image proportionally to reach the target height */
            var factor = targetPoint / currentPoint;
            var currentHScale = graphic.horizontalScale;
            var currentVScale = graphic.verticalScale;
            if (!isValidScale(currentHScale) || !isValidScale(currentVScale)) return false;

            var newHScale = currentHScale * factor;
            var newVScale = currentVScale * factor;
            /* 不正・範囲外のスケール値は適用しない（エラー 11268 回避）/ Skip invalid/out-of-range scales (avoids error 11268) */
            if (!isValidScale(newHScale) || !isValidScale(newVScale)) return false;

            graphic.horizontalScale = newHScale;
            graphic.verticalScale = newVScale;

            /* 縮小後の画像にフレームを合わせ直す / Re-fit the frame to the scaled image */
            safeFit(frame, FitOptions.frameToContent);
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * スケール値として有効かどうかを判定する
     * @param {number} value スケール値
     * @returns {boolean} 有効なら true
     */
    function isValidScale(scale) {
        return typeof scale === "number" && isFinite(scale) && scale > 0 && scale <= 1000000;
    }

    /**
     * インライン画像の前後にある文字サイズを取得する
     * @param {PageItem} frame 対象のフレーム
     * @returns {number} 文字サイズ（pt）
     */
    function getInlineSurroundingPointSize(frame) {
        var anchorCharacter = getAnchorCharacter(frame);
        if (anchorCharacter === null) return 0; // インラインでなければ 0 / Not inline
        try {
            return getSurroundingTextPointSize(anchorCharacter.paragraphs[0]);
        } catch (e) {
            return 0;
        }
    }

    /**
     * 真のインライン配置かどうかを判定する
     * @param {PageItem} frame 対象のフレーム
     * @returns {boolean} インライン配置なら true
     */
    function isInlineAnchored(frame) {
        try {
            var anchoredSettings = frame.anchoredObjectSettings;
            return anchoredSettings && anchoredSettings.anchoredPosition === AnchorPosition.INLINE_POSITION;
        } catch (e) {
            return false;
        }
    }

    /**
     * アンカーオブジェクトを表す文字を取得する
     * @param {PageItem} frame 対象のフレーム
     * @returns {Character|null} アンカー文字。取得できない場合は null
     */
    function getAnchorCharacter(frame) {
        try {
            if (!isInlineAnchored(frame)) return null; // 行揃え・カスタム配置は対象外 / Skip above-line / custom positions
            var parent = frame.parent;
            if (parent !== undefined && parent !== null && parent.constructor.name === "Character") {
                return parent;
            }
        } catch (e) { }
        return null;
    }

    /**
     * アンカー文字の前後にある本文の文字サイズを求める
     * @param {Character} anchorCharacter アンカー文字
     * @returns {number} 文字サイズ（pt）
     */
    function getSurroundingTextPointSize(paragraph) {
        var chars = paragraph.characters.everyItem().getElements();
        var maxSize = 0;
        for (var i = 0; i < chars.length; i++) {
            var content = chars[i].contents;
            if (content === "\r" || content === "\n") continue;              // 改行 / line breaks
            if (content === "\uFFFC") continue;                              // アンカーオブジェクトマーカー / anchored object marker
            if (content === " " || content === "\u3000" || content === "\t") continue; // 空白 / spaces
            try {
                var pointSize = chars[i].pointSize;
                if (pointSize !== NothingEnum.NOTHING) {
                    pointSize = Number(pointSize);
                    if (pointSize > maxSize) maxSize = pointSize;
                }
            } catch (e) { }
        }
        return maxSize;
    }

    /**
     * フレームの高さをポイントで取得する
     * @param {PageItem} frame 対象のフレーム
     * @returns {number} 高さ（pt）
     */
    function getFrameHeightPoint(frame) {
        var view = app.activeDocument.viewPreferences;
        var oldH = view.horizontalMeasurementUnits;
        var oldV = view.verticalMeasurementUnits;
        try {
            view.horizontalMeasurementUnits = MeasurementUnits.POINTS;
            view.verticalMeasurementUnits = MeasurementUnits.POINTS;
            var bounds = frame.geometricBounds; // [y1, x1, y2, x2]（ポイント / points）
            return bounds[2] - bounds[0];
        } catch (e) {
            return 0;
        } finally {
            view.horizontalMeasurementUnits = oldH;
            view.verticalMeasurementUnits = oldV;
        }
    }

})();
