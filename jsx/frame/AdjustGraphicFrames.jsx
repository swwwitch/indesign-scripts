#target indesign

/*
 
# アンカー付きオブジェクトのフレーム調整 / Adjust Anchored Object Frames
 
## 概要
ドキュメント・ストーリー・選択範囲のいずれかから、テキストにアンカーされたグラフィックフレーム
（長方形／楕円／多角形）を集め、「フレームの大きさ（幅・サイズ）」と「フレーム内に配置された画像の
縮尺率（拡大率）」を調整する。具体的には、縦横比の補正／フレーム幅／フレームサイズ／縮尺率の切り捨てを
まとめて適用する。
※調整対象はあくまでフレームと、フレーム内の配置（縮尺率・位置）であり、画像ファイルそのものは変更しない。
※テキストにアンカーされていない独立配置のフレーム、およびテキストフレームは対象外。
 
## 対象（ラジオボタン）
- ドキュメント: ドキュメント内の、テキストにアンカーされたすべてのグラフィックフレーム
- ストーリー: 選択中テキストフレーム、または ^~a$（アンカーオブジェクトマーカー）のストーリーにアンカーされたグラフィックフレーム
  ※「ストーリー」は選択がテキストフレーム／テキスト範囲／挿入点のときのみ選択可能（それ以外はディム表示）
- 選択範囲: 選択中の、テキストにアンカーされたグラフィックフレーム
  ※グループは展開。テキストフレーム／テキスト範囲／挿入点を選んだ場合はそのストーリーを対象にする。
  ※何も選択していないときはディム表示。重複選択でも同一フレームは一度だけ処理する。

## フレーム幅（ラジオ）
- 変更しない / 親フレームに合わせる / マージンに合わせる
- 合わせ先より大きい場合のみ（独立チェックボックス、既定オン）
  ※「親フレームに合わせる」は、テキストフレーム内にアンカーされたフレームにのみ作用する。
  ※「マージンに合わせる」はフレームをページのマージン（版面）幅に合わせて配置する。
  ※どちらも、合わせ先より広い→内容を縦横比で収める／狭い→いっぱいに流し込む、で内容を収め直す。
  ※「合わせ先より大きい場合のみ」は「親フレームに合わせる」「マージンに合わせる」の両方に作用する。
    オンのときは合わせ先より広いフレームだけを縮小し、オフのときは小さいフレームも合わせ先の幅へ広げる。

## フレームサイズ（チェックボックス）
- フレームを内容に合わせる: フレームの大きさを、配置された内容のバウンディングボックスに合わせる。
- インライン画像を文字サイズに合わせる: 真のインライン配置のフレームの高さが同じ段落内の文字サイズに
  なるよう、フレーム内の画像の縮尺率を調整する（行揃え／カスタム配置のアンカーは対象外）。

## 縮尺率の切り捨て
- 縮尺率を切り捨てる（オン時）: 画像（Image / PDF / EPS）を1点だけ含むフレームについて、
  フレーム内の配置の拡大縮小率を指定単位（1% / 5% / 10%）で切り捨てる。
  任意で 72／96／144 ppi の画像のみ・切り捨て後の再フィットを指定できる。

## 縦横比の補正（UIなし・常時実行）
- フレーム内に配置された画像の横スケールと縦スケールが食い違う（縦横比が崩れた）場合、
  横スケールを基準に縦を合わせて同率に正す。
  他の調整より前に実行するため、後続の比例拡縮・再フィットは正しい比率を保つ。

## 補足
- 対象はテキストにアンカーされたグラフィックフレーム（長方形／楕円／多角形）のみ。
  独立配置のフレームとテキストフレームは処理しない。
- 前後に文字がある真のインライン画像は、文字サイズへの高さ合わせを優先し、
  縮尺率の切り捨てとフレーム幅の調整は行わない（縦横比の補正は行う）。
- 全処理はひとつの取り消し単位（Cmd+Z 一回）にまとまる。
 
*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.0.0";

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [15, 20, 15, 10];
    var PANEL_SPACING = 8;

    /* 縮尺率の切り捨て単位の選択肢（％）/ Round-down step choices (%) */
    var ROUND_PRECISION_OPTIONS = [1, 5, 10];
    var DEFAULT_ROUND_PRECISION = 5;

    /* 既定で「切り捨てる」を ON にする / Turn on "Round Down" by default */
    var DEFAULT_ROUND_SCALE = true;

    /* 切り捨て対象を元解像度 72/96/144 ppi の画像に限定する（既定 ON）/ Limit round-down to images whose actual resolution is 72/96/144 ppi (on by default) */
    var DEFAULT_ROUND_ONLY_72_96_144 = true;
    var ROUND_TARGET_PPI = [72, 96, 144];

    /* 縦横比の崩れ（横スケール≠縦スケール）を同率とみなす許容差（％）/ Tolerance (%) for treating horizontal/vertical scales as already uniform */
    var ASPECT_RATIO_TOLERANCE = 0.1;

    /* 既定で「フレームを内容に合わせる」を ON にする / Turn on "Frame to Content" by default */
    var DEFAULT_FIT_FRAME_TO_CONTENT = true;

    /* 既定で「テキスト内のインライン画像を調整」を ON にする / Turn on "Adjust inline images in text" by default */
    var DEFAULT_ADJUST_INLINE_IMAGE_IN_TEXT = true;

    /* 既定で「合わせ先より大きい場合のみ」を ON にする / Turn on "Only when wider than target" by default */
    var DEFAULT_WIDTH_ONLY_IF_LARGER = true;

    /* 既定で切り捨て後の「再フィット」を ON にする / Re-fit after rounding down by default */
    var DEFAULT_REFIT_AFTER_ROUND = true;


    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 言語判定 / Detect UI language */
    var currentLanguage = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* ラベル定義（カテゴリ分け）/ Label definitions (grouped by category) */
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

    /* ラベル取得（ドット区切りキー、{slash} を言語別スラッシュへ置換）/ Get label (dot-path key, replace {slash} with locale slash).
       キー不明・言語欠落時はキー文字列をそのまま返す / Falls back to the raw key when missing */
    function L(key) {
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

    /* {count} を件数で差し替えたラベル / Label with {count} replaced by the given count */
    function labelWithCount(key, count) {
        return L(key).replace(/\{count\}/g, count);
    }


    main();

    // =========================================
    // メイン / Main
    // =========================================

    /* 全体の流れを制御（検証 → ダイアログ → 収集 → undo 単位で適用）/ Orchestrate the flow (validate → dialog → collect → apply in one undo step) */
    function main() {
        /* ドキュメントの有無を確認 / Require an open document */
        if (app.documents.length === 0) {
            alert(L("alert.noDocument"));
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
            alert(L("alert.noFrames"));
            return;
        }

        /* 取り消しをひとまとめに / Apply in a single undo step */
        var processedCount = 0;
        app.doScript(
            function () { processedCount = applyToFrames(targetFrames, settings); },
            ScriptLanguage.JAVASCRIPT,
            undefined,
            UndoModes.ENTIRE_SCRIPT,
            L("dialog.title")
        );

        /* 完了メッセージ（0 件は専用文言）/ Completion message (dedicated text for zero) */
        alert(processedCount > 0 ? labelWithCount("alert.done", processedCount) : L("alert.doneNone"));
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* グループの共通設定（orientation は呼び出し側で指定）/ Apply shared group layout (orientation passed in) */
    function setupGroup(group, orientation, spacing) {
        group.orientation = orientation || "column";
        group.alignChildren = ["fill", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* 設定ダイアログを表示し、選択結果を返す / Show the settings dialog and return the result.
       戻り値 / Returns: settings オブジェクト | null(キャンセル) */
    function showOptionsDialog(hasTextFrameSelection) {
        var dialog = new Window("dialog", L("dialog.title") + "  " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.margins = 15;
        dialog.spacing = 10;

        /* 対象はカラム貫通（全幅）/ Target spans the full width */
        var targetControls = buildTargetPanel(dialog, hasTextFrameSelection);

        /* 1 カラム：フレームの幅 → フレームの大きさ → 縮尺率 / One column: Width → Frame size → Scale */
        var widthControls = buildWidthPanel(dialog);
        var frameToContentControls = buildFrameToContentPanel(dialog);
        var scaleControls = buildScalePanel(dialog);

        /* ボタン / Buttons (Mac: Cancel → OK) */
        var dialogButtons = dialog.add("group");
        dialogButtons.alignment = "right";
        dialogButtons.add("button", undefined, L("button.cancel"), { name: "cancel" });
        dialogButtons.add("button", undefined, "OK", { name: "ok" });

        if (dialog.show() !== 1) return null;

        return collectSettings(targetControls, frameToContentControls, widthControls, scaleControls);
    }

    /* 対象パネル（ドキュメント / ストーリー / 選択範囲）/ Target panel */
    function buildTargetPanel(parent, hasTextFrameSelection) {
        var panel = parent.add("panel", undefined, L("panel.target"));
        setupPanel(panel);

        /* ラジオを横並び＋左右中央に / Radios in a row, centered horizontally */
        var radioGroup = panel.add("group");
        setupGroup(radioGroup, "row");
        radioGroup.alignChildren = ["center", "center"];
        radioGroup.alignment = ["center", "top"];

        var documentRadio = radioGroup.add("radiobutton", undefined, L("target.document"));
        var storyRadio = radioGroup.add("radiobutton", undefined, L("target.story"));
        var selectionRadio = radioGroup.add("radiobutton", undefined, L("target.selection"));
        documentRadio.helpTip = L("tip.targetDocument");
        storyRadio.helpTip = L("tip.targetStory");
        selectionRadio.helpTip = L("tip.targetSelection");

        /* 初期選択は常にドキュメント / Default is always Document */
        documentRadio.value = true;
        /* ストーリーはテキストフレーム選択時のみ選択可能 / Story is selectable only when a text frame is selected */
        storyRadio.enabled = hasTextFrameSelection;
        /* 選択範囲は何か選択しているときのみ選択可能 / Selection is selectable only when something is selected */
        selectionRadio.enabled = (app.selection.length > 0);

        return { documentRadio: documentRadio, storyRadio: storyRadio, selectionRadio: selectionRadio };
    }


    /* フレームの大きさの調整パネル / Frame size adjustment panel */
    function buildFrameToContentPanel(parent) {
        var panel = parent.add("panel", undefined, L("panel.frameSize"));
        setupPanel(panel);

        var frameToContentCheckbox = panel.add("checkbox", undefined, L("fit.frameToContent"));
        frameToContentCheckbox.value = DEFAULT_FIT_FRAME_TO_CONTENT;
        frameToContentCheckbox.helpTip = L("tip.fitFrameToContent");

        var adjustInlineImageInTextCheckbox = panel.add("checkbox", undefined, L("fit.adjustInlineImageInText"));
        adjustInlineImageInTextCheckbox.value = DEFAULT_ADJUST_INLINE_IMAGE_IN_TEXT;
        adjustInlineImageInTextCheckbox.helpTip = L("tip.adjustInlineImageInText");

        return {
            frameToContentCheckbox: frameToContentCheckbox,
            adjustInlineImageInTextCheckbox: adjustInlineImageInTextCheckbox
        };
    }

    /* フレーム幅パネル（ラジオ＋合わせ先より大きい場合のみチェック）/ Frame width panel (radios + only-when-wider checkbox) */
    function buildWidthPanel(parent) {
        var panel = parent.add("panel", undefined, L("panel.width"));
        setupPanel(panel);

        var keepRadio = panel.add("radiobutton", undefined, L("width.keep"));
        var fitToParentRadio = panel.add("radiobutton", undefined, L("width.fitToParent"));
        var fitToMarginRadio = panel.add("radiobutton", undefined, L("width.fitToMargin"));
        var onlyIfLargerCheckbox = panel.add("checkbox", undefined, L("width.onlyIfLarger"));
        fitToParentRadio.value = true;
        onlyIfLargerCheckbox.value = DEFAULT_WIDTH_ONLY_IF_LARGER;
        keepRadio.helpTip = L("tip.widthKeep");
        fitToParentRadio.helpTip = L("tip.fitToParent");
        fitToMarginRadio.helpTip = L("tip.fitToMargin");
        onlyIfLargerCheckbox.helpTip = L("tip.onlyIfLarger");

        /* 「合わせ先より大きい場合のみ」は親フレーム／マージンのどちらかを選んだときに有効 / Enabled for either parent or margin */
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

    /* 縮尺率パネル（切り捨てチェック＋単位ラジオ＋再フィット）/ Scale panel */
    function buildScalePanel(parent) {
        var panel = parent.add("panel", undefined, L("panel.scale"));
        setupPanel(panel);

        var roundCheckbox = panel.add("checkbox", undefined, L("scale.round"));
        roundCheckbox.value = DEFAULT_ROUND_SCALE;
        roundCheckbox.helpTip = L("tip.round");

        var precisionGroup = panel.add("group");
        /* コロンは日本語は全角、英語は半角 / Colon: full-width JA, half-width EN */
        var precisionLabel = precisionGroup.add("statictext", undefined, L("scale.precision") + (currentLanguage === "ja" ? "：" : ":"));
        precisionLabel.helpTip = L("tip.precision");

        var precisionRadios = [];
        for (var i = 0; i < ROUND_PRECISION_OPTIONS.length; i++) {
            var radio = precisionGroup.add("radiobutton", undefined, ROUND_PRECISION_OPTIONS[i] + "%");
            if (ROUND_PRECISION_OPTIONS[i] === DEFAULT_ROUND_PRECISION) radio.value = true;
            radio.helpTip = L("tip.precision");
            precisionRadios.push(radio);
        }

        var refitCheckbox = panel.add("checkbox", undefined, L("scale.refit"));
        refitCheckbox.value = DEFAULT_REFIT_AFTER_ROUND;
        refitCheckbox.helpTip = L("tip.refit");

        var only7296144Checkbox = panel.add("checkbox", undefined, L("scale.only7296144"));
        only7296144Checkbox.value = DEFAULT_ROUND_ONLY_72_96_144;
        only7296144Checkbox.helpTip = L("tip.only7296144");

        /* 「縮尺率を切り捨てる」OFF 時は単位・再フィット・解像度限定を無効化 / Disable step, re-fit & PPI filter while "Round Scale Down" is off */
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

    /* ダイアログの入力値を設定オブジェクトへ変換 / Convert dialog inputs into a settings object */
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

    /* 選択からテキストフレームを取得（テキスト選択・挿入点にも対応）/ Get a text frame from the selection (supports text and insertion point selections) */
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

    /* 対象に応じたフレーム配列を返す（中断時は null）/ Return frames for the target (null if aborted) */
    function collectFrames(targetDocument, target, selectedTextFrame) {
        if (target === "story") {
            if (selectedTextFrame === null) {
                alert(L("alert.noTextFrame"));
                return null;
            }
            return collectFramesFromStory(targetDocument, selectedTextFrame.parentStory);
        }
        if (target === "selection") {
            if (app.selection.length === 0) {
                alert(L("alert.noSelection"));
                return null;
            }
            return collectFramesFromSelection(app.selection);
        }
        return collectFramesFromDocument(targetDocument);
    }

    /* 処理対象となるグラフィックフレームか（型・アンカー・マスター・非表示・ロックを判定）
       Whether an item is an editable graphic frame (checks type, anchored, master, hidden, locked) */
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

    /* フレームとして扱える page item か（テキストフレームは対象外）/ Whether a page item is treated as a frame (text frames excluded) */
    function isFrameItem(item) {
        var typeName = item.constructor.name;
        return typeName === "Rectangle" || typeName === "Oval" || typeName === "Polygon";
    }

    /* page item がアンカーされているストーリーを返す（なければ null）/ Return the story an item is anchored in (or null) */
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

    /* 指定ストーリーにアンカーされたグラフィックフレーム / Graphic frames anchored in the given story */
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

    /* ドキュメント内の全フレーム（マスター・非表示・ロックは除外）/ All editable frames in a document */
    function collectFramesFromDocument(targetDocument) {
        var allItems = targetDocument.allPageItems;
        var collectedFrames = [];
        for (var i = 0; i < allItems.length; i++) {
            if (isEditableFrame(allItems[i])) collectedFrames.push(allItems[i]);
        }
        return collectedFrames;
    }

    /* 選択範囲のフレーム（グループは展開。テキストフレーム選択時は同一ストーリー内を対象）/ Frames in the selection (groups expanded; text frames collect their story).
       テキストフレーム複数選択やフレーム重複選択でも、同一フレームは一度だけ収集する / De-duplicates by id so overlapping selections collect each frame once */
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

    /* テキスト選択（テキスト範囲・挿入点・文字など）の所属ストーリーを返す（なければ null）
       Return the story of a text selection (text range / insertion point / character), or null */
    function getSelectionStory(item) {
        try {
            if (item.parentStory && item.parentStory.isValid) return item.parentStory;
        } catch (e) { }
        return null;
    }

    /* グループ内の編集可能フレーム / Editable frames inside a group */
    function collectFramesFromGroup(group) {
        var collectedFrames = [];
        var groupItems = group.allPageItems;
        for (var i = 0; i < groupItems.length; i++) {
            if (isEditableFrame(groupItems[i])) collectedFrames.push(groupItems[i]);
        }
        return collectedFrames;
    }

    /* フレーム配列を id 重複を避けて追加する / Append frames, skipping ids already collected */
    function appendFrames(targetFrames, seenIds, sourceFrames) {
        for (var i = 0; i < sourceFrames.length; i++) {
            appendFrame(targetFrames, seenIds, sourceFrames[i]);
        }
    }

    /* フレームを id 重複を避けて追加する / Append a single frame unless its id was already collected */
    function appendFrame(targetFrames, seenIds, frame) {
        var key = "" + frame.id;
        if (seenIds[key]) return;
        seenIds[key] = true;
        targetFrames.push(frame);
    }

    // =========================================
    // 実行 / Apply
    // =========================================

    /* 全フレームへ設定を適用し、実際に変化したフレーム数を返す / Apply settings to all frames and return the count of frames that actually changed */
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

    /* フレームの状態（フレーム枠＋単一画像の縮尺・位置）を表す文字列を返す。処理前後で比較して実際の変化を検出する
       Return a signature of the frame's state (frame bounds + the single image's scale & position) for before/after change detection */
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

    /* 微小な浮動小数の揺れを無視するため一定桁で丸める / Round to a fixed precision so sub-unit float jitter is ignored */
    function roundForCompare(value) {
        if (typeof value !== "number" || !isFinite(value)) return "x";
        return Math.round(value * 10000) / 10000;
    }

    /* 1 フレームへ設定を適用（縦横比補正 → 縮尺率切り捨て → フレームを内容に合わせる → 幅）
       Apply settings to one frame (aspect-fix → round-down → frame-to-content → width).
       変化の有無は applyToFrames 側で前後比較して判定する / The caller detects whether anything changed via before/after comparison */
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

    /* フレームを安全にフィット（失敗しても全体を止めない）/ Fit a frame safely (one failure must not abort the run) */
    function safeFit(frame, fitOption) {
        try {
            frame.fit(fitOption);
            return true;
        } catch (e) {
            return false; // ロック・空フレームなど / Locked, empty, etc.
        }
    }

    /* フレーム内の単一画像を返す（なければ null）/ Return the single image in a frame (or null) */
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

    /* 値を指定ステップで切り捨てる / Floor a value down to the given step */
    function floorToStep(value, step) {
        return Math.floor(value / step) * step;
    }

    /* 画像の元解像度（actualPpi）が 72/96/144 ppi のいずれかか / Whether the image's actual resolution is one of 72/96/144 ppi.
       縦横どちらかが一致すれば対象（拡大縮小で端数が出ても切り捨て前の実体で判定）/ Matches if either axis equals a target PPI */
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

    /* 画像の拡大縮小率を同一比率に切り捨て、必要なら再フィット / Round an image's scale down to a single ratio, then re-fit if requested.
       onlyTargetPpi が true のとき、元解像度が 72/96/144 ppi の画像のみを対象とする / When onlyTargetPpi, only 72/96/144 ppi images are processed */
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

    /* 縦横比の崩れた画像（横スケール≠縦スケール）を縦横同率に正す。横スケールを基準に縦を合わせる。
       設定に依らず常に実行する / Correct a distorted image (non-uniform scaling) to a uniform ratio, matching vertical to horizontal. Always runs. */
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

    /* アンカー元のテキストフレームを返す（なければ null）/ Return the anchoring text frame (or null) */
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

    /* フレーム幅を親テキストフレームの内寸（インセット控除後）に合わせる / Match a frame's width to its parent text frame's content width (insets removed).
       onlyIfLarger が true のときは、親の内寸より広いフレームだけを縮める / When onlyIfLarger, only frames wider than the parent are shrunk */
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

    /* フレームを指定の左右幅に合わせ、内容を収め直す共通処理 / Snap a frame to the given left/right width, then re-fit its content.
       元の幅が目標より広い→内容を縦横比で収める / 狭い→いっぱいに流し込む。最後にフレームを内容に合わせる
       Wider than target→fit proportionally / narrower→fill proportionally, then fit frame to content */
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

    /* テキストフレームの左右インセット量を返す（取得不可なら 0）/ Return a text frame's left/right insets (0 if unavailable) */
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

    /* フレーム幅をページのマージン（版面）幅に合わせ、内容を収め直す / Fit a frame to the page margin width, then re-fit its content.
       版面より広い場合と狭い場合で収め方を変える（参考ロジック準拠）/ Branch by whether the frame is wider than the live area.
       onlyIfLarger が true のときは、版面幅より広いフレームだけを縮める / When onlyIfLarger, only frames wider than the live area are shrunk */
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


    /* インライン画像の高さを前後の文字サイズに合わせる / Match an inline graphic's height to the surrounding character size.
       文字にアンカーされたフレームのみ対象。フレームではなく画像自体を拡大縮小して高さを合わせる。
       Only anchored inline frames; scales the image itself (not just the frame) so its height equals the font size. */
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

    /* InDesign が受け付けるスケール値か（％）。有限かつ正で、おおよその上限内 / Whether a scale (%) is acceptable to InDesign: finite, positive, within bounds */
    function isValidScale(scale) {
        return typeof scale === "number" && isFinite(scale) && scale > 0 && scale <= 1000000;
    }

    /* インライン画像で前後に文字がある場合の目標文字サイズ（pt）を返す。無ければ 0
       Return the surrounding text size (pt) for an inline image, or 0 when there is none */
    function getInlineSurroundingPointSize(frame) {
        var anchorCharacter = getAnchorCharacter(frame);
        if (anchorCharacter === null) return 0; // インラインでなければ 0 / Not inline
        try {
            return getSurroundingTextPointSize(anchorCharacter.paragraphs[0]);
        } catch (e) {
            return 0;
        }
    }

    /* フレームが真のインライン配置のアンカーオブジェクトか（行揃え・カスタム配置は除外）
       Whether the frame is a true inline-positioned anchored object (above-line / custom positions excluded) */
    function isInlineAnchored(frame) {
        try {
            var anchoredSettings = frame.anchoredObjectSettings;
            return anchoredSettings && anchoredSettings.anchoredPosition === AnchorPosition.INLINE_POSITION;
        } catch (e) {
            return false;
        }
    }

    /* フレームがアンカーされている文字を返す（真のインラインでなければ null）/ Return the character a frame is anchored into (null unless truly inline) */
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

    /* 段落内の画像・空白を除いた文字の最大ポイントサイズ / Largest point size among non-anchor, non-space characters */
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

    /* フレームの高さをポイントで取得（単位に依存しないよう一時的にポイントへ切替）
       Read a frame's height in points (switches ruler units to points temporarily for unit-safety) */
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
