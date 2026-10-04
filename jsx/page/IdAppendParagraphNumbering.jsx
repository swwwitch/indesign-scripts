#target indesign

/*

### 概要

同じ段落スタイルで同じテキストが繰り返す段落の末尾に、連番を付けたり外したりします。重複は直近の親見出しごとに判定し、範囲は選択範囲・ストーリー・ドキュメントから選べます。

詳細は README を参照してください。
https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAppendParagraphNumbering.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nc96549bb60f9

### Overview

Adds or removes sequential numbers at the end of paragraphs that repeat the same text with the same paragraph style. Duplicates are grouped by the nearest parent heading, and the scope can be the selection, a story, or the whole document.

See the README for details.
https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAppendParagraphNumbering.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdAppendParagraphNumbering";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdAppendParagraphNumbering.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdAppendParagraphNumbering.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nc96549bb60f9"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 段落スタイル名と見出しレベルの対応 / Map of paragraph style names to heading levels */
    var HEADING_LEVEL_MAP = {
        "Heading 1": 1, "h1": 1,
        "Heading 2": 2, "h2": 2,
        "Heading 3": 3, "h3": 3,
        "Heading 4": 4, "h4": 4,
        "Heading 5": 5, "h5": 5,
        "Heading 6": 6, "h6": 6
    };

    /* 見出しとして扱う最大レベルと、見出し以外に与えるレベル / Deepest heading level, and the level for non-headings */
    var MAX_HEADING_LEVEL = 6;
    var NON_HEADING_LEVEL = 99;

    /* ナンバリング対象から除外する段落スタイル / Paragraph styles excluded from numbering */
    var IGNORE_STYLE_NAMES = ["p.img", "p.table"];

    /* 末尾ナンバリングを見つける正規表現 / Pattern that matches trailing numbering */
    var NUMBERING_PATTERN = /[（\(][0-9０-９]+[）\)]$/;

    /* 識別キーを連結する区切り / Separator used to join the key */
    var KEY_SEPARATOR = "___";

    // =========================================
    // レイアウト設定 / Layout settings
    // =========================================

    /* 対象リストのサイズ [幅, 高さ]（px）/ Size of the target list [width, height] (px) */
    var TARGET_LIST_SIZE = [400, 400];

    /* 進捗バーのサイズ [幅, 高さ]（px）/ Size of the progress bar [width, height] (px) */
    var PROGRESS_BAR_SIZE = [330, 7];

    /* リスト表示で省略を始める文字数と、省略後に残す文字数 / Length that triggers truncation, and the kept length */
    var LIST_TEXT_MAX_LENGTH  = 28;
    var LIST_TEXT_KEEP_LENGTH = 25;

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

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
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
            title: { ja: "繰り返し段落に連番を追加", en: "Number Repeated Paragraphs" }
        },
        panel: {
            paragraphStyle: { ja: "段落スタイル", en: "Paragraph Style" },
            scope: { ja: "範囲", en: "Scope" }
        },
        fieldLabel: {
            brackets: { ja: "括弧", en: "Brackets" }
        },
        radio: {
            selection: { ja: "選択範囲", en: "Selection" },
            story: { ja: "ストーリー", en: "Story" },
            document: { ja: "ドキュメント", en: "Document" },
            fullWidth: { ja: "全角", en: "Full-width" },
            halfWidth: { ja: "半角", en: "Half-width" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            removeNumbering: { ja: "番号を削除", en: "Remove Numbers" },
            addNumbering: { ja: "番号を追加", en: "Add Numbers" }
        },
        tooltip: {
            styleFilter: {
                ja: "オフにした段落スタイルの項目は、リストで選べなくなります",
                en: "Items with an unchecked paragraph style can no longer be selected in the list"
            },
            scopeSelection: {
                ja: "選択した文字を含む段落だけを処理します",
                en: "Processes only paragraphs that contain the selected text"
            },
            scopeStory: {
                ja: "選択中のテキストを含むストーリー全体を処理します",
                en: "Processes the whole story that contains the selection"
            },
            scopeDocument: {
                ja: "ドキュメント内のすべてのストーリーを処理します（親ページ上のテキストを除く）",
                en: "Processes every story in the document (text on parent pages is skipped)"
            },
            brackets: {
                ja: "連番を囲む括弧の種類",
                en: "Bracket style around the number"
            },
            removeNumbering: {
                ja: "リストで選択した項目の末尾にある番号を、範囲内から削除します",
                en: "Removes the trailing numbers of the items selected in the list, within the scope"
            },
            addNumbering: {
                ja: "既存の番号を外してから、出現順に連番を付け直します",
                en: "Removes existing numbers, then renumbers in order of appearance"
            }
        },
        progress: {
            title: { ja: "解析中…", en: "Analyzing…" }
        },
        alert: {
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            noTargets: {
                ja: "繰り返している段落が見つかりませんでした。",
                en: "No repeated paragraphs were found."
            },
            noSelection: {
                ja: "何も選択されていません。ドキュメント全体を対象にします。",
                en: "Nothing is selected. The entire document will be processed."
            },
            noRange: {
                ja: "文字が選択されていません。ストーリー全体を対象にします。",
                en: "No text range is selected. The entire story will be processed."
            },
            notStory: {
                ja: "選択したオブジェクトはストーリーとして認識できません。ドキュメント全体を対象にします。",
                en: "The selected object is not recognized as a story. The entire document will be processed."
            },
            removed: {
                ja: "選択した項目から番号を削除しました。",
                en: "Removed numbering from the selected items."
            }
        }
    };

    // =========================================
    // ヘルパー / Helpers
    // =========================================

    /**
     * everyItem() の戻り値を配列に正規化する（要素が1件のときスカラーで返るため）
     * @param {*} value everyItem() で取得した値
     * @returns {Array} 正規化した配列
     */
    function toArray(value) {
        return (value instanceof Array) ? value : [value];
    }

    /**
     * ストーリーが親（マスター）ページ上にあるかを判定する
     * @param {Story} story 対象のストーリー
     * @returns {boolean} 親ページ上なら true
     */
    function isMasterStory(story) {
        var textContainers = story.textContainers;
        if (!textContainers || textContainers.length === 0) return false;
        /* グループ内のフレームもあるのでスプレッドに達するまで親を遡る / Walk up until a spread is reached */
        var ancestor = textContainers[0];
        while (ancestor) {
            var typeName = ancestor.constructor.name;
            if (typeName === "MasterSpread") return true;
            if (typeName === "Spread" || typeName === "Document" || typeName === "Application") return false;
            ancestor = ancestor.parent;
        }
        return false;
    }

    /**
     * 走査対象から外す段落かを判定する（解析・付与・削除で共通）
     * @param {string} cleanedText ナンバリングを除いた本文
     * @param {string} styleName 段落スタイル名
     * @returns {boolean} 除外するなら true
     */
    function isSkippedParagraph(cleanedText, styleName) {
        /* 空行・1文字以下・空白のみはスキップ / Skip empty, single-char, or whitespace-only */
        if (cleanedText.length <= 1) return true;
        if (/^\s+$/.test(cleanedText)) return true;
        for (var i = 0; i < IGNORE_STYLE_NAMES.length; i++) {
            if (IGNORE_STYLE_NAMES[i] === styleName) return true;
        }
        return false;
    }

    /**
     * 末尾の改行を取り除く
     * @param {string} text 対象の文字列
     * @returns {string} 末尾の改行を除いた文字列
     */
    function trimTrailingBreaks(text) {
        return String(text).replace(/[\r\n]+$/, "");
    }

    /**
     * 段落本文からナンバリングを除いた比較用テキストを作る
     * @param {string} contents 段落の内容
     * @returns {string} 比較用テキスト
     */
    function toCleanedText(contents) {
        return trimTrailingBreaks(contents).replace(NUMBERING_PATTERN, "");
    }

    /**
     * 段落末尾の改行を除いた最後の文字位置を返す
     * @param {Paragraph} paragraph 対象の段落
     * @returns {number} 文字インデックス。該当がなければ -1
     */
    function getLastVisibleIndex(paragraph) {
        var index = paragraph.characters.length - 1;
        while (index >= 0) {
            var charContents = paragraph.characters[index].contents;
            if (charContents !== "\r" && charContents !== "\n") break;
            index--;
        }
        return index;
    }

    /**
     * 進捗バーを表示しながら処理を実行する
     * @param {number} maxValue 進捗の最大値
     * @param {function} progressTask 進捗更新関数を受け取る処理
     * @returns {void}
     */
    function withProgressBar(maxValue, progressTask) {
        var progressWindow = new Window("palette", getLabel("progress.title"));
        setupWindow(progressWindow);
        var progressBar = progressWindow.add("progressbar", undefined, 0, maxValue);
        progressBar.preferredSize = PROGRESS_BAR_SIZE;
        progressWindow.show();

        /* 解析が失敗してもパレットを残さない / Never leave the palette behind when the scan throws */
        try {
            progressTask(function (value) {
                progressBar.value = value;
                progressWindow.update();
            });
        } finally {
            progressWindow.close();
        }
    }

    // =========================================
    // 段落の走査 / Paragraph scanning
    // =========================================

    /**
     * 見出しスタックを更新しながら、段落の識別キーを作る
     * @param {string} cleanedText ナンバリングを除いた本文
     * @param {string} styleName 段落スタイル名
     * @param {Array<object>} headingStack 見出しの階層スタック
     * @returns {object} 識別キーと関連情報
     */
    function buildKeyForParagraph(cleanedText, styleName, headingStack) {
        var headingLevel = HEADING_LEVEL_MAP[styleName];
        if (typeof headingLevel !== "number") headingLevel = NON_HEADING_LEVEL;

        /* 現在のレベル以上の親をスタックから除去 / Pop parents at the same or deeper level */
        while (headingStack.length > 0 && headingStack[headingStack.length - 1].level >= headingLevel) {
            headingStack.pop();
        }
        /* 見出しなら親としてスタックに積む / Push headings onto the heading stack */
        if (headingLevel <= MAX_HEADING_LEVEL) {
            headingStack.push({ text: cleanedText, style: styleName, level: headingLevel });
        }

        /* 親は直近の見出しだけを見る。祖先まで含めると章ごとに分かれ、繰り返しと見なされなくなる
           / Use only the nearest heading; including ancestors splits repeats per chapter */
        var parentHeading = (headingStack.length > 0) ? headingStack[headingStack.length - 1] : null;
        var parentLabel = parentHeading ? (parentHeading.style + ":" + parentHeading.text) : "";

        return {
            key: styleName + KEY_SEPARATOR + cleanedText + KEY_SEPARATOR + parentLabel,
            style: styleName,
            text: cleanedText,
            parentLabel: parentLabel
        };
    }

    /**
     * ストーリー内の全段落の内容と段落スタイルを読み出す
     * @param {Story} story 対象のストーリー
     * @returns {object} 内容の一覧 contentsList と段落スタイルの一覧 styleList
     */
    function readParagraphContentsAndStyles(story) {
        var paragraphCount = story.paragraphs.length;
        /* 内容とスタイルを一括取得して段落ごとの DOM アクセスを減らす / Bulk-read to cut per-paragraph DOM access */
        var contentsList = toArray(story.paragraphs.everyItem().contents);
        var styleList = toArray(story.paragraphs.everyItem().appliedParagraphStyle);
        if (contentsList.length === paragraphCount && styleList.length === paragraphCount) {
            return { contentsList: contentsList, styleList: styleList };
        }

        /* 段落数と合わなければ取りこぼすので個別取得に切り替える / Fall back per paragraph when the bulk read does not line up */
        contentsList = [];
        styleList = [];
        for (var i = 0; i < paragraphCount; i++) {
            contentsList.push(story.paragraphs[i].contents);
            styleList.push(story.paragraphs[i].appliedParagraphStyle);
        }
        return { contentsList: contentsList, styleList: styleList };
    }

    /**
     * ストーリーを走査して、対象になる段落の情報を集める
     * @param {Story} story 対象のストーリー
     * @returns {Array<object>} 段落番号と識別キーを持つ情報の一覧
     */
    function scanStory(story) {
        var paragraphInfos = [];
        if (isMasterStory(story) || story.paragraphs.length === 0) return paragraphInfos;

        var paragraphData = readParagraphContentsAndStyles(story);
        var headingStack = [];
        for (var i = 0; i < paragraphData.contentsList.length; i++) {
            var cleanedText = toCleanedText(paragraphData.contentsList[i]);
            var styleName = paragraphData.styleList[i].name;
            if (isSkippedParagraph(cleanedText, styleName)) continue;

            var paragraphInfo = buildKeyForParagraph(cleanedText, styleName, headingStack);
            paragraphInfo.index = i;
            paragraphInfos.push(paragraphInfo);
        }
        return paragraphInfos;
    }

    /**
     * 対象範囲を走査し、識別キーが一致した段落を処理する
     * @param {Array<object>} targetScopes 対象範囲 { story, fromOffset, toOffset } の一覧
     * @param {object} keyMap 対象の識別キーを持つマップ
     * @param {function} handler 一致した段落に対する処理（引数: 段落, 識別キー）
     * @returns {void}
     */
    function eachMatchedParagraph(targetScopes, keyMap, handler) {
        for (var i = 0; i < targetScopes.length; i++) {
            var targetScope = targetScopes[i];
            /* 範囲を絞る場合も走査はストーリー全体で行う。手前の見出しを見ないと親が変わってしまう
               / Always scan the whole story: skipping earlier headings would change the parent */
            var paragraphInfos = scanStory(targetScope.story);
            for (var j = 0; j < paragraphInfos.length; j++) {
                if (!(paragraphInfos[j].key in keyMap)) continue;
                /* 付与も削除も段落数を変えないので、走査時の段落番号をそのまま使える
                   / Neither handler changes the paragraph count, so scanned indexes stay valid */
                var paragraph = targetScope.story.paragraphs[paragraphInfos[j].index];
                if (!isInScope(paragraph, targetScope)) continue;
                handler(paragraph, paragraphInfos[j].key);
            }
        }
    }

    // =========================================
    // 解析 / Analysis
    // =========================================

    /**
     * 全ストーリーを走査し、識別キーごとの出現回数を数える
     * @param {Stories} allStories 対象のストーリー
     * @returns {object} 識別キーをキーに、出現回数 count を持つ段落情報を値にしたマップ
     */
    function countOccurrencesByKey(allStories) {
        var occurrenceMap = {};
        withProgressBar(allStories.length, function (setProgress) {
            for (var i = 0; i < allStories.length; i++) {
                var paragraphInfos = scanStory(allStories[i]);
                for (var j = 0; j < paragraphInfos.length; j++) {
                    var countedInfo = occurrenceMap[paragraphInfos[j].key];
                    if (countedInfo) {
                        countedInfo.count++;
                    } else {
                        paragraphInfos[j].count = 1;
                        occurrenceMap[paragraphInfos[j].key] = paragraphInfos[j];
                    }
                }
                setProgress(i + 1);
            }
        });
        return occurrenceMap;
    }

    /**
     * 全ストーリーを解析し、ナンバリング対象の候補を求める
     * @param {Stories} allStories 対象のストーリー
     * @returns {Array<object>} 出現回数の多い順に並べた候補
     */
    function findNumberingTargets(allStories) {
        var occurrenceMap = countOccurrencesByKey(allStories);

        /* 2回以上出現するものを対象に。見出しスタイルは HEADING_LEVEL_MAP にある名前だけを親として
           扱うので、親の有無は条件にしない
           / Keep whatever repeats; only names in HEADING_LEVEL_MAP count as parents, so a parent is not required */
        var numberingTargets = [];
        for (var occurrenceKey in occurrenceMap) {
            if (occurrenceMap[occurrenceKey].count >= 2) {
                numberingTargets.push(occurrenceMap[occurrenceKey]);
            }
        }

        /* 出現回数の多い順、同数ならスタイル名順 / Sort by count desc, then style name */
        numberingTargets.sort(function (a, b) {
            if (b.count !== a.count) return b.count - a.count;
            return a.style.toLowerCase() < b.style.toLowerCase() ? -1 : 1;
        });

        return numberingTargets;
    }

    // =========================================
    // 対象範囲 / Target scope
    // =========================================

    /* 対象範囲の指定 / Scope of the target range */
    var SCOPE_SELECTION = "selection";
    var SCOPE_STORY = "story";
    var SCOPE_DOCUMENT = "document";

    /**
     * ストーリーを丸ごと対象にする範囲へ変換する
     * @param {Array<Story>|Stories} stories 対象のストーリー
     * @returns {Array<object>} 範囲を絞らない対象範囲の一覧
     */
    function toWholeStoryScopes(stories) {
        var targetScopes = [];
        for (var i = 0; i < stories.length; i++) {
            targetScopes.push({ story: stories[i], fromOffset: null, toOffset: null });
        }
        return targetScopes;
    }

    /**
     * 段落が対象範囲に含まれるかを判定する
     * @param {Paragraph} paragraph 対象の段落
     * @param {object} targetScope 対象範囲 { fromOffset, toOffset }
     * @returns {boolean} 含まれるなら true
     */
    function isInScope(paragraph, targetScope) {
        if (targetScope.fromOffset === null) return true;
        /* 段落の一部でも選択範囲にかかっていれば対象にする / A partial overlap is enough */
        var paraStart = paragraph.characters[0].index;
        var paraEnd = paragraph.characters[-1].index;
        return paraEnd >= targetScope.fromOffset && paraStart <= targetScope.toOffset;
    }

    /**
     * 選択オブジェクトが属するストーリーを求める
     * @param {object} selectedItem app.selection の要素
     * @returns {Story|null} 属するストーリー。求められなければ null
     */
    function getParentStory(selectedItem) {
        if (selectedItem.hasOwnProperty("parentStory")) return selectedItem.parentStory;
        /* 表のセルなどは親のほうがストーリーを持つ / Cells and similar carry the story on the parent */
        if (selectedItem.parent && selectedItem.parent.hasOwnProperty("parentStory")) return selectedItem.parent.parentStory;
        return null;
    }

    /**
     * 対象範囲に応じて処理するストーリーと文字位置を求める
     * @param {string} scopeMode SCOPE_SELECTION / SCOPE_STORY / SCOPE_DOCUMENT のいずれか
     * @param {Stories} allStories ドキュメント内の全ストーリー
     * @returns {Array<object>} 対象範囲 { story, fromOffset, toOffset } の一覧
     */
    function resolveTargetScopes(scopeMode, allStories) {
        if (scopeMode === SCOPE_DOCUMENT) return toWholeStoryScopes(allStories);

        if (app.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return toWholeStoryScopes(allStories);
        }
        var selectedItem = app.selection[0];
        var parentStory = getParentStory(selectedItem);
        if (!parentStory) {
            alert(getLabel("alert.notStory"));
            return toWholeStoryScopes(allStories);
        }
        if (scopeMode === SCOPE_STORY) return toWholeStoryScopes([parentStory]);

        /* 選択範囲：選択した文字の範囲だけに絞る / Selection: narrow down to the selected characters */
        var selectedChars = selectedItem.hasOwnProperty("characters") ? selectedItem.characters : null;
        if (!selectedChars || selectedChars.length === 0) {
            alert(getLabel("alert.noRange"));
            return toWholeStoryScopes([parentStory]);
        }
        return [{
            story: parentStory,
            fromOffset: selectedChars[0].index,
            toOffset: selectedChars[-1].index
        }];
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 対象リストに表示する文字列を作る（長い本文は省略し、出現回数を添える）
     * @param {object} numberingTarget ナンバリング対象の候補
     * @returns {string} 表示用の文字列
     */
    function formatTargetListText(numberingTarget) {
        var bodyText = numberingTarget.text;
        if (bodyText.length > LIST_TEXT_MAX_LENGTH) bodyText = bodyText.substring(0, LIST_TEXT_KEEP_LENGTH) + "…";
        var countText = (uiLang === "ja") ? "（" + numberingTarget.count + "）" : " (" + numberingTarget.count + ")";
        return numberingTarget.style + ": " + bodyText + countText;
    }

    /**
     * 対象リストを作成して候補を並べる
     * @param {Group} parentGroup 追加先のグループ
     * @param {Array<object>} numberingTargets ナンバリング対象の候補
     * @returns {ListBox} 作成したリストボックス
     */
    function buildTargetListBox(parentGroup, numberingTargets) {
        var targetListBox = parentGroup.add("listbox", undefined, "", { multiselect: true });
        targetListBox.preferredSize = TARGET_LIST_SIZE;

        for (var i = 0; i < numberingTargets.length; i++) {
            var numberingTarget = numberingTargets[i];
            var listItem = targetListBox.add("item", formatTargetListText(numberingTarget));
            /* 省略された全文と、親見出しがあればその見出しを添える / Show the full text, plus the parent heading when there is one */
            listItem.helpTip = numberingTarget.parentLabel ? (numberingTarget.text + "\n" + numberingTarget.parentLabel) : numberingTarget.text;
        }
        if (targetListBox.items.length > 0) {
            targetListBox.items[0].selected = true;
        }
        return targetListBox;
    }

    /**
     * 候補に現れる段落スタイル名を重複なしで集め、名前順に並べる
     * @param {Array<object>} numberingTargets ナンバリング対象の候補
     * @returns {Array<string>} 段落スタイル名の一覧
     */
    function collectStyleNames(numberingTargets) {
        var styleNames = [];
        var seenStyles = {};
        for (var i = 0; i < numberingTargets.length; i++) {
            var styleName = numberingTargets[i].style;
            if (seenStyles[styleName] === true) continue;
            seenStyles[styleName] = true;
            styleNames.push(styleName);
        }
        return styleNames.sort();
    }

    /**
     * 段落スタイルの絞り込みチェックボックスを作り、対象リストと連動させる
     * @param {Panel} stylePanel 追加先のパネル
     * @param {Array<object>} numberingTargets ナンバリング対象の候補
     * @param {ListBox} targetListBox 連動させる対象リスト
     * @returns {void}
     */
    function buildStyleFilter(stylePanel, numberingTargets, targetListBox) {
        var styleNames = collectStyleNames(numberingTargets);
        var styleCheckboxes = {};

        /**
         * 対象リストの有効／無効を現在のチェック状態に合わせて切り替える
         * @returns {void}
         */
        function updateListBoxEnabled() {
            for (var i = 0; i < targetListBox.items.length; i++) {
                var listItem = targetListBox.items[i];
                listItem.enabled = styleCheckboxes[numberingTargets[i].style].value;
                if (!listItem.enabled) listItem.selected = false;
            }
        }

        for (var i = 0; i < styleNames.length; i++) {
            var styleCheckbox = stylePanel.add("checkbox", undefined, styleNames[i]);
            styleCheckbox.value = true;
            styleCheckbox.helpTip = getLabel("tooltip.styleFilter");
            styleCheckbox.onClick = updateListBoxEnabled;
            styleCheckboxes[styleNames[i]] = styleCheckbox;
        }
    }

    /**
     * 左カラム（段落スタイル・対象・括弧の種類）を組み立てる
     * @param {Group} columnsGroup 追加先のグループ
     * @returns {object} 段落スタイルパネルと、選択状態を読むラジオボタン
     */
    function buildOptionColumn(columnsGroup) {
        var optionColumnGroup = columnsGroup.add("group");
        optionColumnGroup.orientation = "column";
        optionColumnGroup.alignChildren = ["fill", "top"];
        optionColumnGroup.spacing = PANEL_SPACING;

        var stylePanel = optionColumnGroup.add("panel", undefined, getLabel("panel.paragraphStyle"));
        setupPanel(stylePanel, 6);
        stylePanel.alignChildren = ["left", "top"];

        var scopePanel = optionColumnGroup.add("panel", undefined, getLabel("panel.scope"));
        setupPanel(scopePanel, 6);
        scopePanel.alignChildren = ["left", "top"];
        var selectionRadio = scopePanel.add("radiobutton", undefined, getLabel("radio.selection"));
        var storyRadio = scopePanel.add("radiobutton", undefined, getLabel("radio.story"));
        var documentRadio = scopePanel.add("radiobutton", undefined, getLabel("radio.document"));
        selectionRadio.helpTip = getLabel("tooltip.scopeSelection");
        storyRadio.helpTip = getLabel("tooltip.scopeStory");
        documentRadio.helpTip = getLabel("tooltip.scopeDocument");
        storyRadio.value = true;

        /* 全角／半角選択（日本語UIのみ）/ Full/half-width selection (Japanese UI only) */
        var halfWidthRadio = null;
        if (uiLang === "ja") {
            var bracketRadioGroup = optionColumnGroup.add("group");
            setupRow(bracketRadioGroup, "center", 8);
            bracketRadioGroup.add("statictext", undefined, labelText("fieldLabel.brackets"));
            var fullWidthRadio = bracketRadioGroup.add("radiobutton", undefined, getLabel("radio.fullWidth"));
            halfWidthRadio = bracketRadioGroup.add("radiobutton", undefined, getLabel("radio.halfWidth"));
            fullWidthRadio.helpTip = halfWidthRadio.helpTip = getLabel("tooltip.brackets");
            fullWidthRadio.value = true;
        }

        return {
            stylePanel: stylePanel,
            selectionRadio: selectionRadio,
            storyRadio: storyRadio,
            halfWidthRadio: halfWidthRadio
        };
    }

    /**
     * ボタン行（左に削除、右にキャンセル・追加）を組み立てる
     * @param {Window} dialogWindow 追加先のダイアログ
     * @returns {Button} 削除ボタン
     */
    function buildButtonRow(dialogWindow) {
        var buttonRow = addButtonRow(dialogWindow);
        var btnRemove = buttonRow.leftGroup.add("button", undefined, getLabel("button.removeNumbering"));
        btnRemove.helpTip = getLabel("tooltip.removeNumbering");
        /* ラベルが "OK" / "Cancel" でないと既定の割り当てが効かないので name を明示
           / Labels other than "OK" / "Cancel" need an explicit name */
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.addNumbering"), { name: "ok" });
        btnOK.helpTip = getLabel("tooltip.addNumbering");
        alignRightOnlyButtonRow(buttonRow);

        return btnRemove;
    }

    /**
     * 対象選択ダイアログを組み立てる
     * @param {Array<object>} numberingTargets ナンバリング対象の候補
     * @param {Stories} allStories ドキュメント内の全ストーリー
     * @returns {object} ウィンドウと選択内容を取り出す関数を持つオブジェクト
     */
    function buildDialog(numberingTargets, allStories) {
        var dialogWindow = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialogWindow, 10);

        var columnsGroup = dialogWindow.add("group");
        setupRow(columnsGroup, "fill", COLUMN_SPACING);
        columnsGroup.alignChildren = ["fill", "top"];

        /* 左カラム：スタイル・対象・全角半角 / Left column: styles, target, brackets */
        var optionControls = buildOptionColumn(columnsGroup);

        /* 右カラム：対象リスト / Right column: target list */
        var listColumnGroup = columnsGroup.add("group");
        listColumnGroup.orientation = "column";
        listColumnGroup.alignChildren = ["fill", "top"];
        var targetListBox = buildTargetListBox(listColumnGroup, numberingTargets);
        buildStyleFilter(optionControls.stylePanel, numberingTargets, targetListBox);

        var btnRemove = buildButtonRow(dialogWindow);

        /**
         * リストで選択中の識別キーを集める
         * @returns {object} 識別キーをキーに持つマップ
         */
        function getSelectedKeys() {
            var selectedKeyMap = {};
            for (var i = 0; i < targetListBox.items.length; i++) {
                if (targetListBox.items[i].selected) {
                    selectedKeyMap[numberingTargets[i].key] = 1;
                }
            }
            return selectedKeyMap;
        }

        /**
         * 現在の対象範囲に応じた処理対象を求める
         * @returns {Array<object>} 対象範囲 { story, fromOffset, toOffset } の一覧
         */
        function getTargetScopes() {
            var scopeMode = SCOPE_DOCUMENT;
            if (optionControls.selectionRadio.value) scopeMode = SCOPE_SELECTION;
            else if (optionControls.storyRadio.value) scopeMode = SCOPE_STORY;
            return resolveTargetScopes(scopeMode, allStories);
        }

        /**
         * ナンバリングに使う括弧を求める
         * @returns {object} 左右の括弧 { left, right }
         */
        function getBrackets() {
            var halfWidthRadio = optionControls.halfWidthRadio;
            if (halfWidthRadio && halfWidthRadio.value) return { left: "(", right: ")" };
            return { left: "（", right: "）" };
        }

        btnRemove.onClick = function () {
            var targetScopes = getTargetScopes();
            var selectedKeyMap = getSelectedKeys();
            /* リストが古くなるので、書き換える前にダイアログを閉じる / Close first: the list goes stale once text changes */
            dialogWindow.close(2);
            removeNumbering(targetScopes, selectedKeyMap);
            alert(getLabel("alert.removed"));
        };

        return {
            window: dialogWindow,
            getSelectedKeys: getSelectedKeys,
            getTargetScopes: getTargetScopes,
            getBrackets: getBrackets
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 段落末尾の既存ナンバリングを削除する
     * @param {Paragraph} paragraph 対象の段落
     * @returns {void}
     */
    function removeExistingNumbering(paragraph) {
        var numberingMatch = trimTrailingBreaks(paragraph.contents).match(NUMBERING_PATTERN);
        if (!numberingMatch) return;
        var endIndex = getLastVisibleIndex(paragraph);
        var startIndex = endIndex - numberingMatch[0].length + 1;
        if (startIndex < 0) return;
        paragraph.characters.itemByRange(startIndex, endIndex).remove();
    }

    /**
     * 選択した対象の末尾からナンバリングを削除する（取り消しは1回で戻る）
     * @param {Array<object>} targetScopes 対象範囲 { story, fromOffset, toOffset } の一覧
     * @param {object} keyMap 対象の識別キーを持つマップ
     * @returns {void}
     */
    function removeNumbering(targetScopes, keyMap) {
        app.doScript(function () {
            eachMatchedParagraph(targetScopes, keyMap, removeExistingNumbering);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Remove Paragraph Numbering");
    }

    /**
     * 選択した対象の末尾にナンバリングを付与する（取り消しは1回で戻る）
     * @param {Array<object>} targetScopes 対象範囲 { story, fromOffset, toOffset } の一覧
     * @param {object} keyMap 対象の識別キーを持つマップ
     * @param {object} brackets 使用する括弧 { left, right }
     * @returns {void}
     */
    function applyNumbering(targetScopes, keyMap, brackets) {
        var counterByKey = {};
        app.doScript(function () {
            eachMatchedParagraph(targetScopes, keyMap, function (paragraph, key) {
                removeExistingNumbering(paragraph);
                var counter = (counterByKey[key] || 0) + 1;
                counterByKey[key] = counter;
                paragraph.insertionPoints[getLastVisibleIndex(paragraph) + 1].contents = brackets.left + counter + brackets.right;
            });
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Append Paragraph Numbering");
    }

    /**
     * 対象を解析し、ダイアログの指定に従ってナンバリングを付与する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var allStories = app.activeDocument.stories;
        var numberingTargets = findNumberingTargets(allStories);
        if (numberingTargets.length === 0) {
            alert(getLabel("alert.noTargets"));
            return;
        }

        var numberingDialog = buildDialog(numberingTargets, allStories);
        if (numberingDialog.window.show() != 1) return;

        applyNumbering(numberingDialog.getTargetScopes(), numberingDialog.getSelectedKeys(), numberingDialog.getBrackets());
    }

    main();

})();
