#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdConvertDocumentToGrayscale
Meta-Default-language=ja
Name=ドキュメント全体を CMYK のモノクロにする（テキストは K100）
Name[en]=Turn the whole document into CMYK monochrome (text in K100)
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.1.0
Release-Date=2026-11-21
Tags=document,ドキュメント
X-com.swwwitch-Script-Name=IdConvertDocumentToGrayscale
X-com.swwwitch-Updated=2026-11-21
X-com.swwwitch-README-JA=https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdConvertDocumentToGrayscale.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdConvertDocumentToGrayscale.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdConvertDocumentToGrayscale.jsx
Description-BEGIN
[ja]
ドキュメントのカラーをすべて CMYK の K だけのグレーに変換し、テキストを K100 にします。リンクした AI・EPS は Illustrator で、PSD は Photoshop でモノクロにし、別名で保存してリンクを差し替えます。
[en]
Converts every color in the document to a CMYK gray that uses only K and sets text to K100. Linked AI and EPS files are turned monochrome in Illustrator and PSD files in Photoshop, saved under another name and relinked.
Description-END
SCRIPTMETA-END
*/

/*
### 注意
- 透明ブレンド領域を CMYK にします。
- 同じ濃さになったスウォッチは1つにまとめて「C=0 M=0 Y=0 K=n」の名前にし（K100 は［黒］）、未使用のスウォッチを削除します。グラデーションの分岐点と濃淡の元のカラーは残します。
- グレーの濃さは、カラーを Lab に変換したときの明度（L*）が同じになる K の値です（ドキュメントのカラー設定で換算）。
- 特色はプロセスカラーにします。混合インキは変換しません。
- 配置した画像・グラフィックの中身は変換しません。
- テキストの塗りと線の単色は K100（［黒］）にします。白（紙色・K0）は白のまま残します（KEEP_WHITE_TEXT）。
- ロックしたレイヤーとオブジェクトも一時的に解除して処理し、終わったら元に戻します。
- リンクした AI・EPS・PSD があるときは、一緒に変換するか確認します。AI・EPS は Illustrator 版 ConvertDocumentToGrayscale と同じ処理、PSD は［グレースケール］にしてテキストレイヤーを黒に。同じフォルダーに「名前_gray」で保存してリンクを差し替えます（元のファイルは変更しない）。
- 名前が「_gray」で終わるリンク、埋め込んだリンク、見つからないリンク、Illustrator・Photoshop で開いたままのファイルは対象外です。
- 1回の取り消しで戻せます（保存したファイルは残ります）。

### Notes
- The transparency blend space is set to CMYK.
- Swatches with the same gray are merged into one named "C=0 M=0 Y=0 K=n" (K100 goes to [Black]) and unused swatches are deleted. Colors used in gradient stops or as the base of tints are kept.
- The gray is the K value whose lightness (L*) in Lab matches the color (converted with the document's color settings).
- Spot colors become process colors. Mixed inks are not converted.
- The contents of placed images and graphics are not converted.
- Solid fill and stroke colors of text become K100 ([Black]). White ([Paper], K0) stays white (KEEP_WHITE_TEXT).
- Locked layers and objects are unlocked while processing, then restored.
- When AI, EPS or PSD files are linked, the script asks whether to convert them too. AI and EPS files get the work of the Illustrator script ConvertDocumentToGrayscale; PSD files are set to Grayscale with black text layers. They are saved in the same folder as "name_gray" and relinked (the originals are not changed).
- Links whose names end in "_gray", embedded and missing links, and files open in Illustrator or Photoshop are skipped.
- A single Undo reverts everything (the saved files remain).
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

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 白い文字（紙色・K0）は K100 にせず白のまま残す / Keep white text ([Paper], K0) white instead of K100 */
    var KEEP_WHITE_TEXT = true;

    /* 保存するファイル名の末尾 / Suffix of the saved file names */
    var OUTPUT_SUFFIX = "_gray";

    /* 送り先のアプリ（"illustrator" "photoshop" は起動中の最新版。版を固定するなら "illustrator-31.0" など）
       Target apps ("illustrator" / "photoshop" is the newest running version; pin one with e.g. "illustrator-31.0") */
    var ILLUSTRATOR_TARGET = "illustrator";
    var PHOTOSHOP_TARGET = "photoshop";

    /* 1ファイルの処理を待つ秒数 / Seconds to wait for each file */
    var FILE_TIMEOUT_SECONDS = 300;

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

    var PROGRESS_BAR_WIDTH = 320;            /* 進行状況バーの幅 / progress bar width */

    // =========================================
    // ローカライズ / Localization
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
        undoName: { ja: "ドキュメントを CMYK のモノクロにする", en: "Convert Document to CMYK Monochrome" },
        dialog: {
            progressTitle: { ja: "ドキュメントを CMYK のモノクロにする", en: "Convert Document to CMYK Monochrome" }
        },
        progress: {
            launching: { ja: "{app} を起動しています...", en: "Launching {app}..." },
            converting: { ja: "リンクを変換中（{current} / {total}）: {name}", en: "Converting links ({current} / {total}): {name}" },
            documentColors: { ja: "リンクを差し替えて、ドキュメントのカラーを変換しています...", en: "Relinking and converting the document colors..." }
        },
        alert: {
            confirmLinks: {
                ja: "リンクした {count} 個のファイル（AI・EPS は Illustrator、PSD は Photoshop）もモノクロにし、「{suffix}」を付けた別名で保存してリンクを差し替えますか？\n元のファイルは変更しません。［いいえ］のときは、ドキュメントのカラーだけを変換します。",
                en: "Also turn the {count} linked files (AI and EPS in Illustrator, PSD in Photoshop) into monochrome, save them with \"{suffix}\" added to the name, and relink?\nThe original files are not changed. With No, only the document colors are converted."
            },
            done: { ja: "ドキュメントのカラーを変換し、リンクした {count} 個のファイルを変換して {links} 個のリンクを差し替えました。", en: "Converted the document colors, converted {count} linked files and relinked {links} links." },
            failedHeader: { ja: "変換できなかったファイル :", en: "Files that could not be converted:" },
            notLaunched: { ja: "{app} を起動できませんでした", en: "could not launch {app}" },
            alreadyOpen: { ja: "{app} で開いています", en: "open in {app}" },
            timeout: { ja: "時間切れ", en: "timed out" }
        }
    };

    // =========================================
    // ロックの一時解除 / Temporary unlock
    // =========================================

    /**
     * レイヤーとオブジェクト（グループの中も含む）のロックを解除し、元の状態を控える
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object[]} 解除したもの（restoreLocks に渡す）
     */
    function unlockLayersAndItems(doc) {
        var unlockedTargets = [];
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].locked) {
                doc.layers[i].locked = false;
                unlockedTargets.push(doc.layers[i]);
            }
        }
        var pageItems = doc.allPageItems;
        for (var j = 0; j < pageItems.length; j++) {
            try {
                if (pageItems[j].locked) {
                    pageItems[j].locked = false;
                    unlockedTargets.push(pageItems[j]);
                }
            } catch (e) {
                /* 変更できないオブジェクトはそのまま / Leave items that cannot be changed */
            }
        }
        return unlockedTargets;
    }

    /**
     * 解除したロックを逆順に戻す（オブジェクトを先に、レイヤーを後に）
     * @param {Object[]} unlockedTargets - unlockLayersAndItems の戻り値
     * @returns {void}
     */
    function restoreLocks(unlockedTargets) {
        for (var i = unlockedTargets.length - 1; i >= 0; i--) {
            try {
                unlockedTargets[i].locked = true;
            } catch (e) {
                /* 戻せないものは飛ばす / Skip what cannot be restored */
            }
        }
    }

    // =========================================
    // グレーへの換算 / Conversion to gray
    // =========================================

    /**
     * K0〜K100 の明度（Lab の L*）の表を作る（ドキュメントのカラー設定で換算）
     * @param {Document} doc - 対象ドキュメント
     * @returns {number[]} K の値（添字）→ L*
     */
    function buildGrayLightnessTable(doc) {
        var lightnessByK = [];
        var probeColor = doc.colors.add({ model: ColorModel.PROCESS, space: ColorSpace.CMYK, colorValue: [0, 0, 0, 0] });
        try {
            for (var k = 0; k <= 100; k++) {
                probeColor.space = ColorSpace.CMYK;
                probeColor.colorValue = [0, 0, 0, k];
                probeColor.space = ColorSpace.LAB;
                lightnessByK.push(probeColor.colorValue[0]);
            }
        } finally {
            probeColor.remove();
        }
        return lightnessByK;
    }

    /**
     * 明度（L*）が同じになる K の値を返す（表の間は直線で補う）
     * @param {number} lightness - L*
     * @param {number[]} lightnessByK - buildGrayLightnessTable の表
     * @returns {number} K の値（0〜100 の整数）
     */
    function lightnessToK(lightness, lightnessByK) {
        if (lightness >= lightnessByK[0]) return 0;
        for (var k = 1; k < lightnessByK.length; k++) {
            if (lightness >= lightnessByK[k]) {
                var lighter = lightnessByK[k - 1];
                var darker = lightnessByK[k];
                var ratio = (lighter === darker) ? 0 : (lighter - lightness) / (lighter - darker);
                return Math.round(k - 1 + ratio);
            }
        }
        return 100;
    }

    /**
     * カラーを K だけの CMYK プロセスカラーに変換する（特色はプロセスに）
     * @param {Color} color - 対象のカラー
     * @param {number[]} lightnessByK - buildGrayLightnessTable の表
     * @returns {number} 変換後の K の値
     */
    function convertColorToGray(color, lightnessByK) {
        var grayK;
        var colorValue = color.colorValue;
        if (color.space === ColorSpace.CMYK && colorValue[0] === 0 && colorValue[1] === 0 && colorValue[2] === 0) {
            /* もとから K だけのカラーは値を変えない / Keep colors that already use only K */
            grayK = Math.round(colorValue[3]);
        } else {
            color.space = ColorSpace.LAB;
            grayK = lightnessToK(color.colorValue[0], lightnessByK);
        }
        if (color.model !== ColorModel.PROCESS) color.model = ColorModel.PROCESS;
        color.space = ColorSpace.CMYK;
        color.colorValue = [0, 0, 0, grayK];
        return grayK;
    }

    /**
     * 変換するカラーを集める（スウォッチのカラーと名前の無いカラー。［黒］［紙色］［レジストレーション］と混合インキは除く）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Color[]} 変換するカラー
     */
    function collectConvertibleColors(doc) {
        var fixedNames = { "None": true, "Registration": true, "Paper": true, "Black": true };
        /* doc.colors には隠れたインキのカラー（Cyan など）もあるので、名前のあるものはスウォッチにあるものだけ
           doc.colors also holds hidden ink colors (Cyan etc.); take named ones only when they are swatches */
        var swatchNames = {};
        for (var i = 0; i < doc.swatches.length; i++) swatchNames[doc.swatches[i].name] = true;

        var colors = [];
        for (var j = 0; j < doc.colors.length; j++) {
            var color = doc.colors[j];
            var colorName = color.name;
            if (colorName !== "" && (fixedNames[colorName] || !swatchNames[colorName])) continue;
            if (color.model === ColorModel.MIXEDINKMODEL) continue;
            colors.push(color);
        }
        return colors;
    }

    /**
     * 名前のあるカラーのうち同じ濃さのものを1つにまとめ、「C=0 M=0 Y=0 K=n」の名前にする（K100 は［黒］にまとめる）
     * @param {Document} doc - 対象ドキュメント
     * @param {Color[]} colors - 変換したカラー
     * @param {number[]} grayValues - colors と同じ順の K の値
     * @returns {void}
     */
    function mergeSameGrayColors(doc, colors, grayValues) {
        var keeperByK = { "100": doc.colors.itemByName("Black") };
        var keepers = [];
        for (var i = 0; i < colors.length; i++) {
            if (colors[i].name === "" || grayValues[i] == null) continue;
            var grayKey = String(grayValues[i]);
            if (!keeperByK.hasOwnProperty(grayKey)) {
                keeperByK[grayKey] = colors[i];
                keepers.push({ color: colors[i], grayK: grayValues[i] });
                continue;
            }
            try {
                /* 使っている所（濃淡の元・グラデーションを含む）を残すカラーに置き換えて削除 / Replace usages and delete */
                colors[i].remove(keeperByK[grayKey]);
            } catch (e) {
                /* 削除できないカラーは残す / Keep colors that cannot be removed */
            }
        }
        for (var j = 0; j < keepers.length; j++) {
            try {
                keepers[j].color.name = "C=0 M=0 Y=0 K=" + keepers[j].grayK;
            } catch (e) {
                /* 同じ名前が残っているときは元の名前のまま / Keep the old name when the new one is taken */
            }
        }
    }

    // =========================================
    // テキストの K100 / K100 text
    // =========================================

    /**
     * スウォッチのプロパティを読む（種類によって無いプロパティは例外になる）
     * @param {Swatch} swatch - スウォッチ
     * @param {string} propertyKey - プロパティ名
     * @returns {*} 値。読めなければ undefined
     */
    function readSwatchProperty(swatch, propertyKey) {
        /* try の中で return swatch[key] と書くと例外が catch に届かない（ExtendScript）ので、変数に入れてから返す
           In ExtendScript an exception in "return swatch[key]" inside try escapes the catch; read into a variable first */
        var propertyValue;
        try {
            propertyValue = swatch[propertyKey];
        } catch (e) {
            propertyValue = undefined;
        }
        return propertyValue;
    }

    /**
     * スウォッチの K の濃さを返す（濃淡は元のカラー × 濃淡）。
     * constructor.name は種類によらず "Swatch" なので、読めるプロパティで見分ける
     * @param {Swatch} swatch - スウォッチ
     * @param {number} [tintPercent] - 文字やスタイルの濃淡（-1・省略は 100）
     * @returns {number|null} K の値。CMYK のカラー・濃淡でなければ null
     */
    function getSwatchK(swatch, tintPercent) {
        var tintRatio = (tintPercent == null || tintPercent < 0) ? 1 : tintPercent / 100;
        var swatchName = readSwatchProperty(swatch, "name");
        if (swatchName === "Paper") return 0;
        if (swatchName === "None" || swatchName === "Registration") return null;

        /* 濃淡は tintValue を持ち、カラーは読むと例外 / Tints have tintValue; reading it on a color throws */
        var tintValue = readSwatchProperty(swatch, "tintValue");
        if (tintValue !== undefined) {
            var baseK = getSwatchK(swatch.baseColor);
            return (baseK == null) ? null : baseK * tintValue / 100 * tintRatio;
        }
        /* グラデーションと［なし］は space・colorValue を読むと例外 / Gradients and [None] throw on space and colorValue */
        if (readSwatchProperty(swatch, "space") !== ColorSpace.CMYK) return null;
        var colorValue = readSwatchProperty(swatch, "colorValue");
        return colorValue ? colorValue[3] * tintRatio : null;
    }

    /**
     * K100 にするスウォッチかどうかを返す（単色だけ。白は KEEP_WHITE_TEXT のとき残す）
     * @param {Swatch} swatch - 文字の塗りか線のスウォッチ
     * @param {number} [tintPercent] - 文字やスタイルの濃淡
     * @returns {boolean} K100 にするなら true
     */
    function shouldBecomeK100(swatch, tintPercent) {
        var swatchK = getSwatchK(swatch, tintPercent);
        if (swatchK == null) return false;
        return !(KEEP_WHITE_TEXT && swatchK === 0);
    }

    /**
     * 段落スタイル・文字スタイルの塗りと線を K100 にする
     * @param {Document} doc - 対象ドキュメント
     * @param {Color} blackColor - ［黒］
     * @returns {void}
     */
    function applyK100ToTextStyles(doc, blackColor) {
        var textStyles = doc.allParagraphStyles.concat(doc.allCharacterStyles);
        var colorKeys = [["fillColor", "fillTint"], ["strokeColor", "strokeTint"]];
        for (var i = 0; i < textStyles.length; i++) {
            for (var j = 0; j < colorKeys.length; j++) {
                try {
                    var styleSwatch = textStyles[i][colorKeys[j][0]];
                    if (!styleSwatch || !shouldBecomeK100(styleSwatch, textStyles[i][colorKeys[j][1]])) continue;
                    textStyles[i][colorKeys[j][0]] = blackColor;
                    textStyles[i][colorKeys[j][1]] = 100;
                } catch (e) {
                    /* ［段落スタイルなし］などは変更できない / Built-in styles cannot be changed */
                }
            }
        }
    }

    /**
     * テキストに直接付けた塗りと線を、スウォッチごとに検索と置換で K100 にする（表・脚注・親ページも対象）
     * @param {Document} doc - 対象ドキュメント
     * @param {Color} blackColor - ［黒］
     * @returns {void}
     */
    function applyK100ToText(doc, blackColor) {
        var textSwatches = [];
        for (var i = 0; i < doc.swatches.length; i++) textSwatches.push(doc.swatches[i]);
        for (var j = 0; j < doc.colors.length; j++) {
            if (doc.colors[j].name === "") textSwatches.push(doc.colors[j]);
        }

        var savedOptions = app.findChangeTextOptions.properties;
        app.findChangeTextOptions.properties = {
            includeHiddenLayers: true,
            includeLockedLayersForFind: true,
            includeLockedStoriesForFind: true,
            includeMasterPages: true,
            includeFootnotes: true,
            caseSensitive: false,
            wholeWord: false
        };
        try {
            var colorKeys = [["fillColor", "fillTint"], ["strokeColor", "strokeTint"]];
            for (var k = 0; k < textSwatches.length; k++) {
                /* 濃淡は文字ごとに違うので、白の判定はスウォッチだけで行う / Tints vary per character; judge white by the swatch alone */
                if (!shouldBecomeK100(textSwatches[k])) continue;
                for (var m = 0; m < colorKeys.length; m++) {
                    app.findTextPreferences = NothingEnum.NOTHING;
                    app.changeTextPreferences = NothingEnum.NOTHING;
                    app.findTextPreferences[colorKeys[m][0]] = textSwatches[k];
                    app.changeTextPreferences[colorKeys[m][0]] = blackColor;
                    app.changeTextPreferences[colorKeys[m][1]] = 100;
                    try {
                        doc.changeText();
                    } catch (e) {
                        /* 置換できないスウォッチは飛ばす / Skip swatches that cannot be replaced */
                    }
                }
            }
        } finally {
            app.findTextPreferences = NothingEnum.NOTHING;
            app.changeTextPreferences = NothingEnum.NOTHING;
            app.findChangeTextOptions.properties = savedOptions;
        }
    }

    // =========================================
    // スウォッチ / Swatches
    // =========================================

    /**
     * 未使用のスウォッチを削除する（グラデーションの分岐点と濃淡の元のカラーは使用中とみなして残す）
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function removeUnusedSwatches(doc) {
        var referencedIds = {};
        for (var i = 0; i < doc.gradients.length; i++) {
            var gradientStops = doc.gradients[i].gradientStops;
            for (var j = 0; j < gradientStops.length; j++) {
                try {
                    referencedIds[gradientStops[j].stopColor.id] = true;
                } catch (e) {
                    /* 読めない分岐点は飛ばす / Skip stops that cannot be read */
                }
            }
        }
        for (var k = 0; k < doc.tints.length; k++) {
            referencedIds[doc.tints[k].baseColor.id] = true;
        }

        var unusedSwatches = doc.unusedSwatches;
        for (var m = 0; m < unusedSwatches.length; m++) {
            if (unusedSwatches[m].name === "" || referencedIds[unusedSwatches[m].id]) continue;
            try {
                unusedSwatches[m].remove();
            } catch (e) {
                /* ［紙色］など削除できないスウォッチは残す / Keep swatches that cannot be removed, such as [Paper] */
            }
        }
    }

    // =========================================
    // リンクの収集 / Collect links
    // =========================================

    /**
     * ファイルの拡張子から、処理するアプリの情報を返す
     * @param {string} filePath - ファイルのパス
     * @returns {Object|null} {target, appName, workerFuncs, entryName}。対象外の拡張子は null
     */
    function getConverterForFile(filePath) {
        var extensionMatch = filePath.match(/\.([^.\/:]+)$/);
        var extension = extensionMatch ? extensionMatch[1].toLowerCase() : "";
        if (extension === "ai" || extension === "eps") return ILLUSTRATOR_CONVERTER;
        if (extension === "psd") return PHOTOSHOP_CONVERTER;
        return null;
    }

    /**
     * 対象の AI・EPS・PSD のリンクを、ファイルごとにまとめて返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object[]} {filePath, converter, links} の配列（同じファイルのリンクは1つにまとめる）
     */
    function collectTargetLinkGroups(doc) {
        var groupsByPath = {};
        var linkGroups = [];
        var convertedSuffix = OUTPUT_SUFFIX.toLowerCase();
        for (var i = 0; i < doc.links.length; i++) {
            var link = doc.links[i];
            var filePath = link.filePath;
            var converter = getConverterForFile(filePath);
            if (!converter) continue;
            /* 名前が OUTPUT_SUFFIX で終わるものは変換済み / Names ending in OUTPUT_SUFFIX are already converted */
            var baseName = filePath.replace(/\.[^.\/:]+$/, "").toLowerCase();
            if (baseName.slice(-convertedSuffix.length) === convertedSuffix) continue;
            if (link.status === LinkStatus.LINK_EMBEDDED || link.status === LinkStatus.LINK_MISSING || link.status === LinkStatus.LINK_INACCESSIBLE) continue;
            if (!groupsByPath.hasOwnProperty(filePath)) {
                groupsByPath[filePath] = { filePath: filePath, converter: converter, links: [] };
                linkGroups.push(groupsByPath[filePath]);
            }
            groupsByPath[filePath].links.push(link);
        }
        return linkGroups;
    }

    /**
     * 保存先のファイルを返す（同じフォルダー、名前の末尾に OUTPUT_SUFFIX）
     * @param {string} filePath - 元のファイルのパス
     * @returns {File} 保存先のファイル
     */
    function getOutputFile(filePath) {
        var sourceFile = new File(filePath);
        var fileName = decodeURI(sourceFile.name);
        var dotIndex = fileName.lastIndexOf(".");
        return new File(sourceFile.parent.fsName + "/" + fileName.substring(0, dotIndex) + OUTPUT_SUFFIX + fileName.substring(dotIndex));
    }

    // =========================================
    // BridgeTalk / BridgeTalk
    // =========================================

    /**
     * アプリが起動していなければ起動し、応答できるまで待つ
     * @param {string} target - BridgeTalk の送り先
     * @returns {boolean} 起動していれば true
     */
    function ensureAppRunning(target) {
        if (BridgeTalk.isRunning(target)) return true;
        BridgeTalk.launch(target);
        for (var waited = 0; waited < 120; waited++) {
            if (BridgeTalk.isRunning(target)) return true;
            $.sleep(1000);
        }
        return false;
    }

    /**
     * BridgeTalk でアプリに送り、結果が届くまで待つ
     * @param {string} target - BridgeTalk の送り先
     * @param {string} bodyText - 送り先で実行するスクリプト
     * @param {number} timeoutSeconds - 待つ秒数
     * @returns {string} 送り先からの戻り値。時間切れは "ERR:TIMEOUT"
     */
    function sendAndWait(target, bodyText, timeoutSeconds) {
        var resultText = null;
        var message = new BridgeTalk();
        message.target = target;
        message.body = bodyText;
        message.onResult = function (response) {
            resultText = String(response.body);
        };
        message.onError = function (response) {
            resultText = "ERR:" + encodeURI(String(response.body));
        };
        message.send();
        var startTime = new Date().getTime();
        while (resultText === null) {
            BridgeTalk.pump();
            if (new Date().getTime() - startTime > timeoutSeconds * 1000) return "ERR:TIMEOUT";
            $.sleep(100);
        }
        return resultText;
    }

    // =========================================
    // Illustrator・Photoshop 側の処理 / Work done in Illustrator and Photoshop
    // =========================================
    // toString() で送るので、関数の前後と中に JSDoc・コメントを置かない。ASCII のみ。
    // aiWorker は Illustrator 版 ConvertDocumentToGrayscale と同じ処理の写し、psWorker は Photoshop 用
    // Sent with toString(): no JSDoc or comments around or inside, ASCII only.
    // aiWorker copies the work of the Illustrator script ConvertDocumentToGrayscale; psWorker is for Photoshop

    var WORKER_SEPARATOR = ";";

    function aiWorkerUnlockAndShowLayers(layers, records) {
        for (var i = 0; i < layers.length; i++) {
            var layer = layers[i];
            if (!layer.visible) {
                layer.visible = true;
                records.push({ target: layer, key: "visible", value: false });
            }
            if (layer.locked) {
                layer.locked = false;
                records.push({ target: layer, key: "locked", value: true });
            }
            aiWorkerUnlockAndShowLayers(layer.layers, records);
        }
    }

    function aiWorkerUnlockAndShowItems(doc, records) {
        var pageItems = doc.pageItems;
        for (var i = 0; i < pageItems.length; i++) {
            var item = pageItems[i];
            try {
                if (item.hidden) {
                    item.hidden = false;
                    records.push({ target: item, key: "hidden", value: true });
                }
                if (item.locked) {
                    item.locked = false;
                    records.push({ target: item, key: "locked", value: true });
                }
            } catch (e) {
                continue;
            }
        }
    }

    function aiWorkerRestoreStates(records) {
        for (var i = records.length - 1; i >= 0; i--) {
            try {
                records[i].target[records[i].key] = records[i].value;
            } catch (e) {
                continue;
            }
        }
    }

    function aiWorkerMoveLayerContentsIntoGroup(layer, artGroup) {
        var items = [];
        for (var i = 0; i < layer.pageItems.length; i++) {
            if (layer.pageItems[i] !== artGroup) items.push(layer.pageItems[i]);
        }
        for (var j = 0; j < items.length; j++) {
            items[j].move(artGroup, ElementPlacement.PLACEATEND);
        }
        for (var k = 0; k < layer.layers.length; k++) {
            aiWorkerMoveLayerContentsIntoGroup(layer.layers[k], artGroup);
        }
    }

    function aiWorkerExpandSymbolArt(doc, symbol) {
        var workLayer = doc.layers.add();
        var instance = workLayer.symbolItems.add(symbol);
        instance.breakLink();
        var artGroup = workLayer.groupItems.add();
        aiWorkerMoveLayerContentsIntoGroup(workLayer, artGroup);
        return { layer: workLayer, art: artGroup };
    }

    function aiWorkerCollectNestedItems(container, itemTypename, found) {
        for (var i = 0; i < container.pageItems.length; i++) {
            var item = container.pageItems[i];
            if (item.typename === itemTypename) {
                found.push(item);
            } else if (item.typename === "GroupItem") {
                aiWorkerCollectNestedItems(item, itemTypename, found);
            }
        }
        return found;
    }

    function aiWorkerConvertSymbol(doc, symbolName, converted, completedNames, keepWhiteText) {
        if (converted.hasOwnProperty(symbolName)) return converted[symbolName];
        converted[symbolName] = null;
        var expanded = null;
        try {
            expanded = aiWorkerExpandSymbolArt(doc, doc.symbols.getByName(symbolName));
            var nestedItems = aiWorkerCollectNestedItems(expanded.art, "SymbolItem", []);
            for (var i = 0; i < nestedItems.length; i++) {
                var nestedSymbol = aiWorkerConvertSymbol(doc, nestedItems[i].symbol.name, converted, completedNames, keepWhiteText);
                if (nestedSymbol) nestedItems[i].symbol = nestedSymbol;
            }
            doc.selection = null;
            expanded.art.selected = true;
            app.executeMenuCommand("Colors7");
            doc.selection = null;
            aiWorkerApplyK100ToTextFrames(aiWorkerCollectNestedItems(expanded.art, "TextFrame", []), aiWorkerCreateK100Color(), keepWhiteText);
            converted[symbolName] = doc.symbols.add(expanded.art);
            completedNames.push(symbolName);
        } catch (e) {
            converted[symbolName] = null;
        } finally {
            if (expanded) {
                try {
                    expanded.layer.remove();
                } catch (removeError) {
                    expanded = null;
                }
            }
        }
        return converted[symbolName];
    }

    function aiWorkerConvertAllSymbols(doc, keepWhiteText) {
        if (!doc.symbols.length) return;
        var originalLayer = doc.activeLayer;
        var symbolNames = [];
        for (var i = 0; i < doc.symbols.length; i++) {
            symbolNames.push(doc.symbols[i].name);
        }
        var instancesByName = {};
        for (var j = 0; j < doc.symbolItems.length; j++) {
            try {
                var instance = doc.symbolItems[j];
                var instanceSymbolName = instance.symbol.name;
                if (!instancesByName.hasOwnProperty(instanceSymbolName)) instancesByName[instanceSymbolName] = [];
                instancesByName[instanceSymbolName].push(instance);
            } catch (e) {
                continue;
            }
        }
        var converted = {};
        var completedNames = [];
        for (var k = 0; k < symbolNames.length; k++) {
            if (instancesByName.hasOwnProperty(symbolNames[k])) aiWorkerConvertSymbol(doc, symbolNames[k], converted, completedNames, keepWhiteText);
        }
        for (var n = completedNames.length - 1; n >= 0; n--) {
            var symbolName = completedNames[n];
            var newSymbol = converted[symbolName];
            var instances = instancesByName[symbolName] || [];
            for (var m = 0; m < instances.length; m++) {
                try {
                    instances[m].symbol = newSymbol;
                } catch (e) {
                    continue;
                }
            }
            try {
                doc.symbols.getByName(symbolName).remove();
                newSymbol.name = symbolName;
            } catch (e) {
                try {
                    newSymbol.name = symbolName + " (Grayscale)";
                } catch (renameError) {
                    newSymbol = null;
                }
            }
        }
        aiWorkerRemoveUnusedSymbols(doc, symbolNames, converted);
        try {
            doc.activeLayer = originalLayer;
        } catch (e) {
            originalLayer = null;
        }
    }

    function aiWorkerRemoveUnusedSymbols(doc, symbolNames, converted) {
        var pendingNames = [];
        for (var i = 0; i < symbolNames.length; i++) {
            if (!converted.hasOwnProperty(symbolNames[i])) pendingNames.push(symbolNames[i]);
        }
        var removedAny = true;
        while (pendingNames.length && removedAny) {
            removedAny = false;
            var remainingNames = [];
            for (var j = 0; j < pendingNames.length; j++) {
                try {
                    doc.symbols.getByName(pendingNames[j]).remove();
                    removedAny = true;
                } catch (e) {
                    remainingNames.push(pendingNames[j]);
                }
            }
            pendingNames = remainingNames;
        }
    }

    function aiWorkerCreateK100Color() {
        var blackColor = new CMYKColor();
        blackColor.black = 100;
        return blackColor;
    }

    function aiWorkerIsWhiteColor(color) {
        if (color.typename === "GrayColor") return color.gray === 0;
        if (color.typename === "CMYKColor") return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
        if (color.typename === "RGBColor") return color.red === 255 && color.green === 255 && color.blue === 255;
        return false;
    }

    function aiWorkerShouldBecomeK100(color, keepWhiteText) {
        var solidTypes = { "GrayColor": true, "CMYKColor": true, "RGBColor": true, "SpotColor": true, "LabColor": true };
        if (!color || !solidTypes[color.typename]) return false;
        return !(keepWhiteText && aiWorkerIsWhiteColor(color));
    }

    function aiWorkerApplyColorToCharacterRuns(characters, targets, colorKey, color) {
        var i = 0;
        while (i < targets.length) {
            if (!targets[i]) {
                i++;
                continue;
            }
            var runStart = i;
            while (i < targets.length && targets[i]) i++;
            var characterRun = characters[runStart];
            characterRun.length = i - runStart;
            characterRun.characterAttributes[colorKey] = color;
        }
    }

    function aiWorkerApplyK100ToTextFrames(textFrames, blackColor, keepWhiteText) {
        for (var i = 0; i < textFrames.length; i++) {
            try {
                var characters = textFrames[i].characters;
                var fillTargets = [];
                var strokeTargets = [];
                for (var j = 0; j < characters.length; j++) {
                    var attributes = characters[j].characterAttributes;
                    fillTargets.push(aiWorkerShouldBecomeK100(attributes.fillColor, keepWhiteText));
                    strokeTargets.push(aiWorkerShouldBecomeK100(attributes.strokeColor, keepWhiteText));
                }
                aiWorkerApplyColorToCharacterRuns(characters, fillTargets, "fillColor", blackColor);
                aiWorkerApplyColorToCharacterRuns(characters, strokeTargets, "strokeColor", blackColor);
            } catch (e) {
                continue;
            }
        }
    }

    function aiWorkerDeleteAllSwatches(doc) {
        for (var i = doc.swatches.length - 1; i >= 0; i--) {
            try {
                doc.swatches[i].remove();
            } catch (e) {
                continue;
            }
        }
        for (var j = doc.swatchGroups.length - 1; j >= 0; j--) {
            try {
                if (doc.swatchGroups[j].getAllSwatches().length === 0) doc.swatchGroups[j].remove();
            } catch (groupError) {
                continue;
            }
        }
    }

    function aiWorkerRunOnAllArtwork(doc, menuCommandId) {
        app.executeMenuCommand("selectall");
        if (doc.selection.length) app.executeMenuCommand(menuCommandId);
        doc.selection = null;
    }

    function aiWorkerConvertDocument(doc, keepWhiteText) {
        try {
            app.selectTool("Adobe Select Tool");
        } catch (e) {
            doc.selection = null;
        }
        doc.selection = null;
        var stateRecords = [];
        aiWorkerUnlockAndShowLayers(doc.layers, stateRecords);
        aiWorkerUnlockAndShowItems(doc, stateRecords);
        try {
            aiWorkerRunOnAllArtwork(doc, "Colors8");
            if (doc.documentColorSpace !== DocumentColorSpace.CMYK) app.executeMenuCommand("doc-color-cmyk");
            aiWorkerConvertAllSymbols(doc, keepWhiteText);
            aiWorkerRunOnAllArtwork(doc, "Colors7");
            aiWorkerApplyK100ToTextFrames(doc.textFrames, aiWorkerCreateK100Color(), keepWhiteText);
        } finally {
            aiWorkerRestoreStates(stateRecords);
        }
        aiWorkerDeleteAllSwatches(doc);
    }

    function aiWorkerConvertFile(sourceUri, outputUri, keepWhiteText) {
        var resultText = "OK";
        var sourceFile = new File(decodeURI(sourceUri));
        var outputFile = new File(decodeURI(outputUri));
        var savedInteraction = app.userInteractionLevel;
        var openedDoc = null;
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        try {
            var documentCountBefore = app.documents.length;
            openedDoc = app.open(sourceFile);
            if (app.documents.length === documentCountBefore) {
                openedDoc = null;
                resultText = "ERR:ALREADY_OPEN";
            } else {
                aiWorkerConvertDocument(openedDoc, keepWhiteText);
                if (outputFile.name.toLowerCase().slice(-4) === ".eps") {
                    openedDoc.saveAs(outputFile, new EPSSaveOptions());
                } else {
                    var aiSaveOptions = new IllustratorSaveOptions();
                    aiSaveOptions.pdfCompatible = true;
                    openedDoc.saveAs(outputFile, aiSaveOptions);
                }
            }
        } catch (e) {
            resultText = "ERR:" + encodeURI(String(e));
        } finally {
            if (openedDoc) {
                try {
                    openedDoc.close(SaveOptions.DONOTSAVECHANGES);
                } catch (closeError) {
                    openedDoc = null;
                }
            }
            app.userInteractionLevel = savedInteraction;
        }
        return resultText;
    }


    function psWorkerSetTextLayersBlack(layerContainer, keepWhiteText) {
        for (var i = 0; i < layerContainer.artLayers.length; i++) {
            var artLayer = layerContainer.artLayers[i];
            if (artLayer.kind !== LayerKind.TEXT) continue;
            var wasLocked = artLayer.allLocked;
            try {
                if (wasLocked) artLayer.allLocked = false;
                var isWhite = Math.round(artLayer.textItem.color.gray.gray) === 0;
                if (!(keepWhiteText && isWhite)) {
                    var blackColor = new SolidColor();
                    blackColor.gray.gray = 100;
                    artLayer.textItem.color = blackColor;
                }
            } catch (e) {
                isWhite = false;
            }
            if (wasLocked) artLayer.allLocked = true;
        }
        for (var j = 0; j < layerContainer.layerSets.length; j++) {
            psWorkerSetTextLayersBlack(layerContainer.layerSets[j], keepWhiteText);
        }
    }

    function psWorkerConvertFile(sourceUri, outputUri, keepWhiteText) {
        var resultText = "OK";
        var sourceFile = new File(decodeURI(sourceUri));
        var outputFile = new File(decodeURI(outputUri));
        var savedDialogs = app.displayDialogs;
        var openedDoc = null;
        app.displayDialogs = DialogModes.NO;
        try {
            var documentCountBefore = app.documents.length;
            openedDoc = app.open(sourceFile);
            if (app.documents.length === documentCountBefore) {
                openedDoc = null;
                resultText = "ERR:ALREADY_OPEN";
            } else {
                app.activeDocument = openedDoc;
                if (openedDoc.mode !== DocumentMode.GRAYSCALE) openedDoc.changeMode(ChangeMode.GRAYSCALE);
                psWorkerSetTextLayersBlack(openedDoc, keepWhiteText);
                var psdSaveOptions = new PhotoshopSaveOptions();
                psdSaveOptions.layers = true;
                psdSaveOptions.embedColorProfile = true;
                psdSaveOptions.alphaChannels = true;
                psdSaveOptions.spotColors = true;
                openedDoc.saveAs(outputFile, psdSaveOptions, false, Extension.LOWERCASE);
            }
        } catch (e) {
            resultText = "ERR:" + encodeURI(String(e));
        } finally {
            if (openedDoc) {
                try {
                    openedDoc.close(SaveOptions.DONOTSAVECHANGES);
                } catch (closeError) {
                    openedDoc = null;
                }
            }
            app.displayDialogs = savedDialogs;
        }
        return resultText;
    }

    var AI_WORKER_FUNCS = [
        aiWorkerUnlockAndShowLayers, aiWorkerUnlockAndShowItems, aiWorkerRestoreStates,
        aiWorkerMoveLayerContentsIntoGroup, aiWorkerExpandSymbolArt, aiWorkerCollectNestedItems,
        aiWorkerConvertSymbol, aiWorkerConvertAllSymbols, aiWorkerRemoveUnusedSymbols, aiWorkerCreateK100Color, aiWorkerIsWhiteColor,
        aiWorkerShouldBecomeK100, aiWorkerApplyColorToCharacterRuns, aiWorkerApplyK100ToTextFrames,
        aiWorkerDeleteAllSwatches, aiWorkerRunOnAllArtwork, aiWorkerConvertDocument, aiWorkerConvertFile
    ];

    var PS_WORKER_FUNCS = [psWorkerSetTextLayersBlack, psWorkerConvertFile];

    /* 拡張子ごとの送り先と送る関数 / Target app and worker functions per file type */
    var ILLUSTRATOR_CONVERTER = { target: ILLUSTRATOR_TARGET, appName: "Illustrator", workerFuncs: AI_WORKER_FUNCS, entryName: "aiWorkerConvertFile" };
    var PHOTOSHOP_CONVERTER = { target: PHOTOSHOP_TARGET, appName: "Photoshop", workerFuncs: PS_WORKER_FUNCS, entryName: "psWorkerConvertFile" };

    /**
     * 1つのファイルを変換する送り先用のスクリプトを組み立てる
     * @param {Object} converter - getConverterForFile の戻り値
     * @param {File} sourceFile - 元のファイル
     * @param {File} outputFile - 保存先のファイル
     * @returns {string} BridgeTalk の本文
     */
    function buildWorkerScript(converter, sourceFile, outputFile) {
        var lineBreak = WORKER_SEPARATOR + String.fromCharCode(10);
        var functionSources = [];
        for (var i = 0; i < converter.workerFuncs.length; i++) {
            functionSources.push(converter.workerFuncs[i].toString());
        }
        /* パスは encodeURI で ASCII にして渡す（日本語は BridgeTalk で化けるため） / Pass paths as ASCII via encodeURI */
        return functionSources.join(lineBreak) + lineBreak +
            converter.entryName + "(\"" + encodeURI(sourceFile.fsName) + "\", \"" + encodeURI(outputFile.fsName) + "\", " + (KEEP_WHITE_TEXT ? "true" : "false") + ");";
    }

    // =========================================
    // 進行状況 / Progress
    // =========================================

    /**
     * 進行状況バーの小さなパレットを表示する
     * @param {number} maxValue - バーの最大値
     * @returns {{setMessage: function(string): void, update: function(number): void, close: function(): void}} 操作用の関数（close は2回呼んでもよい）
     */
    function createProgressWindow(maxValue) {
        var progressPalette = new Window("palette", getLabel(LABELS.dialog.progressTitle));
        setupWindow(progressPalette);
        progressPalette.alignChildren = ["fill", "top"];

        var messageText = progressPalette.add("statictext", undefined, "");
        /* あとから入れる文言のぶん、幅を先に取る / Reserve the width for the text set later */
        messageText.preferredSize.width = PROGRESS_BAR_WIDTH;
        var progressBar = progressPalette.add("progressbar", undefined, 0, maxValue);
        progressBar.preferredSize.width = PROGRESS_BAR_WIDTH;

        progressPalette.show();
        progressPalette.update();
        var isClosed = false;

        return {
            setMessage: function (messageString) {
                messageText.text = messageString;
                progressPalette.update();
            },
            update: function (value) {
                progressBar.value = value;
                progressPalette.update();
            },
            close: function () {
                if (!isClosed) {
                    isClosed = true;
                    progressPalette.close();
                }
            }
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントのカラーを K だけのグレーにし、テキストを K100 にしてスウォッチを整理する
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function convertDocumentToGrayscale(doc) {
        var unlockedTargets = unlockLayersAndItems(doc);
        try {
            /* 透明の分割・統合に使うカラースペースも CMYK に / Use CMYK as the transparency blend space */
            doc.transparencyPreferences.blendingSpace = BlendingSpace.CMYK;

            var lightnessByK = buildGrayLightnessTable(doc);
            var colors = collectConvertibleColors(doc);
            var grayValues = [];
            for (var i = 0; i < colors.length; i++) {
                try {
                    grayValues.push(convertColorToGray(colors[i], lightnessByK));
                } catch (e) {
                    /* 変換できないカラーは null（まとめない） / null for colors that cannot be converted (not merged) */
                    grayValues.push(null);
                }
            }

            var blackColor = doc.colors.itemByName("Black");
            applyK100ToTextStyles(doc, blackColor);
            applyK100ToText(doc, blackColor);

            mergeSameGrayColors(doc, colors, grayValues);
            removeUnusedSwatches(doc);
        } finally {
            restoreLocks(unlockedTargets);
        }
    }

    /**
     * 送り先からの戻り値を、失敗の理由の文言にする
     * @param {string} resultText - 送り先からの戻り値（"ERR:…"）
     * @param {string} appName - アプリ名
     * @returns {string} 理由
     */
    function describeFailure(resultText, appName) {
        var reason = resultText.substring(4);
        if (reason === "ALREADY_OPEN") return getLabel(LABELS.alert.alreadyOpen, { app: appName });
        if (reason === "TIMEOUT") return getLabel(LABELS.alert.timeout);
        if (reason === "NOT_LAUNCHED") return getLabel(LABELS.alert.notLaunched, { app: appName });
        var decodedReason = reason;
        try {
            decodedReason = decodeURI(reason);
        } catch (e) {
            decodedReason = reason;
        }
        return decodedReason;
    }

    /**
     * 1つのファイルを送り先のアプリで変換する（アプリは最初に使うときに起動する）
     * @param {Object} linkGroup - collectTargetLinkGroups の1件
     * @param {File} outputFile - 保存先のファイル
     * @param {Object} launchedTargets - 起動を確かめた送り先（結果を書き込む）
     * @param {Object} progressWindow - createProgressWindow の戻り値
     * @returns {string} 送り先からの戻り値（"OK" か "ERR:…"）
     */
    function convertLinkedFile(linkGroup, outputFile, launchedTargets, progressWindow) {
        var converter = linkGroup.converter;
        if (!launchedTargets.hasOwnProperty(converter.target)) {
            if (!BridgeTalk.isRunning(converter.target)) progressWindow.setMessage(getLabel(LABELS.progress.launching, { app: converter.appName }));
            launchedTargets[converter.target] = ensureAppRunning(converter.target);
        }
        if (!launchedTargets[converter.target]) return "ERR:NOT_LAUNCHED";
        return sendAndWait(converter.target, buildWorkerScript(converter, new File(linkGroup.filePath), outputFile), FILE_TIMEOUT_SECONDS);
    }

    /**
     * リンクしたファイルをアプリで変換して別名で保存する（進行状況バーを表示）
     * @param {Object[]} linkGroups - collectTargetLinkGroups の戻り値
     * @param {Object} progressWindow - createProgressWindow の戻り値
     * @param {string[]} failureLines - 変換できなかったファイルの行（追加する）
     * @returns {Object[]} 変換できたファイルの {links, outputFile}
     */
    function convertLinkedFiles(linkGroups, progressWindow, failureLines) {
        var convertedGroups = [];
        var launchedTargets = {};
        for (var i = 0; i < linkGroups.length; i++) {
            var sourceName = decodeURI(new File(linkGroups[i].filePath).name);
            var outputFile = getOutputFile(linkGroups[i].filePath);
            progressWindow.setMessage(getLabel(LABELS.progress.converting, { current: i + 1, total: linkGroups.length, name: sourceName }));
            var resultText = convertLinkedFile(linkGroups[i], outputFile, launchedTargets, progressWindow);
            if (resultText === "OK" && outputFile.exists) {
                convertedGroups.push({ links: linkGroups[i].links, outputFile: outputFile });
            } else {
                failureLines.push(sourceName + " : " + (resultText.indexOf("ERR:") === 0 ? describeFailure(resultText, linkGroups[i].converter.appName) : resultText));
            }
            progressWindow.update(i + 1);
        }
        return convertedGroups;
    }

    /**
     * 変換したファイルにリンクを差し替える
     * @param {Object[]} convertedGroups - convertLinkedFiles の戻り値
     * @param {string[]} failureLines - 差し替えられなかったリンクの行（追加する）
     * @returns {number} 差し替えたリンクの数
     */
    function relinkConvertedFiles(convertedGroups, failureLines) {
        var relinkedCount = 0;
        for (var i = 0; i < convertedGroups.length; i++) {
            for (var j = 0; j < convertedGroups[i].links.length; j++) {
                try {
                    convertedGroups[i].links[j].relink(convertedGroups[i].outputFile);
                    relinkedCount++;
                } catch (e) {
                    failureLines.push(decodeURI(convertedGroups[i].outputFile.name) + " : " + e);
                }
            }
        }
        return relinkedCount;
    }

    /**
     * ドキュメントのカラーをモノクロにする。リンクした AI・EPS・PSD があれば、確認してから一緒に変換して差し替える
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) return;
        var doc = app.activeDocument;

        /* リンクを変換するのは、ファイルを書き出すので確認で［はい］のときだけ / Linked files are written, so ask first */
        var linkGroups = collectTargetLinkGroups(doc);
        var includeLinks = linkGroups.length > 0 &&
            confirm(getLabel(LABELS.alert.confirmLinks, { count: linkGroups.length, suffix: OUTPUT_SUFFIX }));

        if (!includeLinks) {
            app.doScript(function () {
                convertDocumentToGrayscale(doc);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undoName));
            return;
        }

        /* バーはリンクしたファイルの数＋ドキュメントのカラーの1段 / One step per linked file plus one for the document */
        var progressWindow = createProgressWindow(linkGroups.length + 1);
        var failureLines = [];
        var convertedGroups = [];
        var relinkedCount = 0;
        try {
            convertedGroups = convertLinkedFiles(linkGroups, progressWindow, failureLines);

            /* リンクの差し替えとドキュメントのカラーの変換は、1回の取り消しで戻せるようにまとめる
               Relinking and the document conversion are undone together in one step */
            progressWindow.setMessage(getLabel(LABELS.progress.documentColors));
            app.doScript(function () {
                relinkedCount = relinkConvertedFiles(convertedGroups, failureLines);
                convertDocumentToGrayscale(doc);
            }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undoName));
            progressWindow.update(linkGroups.length + 1);
        } finally {
            progressWindow.close();
        }

        var reportText = getLabel(LABELS.alert.done, { count: convertedGroups.length, links: relinkedCount });
        if (failureLines.length) reportText += "\n\n" + getLabel(LABELS.alert.failedHeader) + "\n" + failureLines.join("\n");
        alert(reportText);
    }

    main();

})();
