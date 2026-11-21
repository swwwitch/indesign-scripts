#target indesign

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.indesign.IdConvertDocumentToGrayscale
Meta-Default-language=ja
Name=ドキュメント全体を CMYK のモノクロにする（テキストは K100）
Name[en]=Turn the whole document into CMYK monochrome (text in K100)
Author=Masahiro Takano (@swwwitch)
Target-App=InDesign
Version=1.0.0
Release-Date=2026-11-21
Tags=document,ドキュメント
X-com.swwwitch-Script-Name=IdConvertDocumentToGrayscale
X-com.swwwitch-Updated=2026-11-21
X-com.swwwitch-README-JA=https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdConvertDocumentToGrayscale.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdConvertDocumentToGrayscale.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/indesign-scripts/blob/main/jsx/document/IdConvertDocumentToGrayscale.jsx
Description-BEGIN
[ja]
ドキュメントのカラーをすべて CMYK の K だけのグレーに変換し、テキストを K100 にします。最後に同じ濃さのスウォッチをまとめ、未使用のスウォッチを削除します。
[en]
Converts every color in the document to a CMYK gray that uses only K, and sets text to K100. Swatches of the same gray are merged and unused swatches are deleted at the end.
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
- 1回の取り消しで戻せます。

### Notes
- The transparency blend space is set to CMYK.
- Swatches with the same gray are merged into one named "C=0 M=0 Y=0 K=n" (K100 goes to [Black]) and unused swatches are deleted. Colors used in gradient stops or as the base of tints are kept.
- The gray is the K value whose lightness (L*) in Lab matches the color (converted with the document's color settings).
- Spot colors become process colors. Mixed inks are not converted.
- The contents of placed images and graphics are not converted.
- Solid fill and stroke colors of text become K100 ([Black]). White ([Paper], K0) stays white (KEEP_WHITE_TEXT).
- Locked layers and objects are unlocked while processing, then restored.
- A single Undo reverts everything.
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
        undoName: { ja: "ドキュメントを CMYK のモノクロにする", en: "Convert Document to CMYK Monochrome" }
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
     * 前面のドキュメントを対象に、1回の取り消しで戻せるように実行する
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) return;
        var doc = app.activeDocument;
        app.doScript(function () {
            convertDocumentToGrayscale(doc);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, getLabel(LABELS.undoName));
    }

    main();

})();
