/*
    File: AutoParagraphStyleGenerator.jsx
    InDesign Script: [基本段落]および未設定段落からのスタイル自動生成
    Version: v3.4
    更新日: 2026-02-13

    - 環境設定［単位と増減値］の「他の単位＞テキストサイズ」の単位に合わせて
      自動生成スタイル名のサイズ表記（pt / Q）を連動
    - 内部のグルーピング判定はpt基準で正規化し、pt/Q換算の揺れで別グループ化しにくくする
    - 参照・反映する属性はフォント（ファミリー/スタイル）＋サイズ＋行送りのみに限定
*/

(function () {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        alert("ドキュメントを開いてください。");
        return;
    }

    var doc = app.activeDocument;
    
    // --- Unit helpers (pt / Q) ---
    var PT_PER_MM = 72 / 25.4;
    var PT_PER_Q = PT_PER_MM * 0.25; // 1Q = 0.25mm

    var PT_PER_H = PT_PER_Q; // 1H = 0.25mm (和文組版のH)

    function isUnitQ(unitEnum) {
        try { return unitEnum === MeasurementUnits.Q; } catch (_) { return false; }
    }

    function isUnitH(unitEnum) {
        // InDesignの環境により MeasurementUnits.HA がHを表すケースがある
        try {
            if (typeof MeasurementUnits.HA !== 'undefined' && unitEnum === MeasurementUnits.HA) return true;
        } catch (_) { }
        return false;
    }

    function toPtFromUnit(value, unitEnum) {
        // pointSize / leading 等の取得値は「現在の単位での数値」になり得るため
        // 単位に応じてptへ正規化する
        if (typeof value !== 'number') return value;
        if (isUnitQ(unitEnum)) return value * PT_PER_Q;
        if (isUnitH(unitEnum)) return value * PT_PER_H;
        return value; // pt前提
    }

    function fromPtToUnit(ptValue, unitEnum) {
        // スタイル作成時は、プロパティが「現在の単位での数値」を期待することがあるため
        // ptから現在単位の数値に戻す
        if (typeof ptValue !== 'number') return ptValue;
        if (isUnitQ(unitEnum)) return ptValue / PT_PER_Q;
        if (isUnitH(unitEnum)) return ptValue / PT_PER_H;
        return ptValue; // pt前提
    }

    function roundTo(val, step) {
        return Math.round(val / step) * step;
    }

    function getTextSizeUnit(doc) {
        // 環境設定［単位と増減値］＞［他の単位］＞［テキストサイズ］に相当
        // ドキュメント側の viewPreferences で参照できる
        try {
            return doc.viewPreferences.textSizeMeasurementUnits;
        } catch (_) {
            return null;
        }
    }

    function formatSizeLabelFromPt(ptValue, unitEnum) {
        // 内部値(pt)から、表示用ラベルとサフィックスを返す
        // 返り値: { value: "12" or "12.5", suffix: "pt" or "Q" }
        if (isUnitQ(unitEnum)) {
            var q = ptValue / PT_PER_Q;
            var qRounded = roundTo(q, 0.1);
            return { value: (Math.round(qRounded * 10) / 10).toString(), suffix: "Q" };
        }
        // デフォルトはpt表記
        var ptRounded = roundTo(ptValue, 0.1);
        return { value: (Math.round(ptRounded * 10) / 10).toString(), suffix: "pt" };
    }

    // ターゲットとするスタイル名のリスト
    var targetStyleNames = ["[基本段落]", "基本段落", "[Basic Paragraph]", "No Paragraph Style"];

    // スタイル作成用の情報を格納するオブジェクト
    var styleGroups = {};
    
    // 集計用変数をここで宣言（エラー修正箇所）
    var countProcessed = 0;
    var countSkipped = 0;
    var styleCounter = 1; 

    // 単位は毎段落で都度参照しない（参照値はドキュメントの viewPreferences）
    var __textSizeUnit = getTextSizeUnit(doc);
    var __typoUnit = null;
    try { __typoUnit = doc.viewPreferences.typographicMeasurementUnits; } catch (_) { __typoUnit = null; }

    // --- 1. ドキュメント全体をスキャンしてグループ化 ---
    
    for (var i = 0; i < doc.stories.length; i++) {
        var story = doc.stories[i];
        
        for (var j = 0; j < story.paragraphs.length; j++) {
            var para = story.paragraphs[j];
            var pStyleName = para.appliedParagraphStyle.name;

            // 配列内に該当するスタイル名があるかチェック
            var isTarget = false;
            for (var k = 0; k < targetStyleNames.length; k++) {
                if (pStyleName === targetStyleNames[k]) {
                    isTarget = true;
                    break;
                }
            }

            if (isTarget) {
                // pointSize / leading は「現在の単位での数値」になり得るため、ptへ正規化して扱う
                // フォント情報
                var fFamily = para.appliedFont.name;
                var fStyle = para.fontStyle;

                var fSizeRaw = para.pointSize;
                var fLeadingRawOrAuto = para.leading;

                var fSizePt = toPtFromUnit(fSizeRaw, __textSizeUnit);
                var fLeadingPtOrAuto = fLeadingRawOrAuto;
                if (fLeadingRawOrAuto !== Leading.AUTO) {
                    // 行送りは「組版」単位に従うことが多い
                    fLeadingPtOrAuto = toPtFromUnit(fLeadingRawOrAuto, __typoUnit);
                }

                // 属性が混合(Mixed)している場合はスキップ
                // (1つの段落内で文字サイズが違う箇所がある場合などは数値ではなくObject等を返すため)
                if (
                    typeof fSizePt !== 'number' ||
                    (fLeadingPtOrAuto !== Leading.AUTO && typeof fLeadingPtOrAuto !== 'number')
                ) {
                    countSkipped++;
                    continue;
                }

                // 比較用にptを正規化（浮動小数の揺れで別グループ化しないように）
                var sizeNorm = roundTo(fSizePt, 0.01);

                var leadingPtNorm;
                var leadingKey;
                if (fLeadingPtOrAuto === Leading.AUTO) {
                    leadingPtNorm = Leading.AUTO;
                    leadingKey = "Auto";
                } else {
                    leadingPtNorm = roundTo(fLeadingPtOrAuto, 0.01);
                    leadingKey = leadingPtNorm;
                }

                // 識別キーを作成（内部はpt基準）
                var key = fFamily + "_" + fStyle + "_" + sizeNorm + "_" + leadingKey;

                if (!styleGroups[key]) {
                    styleGroups[key] = {
                        fontFamily: para.appliedFont,
                        fontStyle: fStyle,
                        pointSizePt: sizeNorm,
                        leadingPt: leadingPtNorm,
                        paragraphs: []
                    };
                }
                styleGroups[key].paragraphs.push(para);
            }
        }
    }

    // --- 2. スタイルの作成と適用 ---

    if (isObjectEmpty(styleGroups)) {
        var msgNoTarget = "対象となる段落が見つかりませんでした。";
        if (countSkipped > 0) {
            msgNoTarget += "\n(書式混在によりスキップされた段落: " + countSkipped + ")";
        }
        alert(msgNoTarget);
        return;
    }

    app.doScript(function() {
        for (var key in styleGroups) {
            if (styleGroups.hasOwnProperty(key)) {
                var group = styleGroups[key];
                
                // 新しいスタイル名を生成（テキストサイズ単位に連動）
                var sizeLabelObj = formatSizeLabelFromPt(group.pointSizePt, __textSizeUnit);
                var newStyleName = "AutoStyle_" + styleCounter + "_" + sizeLabelObj.value + sizeLabelObj.suffix;
                
                // スタイル名重複回避
                var finalStyleName = newStyleName;
                var suffix = 2;
                while (doc.paragraphStyles.itemByName(finalStyleName).isValid) {
                    finalStyleName = newStyleName + "_" + suffix;
                    suffix++;
                }

                // 段落スタイルを作成
                var newStyle = doc.paragraphStyles.add({
                    name: finalStyleName,
                    appliedFont: group.fontFamily,
                    fontStyle: group.fontStyle,
                    pointSize: fromPtToUnit(group.pointSizePt, __textSizeUnit),
                    leading: (group.leadingPt === Leading.AUTO ? Leading.AUTO : fromPtToUnit(group.leadingPt, __typoUnit))
                });

                // 該当する段落にスタイルを適用
                for (var p = 0; p < group.paragraphs.length; p++) {
                    var targetPara = group.paragraphs[p];
                    try {
                        // true = オーバーライド消去
                        targetPara.applyParagraphStyle(newStyle, true); 
                        countProcessed++;
                    } catch(e) {
                        // エラー無視
                    }
                }
                styleCounter++;
            }
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "段落スタイル自動整理");

    // 結果報告
    var msg = "完了しました。\n" +
              "作成したスタイル数: " + (styleCounter - 1) + "\n" +
              "適用した段落数: " + countProcessed;
    
    if (countSkipped > 0) {
        msg += "\n\n※書式混在でスキップ: " + countSkipped;
    }

    alert(msg);

    function isObjectEmpty(obj) {
        for (var prop in obj) {
            if (obj.hasOwnProperty(prop)) return false;
        }
        return true;
    }

})();
