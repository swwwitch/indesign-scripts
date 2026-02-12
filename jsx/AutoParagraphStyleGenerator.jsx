/*
    InDesign Script: [基本段落]および未設定段落からのスタイル自動生成 (v3 修正版)
*/

(function () {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        alert("ドキュメントを開いてください。");
        return;
    }

    var doc = app.activeDocument;
    
    // ターゲットとするスタイル名のリスト
    var targetStyleNames = ["[基本段落]", "基本段落", "[Basic Paragraph]", "No Paragraph Style"];

    // スタイル作成用の情報を格納するオブジェクト
    var styleGroups = {};
    
    // 集計用変数をここで宣言（エラー修正箇所）
    var countProcessed = 0;
    var countSkipped = 0;
    var styleCounter = 1; 

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
                // 属性を取得
                var fFamily = para.appliedFont.name;
                var fStyle = para.fontStyle;
                var fSize = para.pointSize;
                var fLeading = para.leading;

                // 属性が混合(Mixed)している場合はスキップ
                // (1つの段落内で文字サイズが違う箇所がある場合などは数値ではなくObject等を返すため)
                if (typeof fSize !== 'number' || (fLeading !== Leading.AUTO && typeof fLeading !== 'number')) {
                    countSkipped++;
                    continue; 
                }

                if (fLeading === Leading.AUTO) {
                    fLeading = "Auto";
                }

                // 識別キーを作成
                var key = fFamily + "_" + fStyle + "_" + fSize + "_" + fLeading;

                if (!styleGroups[key]) {
                    styleGroups[key] = {
                        fontFamily: para.appliedFont,
                        fontStyle: fStyle,
                        pointSize: fSize,
                        leading: fLeading,
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
                
                // 新しいスタイル名を生成
                var sizeLabel = Math.round(group.pointSize * 10) / 10;
                var newStyleName = "AutoStyle_" + styleCounter + "_" + sizeLabel + "Q"; 
                
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
                    pointSize: group.pointSize,
                    leading: (group.leading === "Auto" ? Leading.AUTO : group.leading)
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
