//=========================================================
// 段落ごとにテキストフレームを分割（位置・書式・幅保持版）
// 更新日：2026-03-16
//
// 概要：
// 選択したテキストフレーム内の各段落を、段落ごとに独立した
// テキストフレームへ分割します。
//
// 特徴：
// ・元段落のベースライン位置を基準にY位置を補正
// ・元フレームの左右座標を維持し、幅を保持
// ・テキストフレーム設定（textFramePreferences）を引き継ぎ
// ・オーバーセットテキストがある場合は確認ダイアログを表示
// ・必要に応じてフレーム高さを自動調整してオーバーセットを解消
//
// 仕様上の注意：
// ・空段落（空行）は出力対象から除外されます
// ・空行は新しいフレームとして作成されず、そのまま削除されます
//
// 想定用途：
// 見出しや段落を個別オブジェクトとして配置・アニメーション・
// レイアウト調整するための前処理スクリプト
//=========================================================

#target indesign

function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var LANG = getCurrentLang();

var LABELS = {
    docRequired: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
    selectFrame: { ja: "テキストフレームを1つだけ選択して実行してください。", en: "Select exactly one text frame before running." },
    overflowTitle: { ja: "オーバーセットテキストの確認", en: "Overset Text Detected" },
    overflowMsg: { ja: "オーバーセットテキストがあります。処理方法を選択してください。", en: "The text frame contains overset text. Choose how to proceed." },
    expand: { ja: "フレームを拡張して解消する", en: "Expand frame to resolve" },
    ignore: { ja: "そのまま実行する（オーバーセットは消失します）", en: "Run anyway (overset text will be lost)" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" },
    overflowFail: { ja: "オーバーセットテキストを解消できなかったため中断しました。", en: "Could not resolve overset text. Process cancelled." }
};

function L(key){ return LABELS[key][LANG]; }

if (app.documents.length === 0) {
    alert(L("docRequired"));
    exit();
}

if (app.selection.length !== 1) {
    alert(L("selectFrame"));
    exit();
}

var sourceFrame = app.selection[0];
if (!sourceFrame || sourceFrame.constructor.name !== "TextFrame") {
    alert(L("selectFrame"));
    exit();
}

if (sourceFrame.overflows) {
    var overflowChoice = askOverflowHandling();
    if (overflowChoice === "cancel") {
        exit();
    }
    if (overflowChoice === "expand") {
        resolveOverflowByExpandingFrame(sourceFrame);
        if (sourceFrame.overflows) {
            alert(L("overflowFail"));
            exit();
        }
    }
}

function askOverflowHandling() {
    var dlg = new Window("dialog", L("overflowTitle"));
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.margins = [15, 15, 15, 10];

    dlg.add("statictext", undefined, L("overflowMsg"));

    var pnl = dlg.add("panel");
    pnl.orientation = "column";
    pnl.alignChildren = ["left", "top"];
    pnl.margins = [12, 12, 12, 10];

    var rbExpand = pnl.add("radiobutton", undefined, L("expand"));
    var rbIgnore = pnl.add("radiobutton", undefined, L("ignore"));
    rbExpand.value = true;

    var btns = dlg.add("group");
    btns.alignment = "right";
    var cancelBtn = btns.add("button", undefined, L("cancel"), {name:"cancel"});
    var okBtn = btns.add("button", undefined, L("ok"), {name:"ok"});

    if (dlg.show() !== 1) {
        return "cancel";
    }
    return rbExpand.value ? "expand" : "ignore";
}

function resolveOverflowByExpandingFrame(tf) {
    var maxGrowCount = 200;
    var step = 12;
    var count = 0;

    while (tf.overflows && count < maxGrowCount) {
        var gb = tf.geometricBounds;
        tf.geometricBounds = [gb[0], gb[1], gb[2] + step, gb[3]];
        count++;
    }
}

var parent = sourceFrame.parent;
var paras = sourceFrame.paragraphs.everyItem().getElements();

// 元のフレームの座標 [上, 左, 下, 右] を取得
var sourceBounds = sourceFrame.geometricBounds;
var origX1 = sourceBounds[1]; // 元のフレームの左端
var origX2 = sourceBounds[3]; // 元のフレームの右端

for (var i = 0; i < paras.length; i++) {
    var p = paras[i];

    // 空行はスキップ（仕様）
    if (p.contents.replace(/\s/g, '') === "") {
        continue;
    }

    // 念のため、行を持たない段落はスキップ
    if (p.lines.length === 0) {
        continue;
    }

    // 元のテキストの正確なY座標（ベースライン）を取得
    var origY = p.lines[0].baseline;

    // 新しいテキストフレームを作成
    var newFrame = parent.textFrames.add();
    newFrame.textFramePreferences.properties = sourceFrame.textFramePreferences.properties;

    // 一旦元のフレームと全く同じサイズ・位置にする
    newFrame.geometricBounds = sourceBounds;

    // 段落テキストを複製
    p.duplicate(LocationOptions.AT_BEGINNING, newFrame.insertionPoints.item(0));

    // 余分な空段落や改行を削除
    if (newFrame.paragraphs.length > 1) {
        newFrame.paragraphs.item(-1).remove();
    }
    if (newFrame.characters.length > 0 && newFrame.characters.item(-1).contents === "\r") {
        newFrame.characters.item(-1).remove();
    }

    // いったん、フレームをコンテンツに合わせる（高さと幅が縮む）
    newFrame.fit(FitOptions.FRAME_TO_CONTENT);

    if (newFrame.lines.length > 0) {
        // 縮んだ後のY座標を取得し、本来の位置とのズレを計算
        var newY = newFrame.lines[0].baseline;
        var dy = origY - newY;

        var b = newFrame.geometricBounds;

        // Y座標（上下）はズレを補正し、X座標（左右）は元のフレーム幅へ戻す
        newFrame.geometricBounds = [
            b[0] + dy,
            origX1,
            b[2] + dy,
            origX2
        ];

        // 幅を戻したことで再改行され、まれに高さ不足になるケースを救済
        var growCount = 0;
        while (newFrame.overflows && growCount < 20) {
            var gb = newFrame.geometricBounds;
            newFrame.geometricBounds = [gb[0], gb[1], gb[2] + 12, gb[3]];
            growCount++;
        }
    } else {
        newFrame.remove();
    }
}

// 元のテキストフレームを削除
sourceFrame.remove();