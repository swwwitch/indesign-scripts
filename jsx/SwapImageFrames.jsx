/target "InDesign"
var SCRIPT_VERSION = "v1.0";

/*
 * スクリプトの概要：
 * 選択した複数の画像入りフレームを、ダイアログで指定した方法で入れ替えます。
 * - 「画像リンクを入れ替え」：各フレームの既存画像を削除し、別フレームのリンク画像を順送りに再配置します。
 *   配置後のフィット方法として「フレームに合わせて塗り」または「全体を表示」を選択できます。
 * - 「フレームごと入れ替え」：各フレームの中身はそのままに、フレームの位置のみを順送りに入れ替えます。
 *   フレーム位置の基準は「左上」または「中央」から選択できます。
 * - 画像入りフレームを2つ以上選択している場合に実行できます。
 * - 入れ替え順は、フレームの見た目上の位置を基準に「上から下、左から右」の順で決まります。
 * - 同じフレームが重複選択されていても、1回だけ処理します。
 * - 「画像リンクを入れ替え」は、各フレームに主画像が1点だけある前提で動作します。
 * - 画像削除・再配置・フィットに失敗した場合は処理を中止し、エラーを表示します。
 * - 処理は1回のUndoで戻せます。
 *
 * Overview:
 * Swaps multiple selected frames containing placed images using the method chosen in the dialog.
 * - "Swap linked images": Removes the existing graphic in each frame, then places linked images from other frames in rotating order.
 *   After placement, you can choose either Fill Proportionally or Fit Proportionally.
 * - "Swap frames": Keeps each frame's existing content unchanged and swaps only the frame positions.
 *   The frame position reference can be set to Top Left or Center.
 * - Requires at least two selected frames containing placed images.
 * - The swap order is determined by the visible frame positions: top to bottom, then left to right.
 * - Even if the same frame is selected more than once, it is processed only once.
 * - "Swap linked images" assumes each frame contains a single primary graphic.
 * - If graphic removal, placement, or fitting fails, the operation is canceled and an error message is shown.
 * - The whole operation can be undone in a single Undo step.
 *
 * 作成日：2026-03-28
 * 更新日：2026-03-28
 */

function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var LANG = getCurrentLang();

var LABELS = {
    noDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    selectFrames: {
        ja: "画像入りフレームを2つ以上選択してください。",
        en: "Select at least two frames containing placed images."
    },
    selectLinkedFrames: {
        ja: "リンク画像が入ったフレームを2つ以上選択してください。",
        en: "Select at least two frames containing linked images."
    },
    sortError: {
        ja: "フレームの並び順を決定できませんでした。",
        en: "Could not determine the frame order."
    },
    graphicRemoveError: {
        ja: "既存画像の削除に失敗したため、画像リンクの入れ替えを中止しました。",
        en: "The swap was canceled because an existing graphic could not be removed."
    },
    graphicPlaceError: {
        ja: "画像の再配置に失敗したため、画像リンクの入れ替えを中止しました。",
        en: "The swap was canceled because a linked image could not be placed."
    },
    graphicFitError: {
        ja: "画像は再配置されましたが、フィット処理に失敗しました。",
        en: "The linked image was replaced, but the fit operation failed."
    },
    genericError: {
        ja: "予期しないエラーが発生したため、処理を中止しました。",
        en: "The operation was canceled due to an unexpected error."
    },
    undoName: {
        ja: "フレーム／画像の入れ替え",
        en: "Swap Frames / Images"
    },
    dialogTitle: {
        ja: "フレーム／画像の入れ替え",
        en: "Swap Frames / Images"
    },
    fitPanelTitle: {
        ja: "配置後のフィット",
        en: "Fit After Placing"
    },
    fitFillProportionally: {
        ja: "フレームに合わせて塗り（トリミングあり）",
        en: "Fill Proportionally (crop)"
    },
    fitProportionally: {
        ja: "全体を表示（余白あり）",
        en: "Fit Proportionally"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    swapModePanelTitle: {
        ja: "入れ替えモード",
        en: "Swap Mode"
    },
    swapByFrame: {
        ja: "フレームごと入れ替え",
        en: "Swap frames"
    },
    swapGraphicOnly: {
        ja: "画像リンクを入れ替え",
        en: "Swap linked images"
    },
    frameAnchorPanelTitle: {
        ja: "フレームの位置",
        en: "Frame Position"
    },
    frameAnchorTopLeft: {
        ja: "左上",
        en: "Top Left"
    },
    frameAnchorCenter: {
        ja: "中央",
        en: "Center"
    }
};

function L(key) {
    if (!LABELS[key]) return key;
    return LABELS[key][LANG] || LABELS[key].en;
}

function showSwapDialog() {
    var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.margins = 16;

    var pnlSwapMode = dlg.add("panel", undefined, L("swapModePanelTitle"));
    pnlSwapMode.orientation = "column";
    pnlSwapMode.alignChildren = ["left", "top"];
    pnlSwapMode.margins = [15, 20, 15, 10];

    var rdoSwapByFrame = pnlSwapMode.add("radiobutton", undefined, L("swapByFrame"));
    var rdoSwapGraphicOnly = pnlSwapMode.add("radiobutton", undefined, L("swapGraphicOnly"));
    rdoSwapGraphicOnly.value = true;

    var pnlAnchor = dlg.add("panel", undefined, L("frameAnchorPanelTitle"));
    pnlAnchor.orientation = "column";
    pnlAnchor.alignChildren = ["left", "top"];
    pnlAnchor.margins = [15, 20, 15, 10];

    var rdoAnchorTopLeft = pnlAnchor.add("radiobutton", undefined, L("frameAnchorTopLeft"));
    var rdoAnchorCenter = pnlAnchor.add("radiobutton", undefined, L("frameAnchorCenter"));
    rdoAnchorTopLeft.value = true;

    var pnlFit = dlg.add("panel", undefined, L("fitPanelTitle"));
    pnlFit.orientation = "column";
    pnlFit.alignChildren = ["left", "top"];
    pnlFit.margins = [15, 20, 15, 10];

    var rdoFill = pnlFit.add("radiobutton", undefined, L("fitFillProportionally"));
    var rdoFit = pnlFit.add("radiobutton", undefined, L("fitProportionally"));
    rdoFill.value = true;

    function reflectDialogState() {
        var isGraphicOnly = rdoSwapGraphicOnly.value === true;

        pnlFit.enabled = isGraphicOnly;
        rdoFill.enabled = isGraphicOnly;
        rdoFit.enabled = isGraphicOnly;

        pnlAnchor.enabled = !isGraphicOnly;
        rdoAnchorTopLeft.enabled = !isGraphicOnly;
        rdoAnchorCenter.enabled = !isGraphicOnly;
    }

    rdoSwapByFrame.onClick = reflectDialogState;
    rdoSwapGraphicOnly.onClick = reflectDialogState;

    reflectDialogState();

    var grpButtons = dlg.add("group");
    grpButtons.alignment = ["right", "center"];
    grpButtons.add("button", undefined, L("cancel"), { name: "cancel" });
    grpButtons.add("button", undefined, L("ok"), { name: "ok" });

    if (dlg.show() !== 1) {
        return null;
    }

    return {
        swapMode: rdoSwapByFrame.value ? "frame" : "graphicOnly",
        fitOption: rdoFill.value ? FitOptions.FILL_PROPORTIONALLY : FitOptions.PROPORTIONALLY,
        anchorMode: rdoAnchorTopLeft.value ? "topLeft" : "center"
    };
}

(function () {
    var ROW_TOLERANCE = 12; // 同一行とみなすY差の許容値
    if (app.documents.length === 0) {
        alert(L("noDocument"));
        return;
    }

    var sel = app.selection;
    if (!sel || sel.length < 2) {
        alert(L("selectFrames"));
        return;
    }

    var frames = [];
    var files = [];
    var seenFrameIds = {};

    function getFrameBounds(frame) {
        var gb = frame.geometricBounds;
        return {
            top: gb[0],
            left: gb[1],
            bottom: gb[2],
            right: gb[3],
            width: gb[3] - gb[1],
            height: gb[2] - gb[0],
            centerX: (gb[1] + gb[3]) / 2,
            centerY: (gb[0] + gb[2]) / 2
        };
    }

    function getFrameKey(frame) {
        if (!frame) return "";

        try {
            if (frame.id !== undefined) {
                return String(frame.id);
            }
        } catch (e) { }

        try {
            var gb = frame.geometricBounds;
            return [frame.constructor.name, gb[0], gb[1], gb[2], gb[3]].join("|");
        } catch (e2) { }

        return "";
    }

    function sortFramesByVisualOrder(items) {
        var sorted = items.slice(0);

        sorted.sort(function (a, b) {
            var ab = getFrameBounds(a);
            var bb = getFrameBounds(b);
            var ay = ab.centerY;
            var by = bb.centerY;
            var ax = ab.left;
            var bx = bb.left;

            if (Math.abs(ay - by) > ROW_TOLERANCE) {
                return ay - by;
            }
            return ax - bx;
        });

        return sorted;
    }

    function getPrimaryGraphic(frame) {
        if (!frame || frame.allGraphics.length === 0) return null;
        return frame.allGraphics[0];
    }

    function removeExistingGraphics(frame) {
        if (!frame) return;

        while (frame.allGraphics.length > 0) {
            frame.allGraphics[0].remove();
        }
    }

    function replaceFrameGraphic(frame, file, fitOption) {
        var placedItems;

        if (!frame || !file) {
            throw new Error("Invalid frame or file.");
        }

        try {
            removeExistingGraphics(frame);
        } catch (e) {
            throw new Error(L("graphicRemoveError"));
        }

        placedItems = frame.place(file);
        if (!placedItems || placedItems.length === 0) {
            throw new Error("Place failed.");
        }

        try {
            frame.fit(fitOption);
        } catch (e) {
            throw new Error(L("graphicFitError"));
        }
    }

    function moveFrameToStoredPosition(frame, bounds, anchorMode) {
        if (!frame || !bounds) return;

        var current = getFrameBounds(frame);
        var dx, dy;

        if (anchorMode === "center") {
            dx = bounds.centerX - current.centerX;
            dy = bounds.centerY - current.centerY;
        } else {
            dx = bounds.left - current.left;
            dy = bounds.top - current.top;
        }

        frame.move(undefined, [dx, dy]);
    }

    function getFrameFromSelectionItem(item) {
        if (!item) return null;

        var name = item.constructor.name;

        if (name === "Rectangle" || name === "Oval" || name === "Polygon") {
            return item;
        }

        if (name === "Image" || name === "PDF" || name === "EPS" || name === "ImportedPage") {
            if (item.parent) {
                var parentName = item.parent.constructor.name;
                if (parentName === "Rectangle" || parentName === "Oval" || parentName === "Polygon") {
                    return item.parent;
                }
            }
        }

        if (item.parent) {
            var pname = item.parent.constructor.name;
            if (pname === "Rectangle" || pname === "Oval" || pname === "Polygon") {
                return item.parent;
            }
        }

        return null;
    }

    for (var i = 0; i < sel.length; i++) {
        var frame = getFrameFromSelectionItem(sel[i]);
        var frameKey;

        if (!frame) continue;
        if (frame.allGraphics.length === 0) continue;

        frameKey = getFrameKey(frame);
        if (frameKey && seenFrameIds[frameKey]) continue;

        if (frameKey) {
            seenFrameIds[frameKey] = true;
        }
        frames.push(frame);
    }

    if (frames.length > 1) {
        try {
            frames = sortFramesByVisualOrder(frames);
        } catch (e) {
            alert(L("sortError"));
            return;
        }
    }

    if (frames.length < 2) {
        alert(L("selectFrames"));
        return;
    }

    var dialogResult = showSwapDialog();
    if (dialogResult === null) {
        return;
    }

    var swapMode = dialogResult.swapMode;
    var fitOption = dialogResult.fitOption;
    var anchorMode = dialogResult.anchorMode;

    if (swapMode === "graphicOnly") {
        for (var k = 0; k < frames.length; k++) {
            var graphic = getPrimaryGraphic(frames[k]);
            var link = graphic ? graphic.itemLink : null;
            if (!link || !link.filePath) {
                alert(L("selectLinkedFrames"));
                return;
            }
            files.push(File(link.filePath));
        }
    }

    try {
        app.doScript(function () {
            if (swapMode === "graphicOnly") {
                var lastFile = files[files.length - 1];
                var j;

                for (j = files.length - 1; j > 0; j--) {
                    replaceFrameGraphic(frames[j], files[j - 1], fitOption);
                }

                replaceFrameGraphic(frames[0], lastFile, fitOption);
                return;
            }

            var originalBounds = [];
            var lastBounds;

            for (var j = 0; j < frames.length; j++) {
                originalBounds.push(getFrameBounds(frames[j]));
            }

            lastBounds = originalBounds[originalBounds.length - 1];

            for (var j = frames.length - 1; j > 0; j--) {
                moveFrameToStoredPosition(frames[j], originalBounds[j - 1], anchorMode);
            }

            moveFrameToStoredPosition(frames[0], lastBounds, anchorMode);

        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, L("undoName"));
    } catch (e) {
        var msg = e && e.message ? e.message : "";

        if (msg === L("graphicFitError")) {
            alert(L("graphicFitError"));
        } else if (msg === "Place failed.") {
            alert(L("graphicPlaceError"));
        } else if (msg === "Invalid frame or file.") {
            alert(L("graphicPlaceError"));
        } else if (msg === L("graphicRemoveError")) {
            alert(L("graphicRemoveError"));
        } else {
            alert(L("genericError"));
        }
    }

})();
