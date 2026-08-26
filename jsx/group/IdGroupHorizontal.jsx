#target indesign

/*

### 概要

選択したオブジェクトを縦位置の近さで「行」に分け、行ごとにグループ化します。

詳細は README を参照してください。

### Overview

Splits the selected objects into rows by vertical proximity and groups each row.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IdGroupHorizontal";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-04-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-ja/IdGroupHorizontal.md
// README (English)
// https://github.com/swwwitch/indesign-scripts/blob/main/readme-en/IdGroupHorizontal.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 同じ行とみなす垂直方向のズレの許容値（現在のルーラー単位）/ Vertical tolerance that still counts as the same row (current ruler units) */
    var ROW_TOLERANCE = 5;

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

    var currentLang = getCurrentLang();

    var LABELS = {
        alert: {
            noSelection: { ja: "アイテムを選択してください。", en: "Please select one or more items." },
            resultSuffix: { ja: " 個のグループを作成しました。", en: " group(s) created." }
        },
        undo: {
            groupRows: { ja: "横並びのアイテムをグループ化", en: "Group horizontally aligned items" }
        }
    };

    /**
     * ラベルを現在の言語で取得する
     * @param {object} labelEntry ja / en を持つラベルオブジェクト
     * @returns {string} 現在の言語のラベル文字列
     */
    function localize(labelEntry) {
        return labelEntry[currentLang];
    }

    // =========================================
    // 行の判定 / Row detection
    // =========================================

    /**
     * 選択オブジェクトを垂直方向の中心座標つきの配列に変換する
     * @param {Array} selectedItems 選択オブジェクトの配列
     * @returns {Array<{pageItem: PageItem, centerY: number}>} 中心 Y 座標を添えた配列
     */
    function collectItemsWithCenterY(selectedItems) {
        var itemsWithCenterY = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var pageItem = selectedItems[i];
            /* geometricBounds は [上, 左, 下, 右] / geometricBounds is [top, left, bottom, right] */
            var bounds = pageItem.geometricBounds;
            itemsWithCenterY.push({
                pageItem: pageItem,
                centerY: (bounds[0] + bounds[2]) / 2
            });
        }
        return itemsWithCenterY;
    }

    /**
     * 中心 Y 座標の近さでオブジェクトを行単位にまとめる
     * @param {Array<{pageItem: PageItem, centerY: number}>} itemsWithCenterY 中心 Y 座標を添えた配列
     * @param {number} tolerance 同じ行とみなす許容値
     * @returns {Array<Array<{pageItem: PageItem, centerY: number}>>} 行ごとの配列
     */
    function buildRows(itemsWithCenterY, tolerance) {
        itemsWithCenterY.sort(function(a, b) {
            return a.centerY - b.centerY;
        });

        var rows = [];
        var currentRow = [];
        for (var i = 0; i < itemsWithCenterY.length; i++) {
            if (currentRow.length === 0) {
                currentRow.push(itemsWithCenterY[i]);
            } else if (Math.abs(itemsWithCenterY[i].centerY - currentRow[0].centerY) <= tolerance) {
                currentRow.push(itemsWithCenterY[i]);
            } else {
                rows.push(currentRow);
                currentRow = [itemsWithCenterY[i]];
            }
        }
        if (currentRow.length > 0) rows.push(currentRow);
        return rows;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトを行ごとにグループ化する
     * @returns {void}
     */
    function main() {
        if (app.selection.length === 0) {
            alert(localize(LABELS.alert.noSelection));
            return;
        }

        var activeDoc = app.activeDocument;
        var rows = buildRows(collectItemsWithCenterY(app.selection), ROW_TOLERANCE);

        var createdGroupCount = 0;
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            /* グループ化には 2 つ以上のオブジェクトが必要 / Grouping requires at least two items */
            if (row.length < 2) continue;

            var itemsToGroup = [];
            for (var j = 0; j < row.length; j++) {
                itemsToGroup.push(row[j].pageItem);
            }
            activeDoc.groups.add(itemsToGroup);
            createdGroupCount++;
        }

        alert(createdGroupCount + localize(LABELS.alert.resultSuffix));
    }

    /* 一括で取り消せるように doScript でまとめて実行 / Run through doScript so the whole run is a single undo step */
    app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, localize(LABELS.undo.groupRows));

})();
