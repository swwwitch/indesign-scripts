#target indesign

/*
    任意のスクリプトファイルを選択して実行するシンプルなランチャー

    - JSX / JSXBIN / JS スクリプトを選択して実行
    - 実行前にファイル存在と拡張子を確認
    - 実行エラー時にファイル名・行番号・エラー番号を表示
    - 日本語 / 英語ローカライズ対応
*/

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    dialogTitle: {
        ja: "実行するスクリプトを選択してください",
        en: "Select a script to execute"
    },
    fileNotFound: {
        ja: "選択したスクリプトファイルが存在しません:",
        en: "The selected script file does not exist:"
    },
    invalidFile: {
        ja: "実行できないファイル形式です:",
        en: "Unsupported file type:"
    },
    errorTitle: {
        ja: "スクリプトの実行中にエラーが発生しました:",
        en: "An error occurred while executing the script:"
    },
    file: {
        ja: "ファイル",
        en: "File"
    },
    line: {
        ja: "行番号",
        en: "Line"
    },
    errorNumber: {
        ja: "エラー番号",
        en: "Error Number"
    }
};

(function () {
    var DIALOG_TITLE = LABELS.dialogTitle[lang];
    var FILTER = "ExtendScript:*.jsx;*.jsxbin;*.js";

    function pickScriptFile(startFolder) {
        try {
            if (startFolder && startFolder.exists) {
                Folder.current = startFolder;
            }
        } catch (e) {}
        return File.openDialog(DIALOG_TITLE, FILTER, false);
    }

    function buildErrorText(targetFile, e) {
        var errFile = (e && e.fileName) ? File(e.fileName).name : (targetFile ? targetFile.name : "unknown");
        var errLine = (e && e.line !== undefined) ? e.line : "unknown";
        var errNumber = (e && e.number !== undefined) ? e.number : "unknown";
        var errMessage = (e && e.message) ? e.message : String(e);

        return LABELS.errorTitle[lang] + "\n\n" +
            LABELS.file[lang] + ": " + errFile + "\n" +
            LABELS.line[lang] + ": " + errLine + "\n" +
            LABELS.errorNumber[lang] + ": " + errNumber + "\n\n" +
            errMessage;
    }

    var f = pickScriptFile(null);

    if (!f) {
        return;
    }

    if (!f.exists) {
        alert(LABELS.fileNotFound[lang] + "\n" + f.fsName);
        return;
    }
    if (!/\.(jsx|jsxbin|js)$/i.test(f.name)) { alert(LABELS.invalidFile[lang] + "\n" + f.name); return; }

    try {
        $.evalFile(f);
    } catch (e) {
        alert(buildErrorText(f, e));
    }
})();
