var f = File.openDialog("実行するExtendScriptを選択", "*.jsx;*.js;*.jsxbin");
if (f) {
    $.evalFile(f);
}
