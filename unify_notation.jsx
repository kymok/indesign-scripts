var REPLACEMENT_TABLE = [
    // 約物
    { src: ["?"], dst: "？" },
    { src: ["!"], dst: "！" },
    { src: [":", "："], dst: "：" },
    { src: [";", "；"], dst: "；" },
    { src: ["･", "・"], dst: "・" },

    // 括弧
    { src: ["("], dst: "（" },
    { src: [")"], dst: "）" },
    { src: ["["], dst: "［" },
    { src: ["]"], dst: "］" },
    { src: ["{"], dst: "｛" },
    { src: ["}"], dst: "｝" },
    { src: ["<", "＜"], dst: "〈" },
    { src: [">", "＞"], dst: "〉" },
    { src: ["«", "<<", "≪"], dst: "《" },
    { src: ["»", ">>", "≫"], dst: "》" },

    // 記号
    { src: ["...", "・・・"], dst: "…" }, // 省略記号
    { src: ["“", "\""], dst: "“" }, // U+201C
    { src: ["”",], dst: "”" }, // U+201D

    // ダッシュ
    {
        src: [
            "－－",
            "--",
            "ーー",
            "−−" // 全角マイナスx2
        ],
        dst: "――" // U+2015 全角ダッシュx2
    },
    {
        src: [
            "~",
            "～"
        ],
        dst: "〜" // U+301C 波ダッシュ
    },

];

(function () {
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;
    var targetInfo = getProcessingTargets(doc);
    if (targetInfo.mode === "invalid-selection") {
        alert("Editable text not selected");
        return;
    }
    var targets = targetInfo.targets;

    // 既存の検索条件を退避
    var savedFindTextPrefs = app.findTextPreferences.properties;
    var savedChangeTextPrefs = app.changeTextPreferences.properties;
    var savedFindGrepPrefs = app.findGrepPreferences.properties;
    var savedChangeGrepPrefs = app.changeGrepPreferences.properties;

    try {
        // 検索条件を初期化
        app.findTextPreferences = NothingEnum.NOTHING;
        app.changeTextPreferences = NothingEnum.NOTHING;
        app.findGrepPreferences = NothingEnum.NOTHING;
        app.changeGrepPreferences = NothingEnum.NOTHING;

        // 行頭の全角スペースを削除
        app.findGrepPreferences.findWhat = "^　(?!　)";
        app.changeGrepPreferences.changeTo = "";
        for (var t = 0; t < targets.length; t++) {
            targets[t].changeGrep();
        }

        // 文字置換
        normalizeByTable(targets);

        alert("表記統一を完了しました。");
    } catch (e) {
        alert("エラーが発生しました: " + e);
    } finally {
        // 検索条件を元に戻す
        app.findTextPreferences.properties = savedFindTextPrefs;
        app.changeTextPreferences.properties = savedChangeTextPrefs;
        app.findGrepPreferences.properties = savedFindGrepPrefs;
        app.changeGrepPreferences.properties = savedChangeGrepPrefs;
    }
})();

function normalizeByTable(targets) {
    for (var i = 0; i < REPLACEMENT_TABLE.length; i++) {
        var rule = REPLACEMENT_TABLE[i];
        if (!rule || !rule.src || !rule.dst) continue;

        for (var j = 0; j < rule.src.length; j++) {
            var s = rule.src[j];
            if (!s || s === rule.dst) continue;

            app.findTextPreferences = NothingEnum.NOTHING;
            app.changeTextPreferences = NothingEnum.NOTHING;

            app.findTextPreferences.findWhat = s;
            app.changeTextPreferences.changeTo = rule.dst;

            for (var t = 0; t < targets.length; t++) {
                targets[t].changeText();
            }
        }
    }
}

function getProcessingTargets(doc) {
    if (!app.selection || app.selection.length === 0) {
        return { mode: "document", targets: [doc] };
    }

    var seen = {};
    var targets = [];

    for (var i = 0; i < app.selection.length; i++) {
        var target = resolveTextTarget(app.selection[i]);
        if (!target) continue;

        var key = target.toSpecifier ? target.toSpecifier() : String(target);
        if (!seen[key]) {
            seen[key] = true;
            targets.push(target);
        }
    }

    if (targets.length === 0) {
        return { mode: "invalid-selection", targets: [] };
    }

    return { mode: "selection", targets: targets };
}

function resolveTextTarget(item) {
    if (!item) return null;

    if (typeof item.changeText === "function" && typeof item.changeGrep === "function") {
        return item;
    }

    try {
        if (item.texts && item.texts.length > 0) {
            return item.texts[0];
        }
    } catch (e) {}

    try {
        if (item.parentStory) {
            return item.parentStory;
        }
    } catch (e) {}

    return null;
}
