//==============================================================
// Ruby markup converter
//  [ルビ指定開始マーカー][親字][ルビ開始マーカー][ルビ][ルビ終了マーカー]
//  ex) ｜愛知《あいち》  (group_ruby)
//      〓愛知《あい ち》 (mono_ruby)
//==============================================================

(function () {
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var RUBY_START_MARKERS = ["｜", "〓"]; // 全角ホリゾンタルバー, ゲタ
    var RUBY_OPEN  = "《"; // 全角二重開き山括弧
    var RUBY_CLOSE = "》"; // 全角二重閉じ山括弧
    var doc = app.activeDocument;

    var groupStyle = getOrCreateRubyCharStyle(doc, "group_ruby", RubyTypes.GROUP_RUBY);
    var monoStyle  = getOrCreateRubyCharStyle(doc, "mono_ruby",  RubyTypes.PER_CHARACTER_RUBY);

    app.findGrepPreferences  = NothingEnum.NOTHING;
    app.changeGrepPreferences = NothingEnum.NOTHING;

    var markerClass = "[" + RUBY_START_MARKERS.join("") + "]";

    app.findGrepPreferences.findWhat =
        "(" + markerClass + ")" +
        "([^" + RUBY_OPEN + RUBY_CLOSE + "]+?)" +
        RUBY_OPEN +
        "(.+?)" +
        RUBY_CLOSE;

    var stories = doc.stories;
    var allResults = [];

    for (var i = 0; i < stories.length; i++) {
        var story = stories[i];
        var results = story.findGrep();

        for (var j = 0; j < results.length; j++) {
            allResults.push(results[j]);
        }
    }

    for (var k = allResults.length - 1; k >= 0; k--) {
        processMatch(allResults[k], RUBY_OPEN, RUBY_CLOSE, groupStyle, monoStyle);
    }

    app.findGrepPreferences  = NothingEnum.NOTHING;
    app.changeGrepPreferences = NothingEnum.NOTHING;

    alert("ルビの変換が完了しました。");

    function processMatch(match, rubyOpen, rubyClose, groupStyle, monoStyle) {
        var s = match.contents; // 例: "｜愛知《あい ち》"

        if (!s) {
            return;
        }

        var markerChar   = s.charAt(0);
        var openIndex    = s.indexOf(rubyOpen, 1);
        var closeIndex   = s.lastIndexOf(rubyClose);

        if (openIndex < 0 || closeIndex < 0 || closeIndex <= openIndex) {
            return; // 変な形式なら無視
        }

        var baseText = s.substring(1, openIndex);
        var rubyText = s.substring(openIndex + 1, closeIndex);

        baseText = baseText.replace(/^\s+|\s+$/g, ""); // trim
        rubyText = rubyText.replace(/^\s+|\s+$/g, "");

        if (!baseText || !rubyText) {
            return; // parse error
        }

        var baseChars = match.characters.itemByRange(1, openIndex - 1);

        var isMono = /[ 　]/.test(rubyText);

        if (isMono) {
            applyMonoRuby(baseChars, rubyText, monoStyle);
        } else {
            applyGroupRuby(baseChars, rubyText, groupStyle);
        }

        match.characters.itemByRange(openIndex, s.length - 1).remove();
        match.characters[0].remove();
    }

    function applyGroupRuby(baseChars, rubyText, groupStyle) {
        baseChars.rubyFlag = true;
        baseChars.rubyType = RubyTypes.GROUP_RUBY;
        baseChars.rubyString = rubyText;
        baseChars.appliedCharacterStyle = groupStyle;
    }

    function applyMonoRuby(baseChars, rubyText, monoStyle) {
        var baseLen = baseChars.length;
        var tokens = rubyText.split(/[ 　]+/);

        if (tokens.length !== baseLen) {
            applyGroupRuby(baseChars, rubyText, monoStyle);
            return;
        }

        for (var i = 0; i < baseLen; i++) {
            var ch = baseChars[i];
            ch.rubyFlag = true;
            ch.rubyType = RubyTypes.PER_CHARACTER_RUBY;
            ch.rubyString = tokens[i];
            ch.appliedCharacterStyle = monoStyle;
        }
    }

    function getOrCreateRubyCharStyle(doc, name, rubyTypeEnum) {
        var style;
        try {
            style = doc.characterStyles.itemByName(name);
            style.name;
        } catch (e) {
            style = doc.characterStyles.add({ name: name });
        }

        style.rubyType = rubyTypeEnum;

        return style;
    }

})();
