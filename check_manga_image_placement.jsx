//==============================================================
// Manga image placement checker
// Checks placed graphics against bleed and centering rules.
//==============================================================

var IGNORE_SMALLER_THAN_PAGE_MM = 1;
var CHECK_TOLERANCE_MM = 0.1;
var FLOAT_EPSILON_MM = 0.01;
var NEWLINE = "\r";

(function () {
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;
    var measurementState = saveMeasurementState();

    try {
        setMeasurementStateToMillimeters();
        var result = inspectDocument(doc);
        showReportDialog(result);
    } catch (e) {
        alert("エラーが発生しました: " + e);
    } finally {
        restoreMeasurementState(measurementState);
    }
})();

function inspectDocument(doc) {
    var result = {
        totalGraphics: 0,
        checkedGraphics: 0,
        ignoredGraphics: 0,
        okGraphics: 0,
        issueGraphics: 0,
        skippedGraphics: 0,
        issueFrames: [],
        issueItems: [],
        issueEntries: [],
        skippedEntries: []
    };

    var bleed = getDocumentBleed(doc);
    var graphics = doc.allGraphics;

    for (var g = 0; g < graphics.length; g++) {
        var graphic = graphics[g];
        result.totalGraphics++;

        var frame = getGraphicFrame(graphic);
        var page = getGraphicPage(graphic, frame);
        var pageBounds = getBounds(page);
        var frameBounds = getBounds(frame);
        var graphicBounds = getBounds(graphic);

        if (!page) {
            result.skippedGraphics++;
            result.skippedEntries.push(buildSkippedEntry(page, graphic, "ページに紐づかない画像のため判定できませんでした。"));
            continue;
        }

        if (!frame || !pageBounds || !frameBounds || !graphicBounds) {
            result.skippedGraphics++;
            result.skippedEntries.push(buildSkippedEntry(page, graphic, "境界を取得できませんでした。"));
            continue;
        }

        if (!isAxisAligned(frame) || !isAxisAligned(graphic)) {
            result.skippedGraphics++;
            result.skippedEntries.push(buildSkippedEntry(page, graphic, "回転またはシアーされたオブジェクトは自動判定の対象外です。"));
            continue;
        }

        if (isMuchSmallerThanPage(graphicBounds, pageBounds)) {
            result.ignoredGraphics++;
            continue;
        }

        result.checkedGraphics++;

        var issues = [];
        var actual = getFrameProtrusions(frameBounds, pageBounds, page);
        var expected = getExpectedProtrusions(page, bleed);

        pushProtrusionIssue(issues, "天", actual.top, expected.top);
        pushProtrusionIssue(issues, "地", actual.bottom, expected.bottom);
        pushProtrusionIssue(issues, "小口", actual.outer, expected.outer);
        pushProtrusionIssue(issues, "ノド", actual.spine, expected.spine);
        pushCenterIssue(issues, graphicBounds, pageBounds);

        if (issues.length > 0) {
            result.issueGraphics++;
            result.issueFrames.push(frame);
            result.issueItems.push({
                page: page,
                frame: frame,
                graphic: graphic
            });
            result.issueEntries.push(buildIssueEntry(page, graphic, issues));
        } else {
            result.okGraphics++;
        }
    }

    return result;
}

function getDocumentBleed(doc) {
    var prefs = doc.documentPreferences;
    var top = prefs.documentBleedTopOffset;

    if (prefs.documentBleedUniformSize) {
        return {
            top: top,
            bottom: top,
            inside: top,
            outside: top
        };
    }

    return {
        top: top,
        bottom: prefs.documentBleedBottomOffset,
        inside: prefs.documentBleedInsideOrLeftOffset,
        outside: prefs.documentBleedOutsideOrRightOffset
    };
}

function saveMeasurementState() {
    return {
        measurementUnit: app.scriptPreferences.measurementUnit
    };
}

function setMeasurementStateToMillimeters() {
    app.scriptPreferences.measurementUnit = MeasurementUnits.MILLIMETERS;
}

function restoreMeasurementState(state) {
    if (!state) {
        return;
    }

    app.scriptPreferences.measurementUnit = state.measurementUnit;
}

function getGraphicFrame(graphic) {
    if (!graphic || !graphic.isValid) {
        return null;
    }

    var current = graphic;

    for (var i = 0; i < 6; i++) {
        try {
            current = current.parent;
        } catch (e) {
            return null;
        }

        if (!current || !current.isValid) {
            return null;
        }

        try {
            var bounds = current.geometricBounds;
            if (bounds && bounds.length === 4) {
                return current;
            }
        } catch (e2) {}
    }

    return null;
}

function getBounds(item) {
    if (!item) {
        return null;
    }

    try {
        if (item.isValid === false) {
            return null;
        }
    } catch (e) {}

    var rawBounds = null;

    try {
        rawBounds = item.geometricBounds;
    } catch (e2) {}

    if (!rawBounds) {
        try {
            rawBounds = item.bounds;
        } catch (e3) {}
    }

    if (!rawBounds || rawBounds.length !== 4) {
        return null;
    }

    var top = Number(rawBounds[0]);
    var left = Number(rawBounds[1]);
    var bottom = Number(rawBounds[2]);
    var right = Number(rawBounds[3]);

    if (isNaN(top) || isNaN(left) || isNaN(bottom) || isNaN(right)) {
        return null;
    }

    return {
        top: top,
        left: left,
        bottom: bottom,
        right: right,
        width: right - left,
        height: bottom - top,
        centerX: (left + right) / 2,
        centerY: (top + bottom) / 2
    };
}

function getGraphicPage(graphic, frame) {
    try {
        if (graphic.parentPage && graphic.parentPage.isValid) {
            return graphic.parentPage;
        }
    } catch (e) {}

    try {
        if (frame && frame.parentPage && frame.parentPage.isValid) {
            return frame.parentPage;
        }
    } catch (e2) {}

    return null;
}

function isAxisAligned(item) {
    var angleTolerance = 0.01;

    try {
        var rotation = normalizeAngle(item.absoluteRotationAngle);
        if (rotation > angleTolerance && rotation < 180 - angleTolerance) {
            return false;
        }
    } catch (e) {}

    try {
        var shear = Math.abs(Number(item.absoluteShearAngle));
        if (!isNaN(shear) && shear > angleTolerance) {
            return false;
        }
    } catch (e2) {}

    return true;
}

function normalizeAngle(angle) {
    var normalized = Number(angle) % 180;

    if (normalized < 0) {
        normalized += 180;
    }

    return Math.abs(normalized);
}

function isMuchSmallerThanPage(graphicBounds, pageBounds) {
    var threshold = IGNORE_SMALLER_THAN_PAGE_MM - FLOAT_EPSILON_MM;
    var widthDiff = pageBounds.width - graphicBounds.width;
    var heightDiff = pageBounds.height - graphicBounds.height;

    return widthDiff >= threshold && heightDiff >= threshold;
}

function getFrameProtrusions(frameBounds, pageBounds, page) {
    var left = pageBounds.left - frameBounds.left;
    var right = frameBounds.right - pageBounds.right;
    var side = getPageSideName(page);

    return {
        top: pageBounds.top - frameBounds.top,
        bottom: frameBounds.bottom - pageBounds.bottom,
        outer: side === "left" ? left : right,
        spine: side === "left" ? right : left
    };
}

function getExpectedFrameBounds(pageBounds, page, bleed) {
    var side = getPageSideName(page);
    var spine = pageHasOppositePage(page) ? 0 : bleed.inside;
    var left;
    var right;

    if (side === "left") {
        left = pageBounds.left - bleed.outside;
        right = pageBounds.right + spine;
    } else {
        left = pageBounds.left - spine;
        right = pageBounds.right + bleed.outside;
    }

    return {
        top: pageBounds.top - bleed.top,
        left: left,
        bottom: pageBounds.bottom + bleed.bottom,
        right: right
    };
}

function getExpectedProtrusions(page, bleed) {
    return {
        top: bleed.top,
        bottom: bleed.bottom,
        outer: bleed.outside,
        spine: pageHasOppositePage(page) ? 0 : bleed.inside
    };
}

function getPageSideName(page) {
    try {
        if (page.side === PageSideOptions.LEFT_HAND) {
            return "left";
        }
        if (page.side === PageSideOptions.RIGHT_HAND) {
            return "right";
        }
        if (page.side === PageSideOptions.SINGLE_SIDED) {
            return "single";
        }
    } catch (e) {}

    return "single";
}

function pageHasOppositePage(page) {
    var spread;

    try {
        spread = page.parent;
    } catch (e) {
        return false;
    }

    if (!spread || !spread.isValid) {
        return false;
    }

    var myId = page.id;
    var mySide = getPageSideName(page);
    var pages = spread.pages;

    for (var i = 0; i < pages.length; i++) {
        var otherPage = pages[i];

        if (!otherPage || !otherPage.isValid || otherPage.id === myId) {
            continue;
        }

        if (getPageSideName(otherPage) !== mySide) {
            return true;
        }
    }

    return false;
}

function pushProtrusionIssue(issues, label, actual, expected) {
    if (!nearlyEqual(actual, expected, CHECK_TOLERANCE_MM)) {
        issues.push(label + ": 実測 " + formatMm(actual) + " / 期待 " + formatMm(expected));
    }
}

function pushCenterIssue(issues, graphicBounds, pageBounds) {
    var tolerance = CHECK_TOLERANCE_MM;
    var deltaX = graphicBounds.centerX - pageBounds.centerX;
    var deltaY = graphicBounds.centerY - pageBounds.centerY;

    if (Math.abs(deltaX) > tolerance || Math.abs(deltaY) > tolerance) {
        issues.push("画像中心: X " + formatSignedMm(deltaX) + ", Y " + formatSignedMm(deltaY));
    }
}

function nearlyEqual(a, b, tolerance) {
    return Math.abs(a - b) <= tolerance;
}

function buildIssueEntry(page, graphic, issues) {
    var lines = [describeGraphic(page, graphic)];

    for (var i = 0; i < issues.length; i++) {
        lines.push("  - " + issues[i]);
    }

    return lines.join(NEWLINE);
}

function buildSkippedEntry(page, graphic, reason) {
    return describeGraphic(page, graphic) + NEWLINE + "  - " + reason;
}

function describeGraphic(page, graphic) {
    var pageName = page && page.isValid ? page.name : "ページ不明";
    return "ページ " + pageName + " / " + getGraphicName(graphic);
}

function getGraphicName(graphic) {
    try {
        if (graphic.itemLink && graphic.itemLink.isValid && graphic.itemLink.name) {
            return graphic.itemLink.name;
        }
    } catch (e) {}

    try {
        if (graphic.name) {
            return graphic.name;
        }
    } catch (e2) {}

    try {
        return "Graphic #" + graphic.id;
    } catch (e3) {}

    return "名称不明";
}

function buildReportText(result) {
    var lines = [];

    lines.push("漫画画像貼り込みチェック");
    lines.push("");
    lines.push("総グラフィック数: " + result.totalGraphics);
    lines.push("チェック対象: " + result.checkedGraphics);
    lines.push("問題なし: " + result.okGraphics);
    lines.push("要確認: " + result.issueGraphics);
    lines.push("除外: " + result.ignoredGraphics + "（ページより幅・高さとも1mm以上小さい）");

    if (result.skippedGraphics > 0) {
        lines.push("未判定: " + result.skippedGraphics);
    }

    lines.push("判定許容差: " + CHECK_TOLERANCE_MM + "mm");

    if (result.issueEntries.length === 0 && result.skippedEntries.length === 0) {
        lines.push("");
        lines.push("要確認項目はありませんでした。");
        return lines.join(NEWLINE);
    }

    if (result.issueEntries.length > 0) {
        lines.push("");
        lines.push("要確認項目");
        lines.push("");

        for (var i = 0; i < result.issueEntries.length; i++) {
            lines.push(result.issueEntries[i]);
            if (i < result.issueEntries.length - 1) {
                lines.push("");
            }
        }
    }

    if (result.skippedEntries.length > 0) {
        lines.push("");
        lines.push("未判定項目");
        lines.push("");

        for (var j = 0; j < result.skippedEntries.length; j++) {
            lines.push(result.skippedEntries[j]);
            if (j < result.skippedEntries.length - 1) {
                lines.push("");
            }
        }
    }

    return lines.join(NEWLINE);
}

function showReportDialog(result, notice) {
    var reportText = buildReportText(result);
    if (notice) {
        reportText = notice + NEWLINE + NEWLINE + reportText;
    }

    var win = new Window("dialog", "漫画画像貼り込みチェック");
    var shouldSelectIssues = false;
    var shouldAutoFix = false;

    win.orientation = "column";
    win.alignChildren = ["fill", "fill"];

    var reportField = win.add("edittext", undefined, reportText, {
        multiline: true,
        readonly: true,
        scrolling: true
    });
    reportField.preferredSize = [760, 520];
    reportField.active = true;

    var buttonGroup = win.add("group");
    buttonGroup.alignment = ["right", "center"];

    if (result.issueFrames.length > 0) {
        var fixButton = buttonGroup.add("button", undefined, "問題を自動修正");
        fixButton.onClick = function () {
            shouldAutoFix = true;
            win.close(1);
        };

        var selectButton = buttonGroup.add("button", undefined, "問題のフレームを選択");
        selectButton.onClick = function () {
            shouldSelectIssues = true;
            win.close(1);
        };
    }

    buttonGroup.add("button", undefined, "閉じる", { name: "ok" });

    win.show();

    if (shouldSelectIssues) {
        selectItems(result.issueFrames);
        return;
    }

    if (shouldAutoFix) {
        var fixResult = autoFixIssues(app.activeDocument, result.issueItems);
        var rerunResult = inspectDocument(app.activeDocument);
        showReportDialog(rerunResult, buildFixNotice(fixResult));
    }
}

function autoFixIssues(doc, issueItems) {
    var result = {
        attempted: issueItems.length,
        fixed: 0,
        failed: 0,
        failureEntries: []
    };
    var bleed = getDocumentBleed(doc);

    for (var i = 0; i < issueItems.length; i++) {
        var item = issueItems[i];

        try {
            if (!item || !item.page || !item.frame || !item.graphic) {
                throw new Error("必要なオブジェクト参照がありません。");
            }
            if (!item.page.isValid || !item.frame.isValid || !item.graphic.isValid) {
                throw new Error("オブジェクトが無効です。");
            }

            fixGraphicPlacement(item.page, item.frame, item.graphic, bleed);
            result.fixed++;
        } catch (e) {
            result.failed++;
            result.failureEntries.push(describeGraphic(item.page, item.graphic) + ": " + e);
        }
    }

    return result;
}

function fixGraphicPlacement(page, frame, graphic, bleed) {
    var pageBounds = getBounds(page);

    if (!pageBounds) {
        throw new Error("境界を取得できませんでした。");
    }

    var targetFrameBounds = getExpectedFrameBounds(pageBounds, page, bleed);
    setItemBounds(frame, targetFrameBounds);

    var graphicBounds = getBounds(graphic);

    if (!graphicBounds) {
        throw new Error("画像の境界を取得できませんでした。");
    }

    var deltaX = pageBounds.centerX - graphicBounds.centerX;
    var deltaY = pageBounds.centerY - graphicBounds.centerY;

    if (Math.abs(deltaX) > 0 || Math.abs(deltaY) > 0) {
        offsetItemBounds(graphic, deltaX, deltaY);
    }
}

function setItemBounds(item, bounds) {
    item.geometricBounds = [bounds.top, bounds.left, bounds.bottom, bounds.right];
}

function offsetItemBounds(item, deltaX, deltaY) {
    var bounds = getBounds(item);

    if (!bounds) {
        throw new Error("移動対象の境界を取得できませんでした。");
    }

    item.geometricBounds = [
        bounds.top + deltaY,
        bounds.left + deltaX,
        bounds.bottom + deltaY,
        bounds.right + deltaX
    ];
}

function buildFixNotice(fixResult) {
    var lines = [];

    lines.push("自動修正結果: " + fixResult.fixed + "件修正 / " + fixResult.failed + "件失敗");

    if (fixResult.failureEntries.length > 0) {
        lines.push("");
        lines.push("修正できなかった項目");
        lines.push("");

        for (var i = 0; i < fixResult.failureEntries.length; i++) {
            lines.push(fixResult.failureEntries[i]);
        }
    }

    return lines.join(NEWLINE);
}

function selectItems(items) {
    var validItems = [];

    for (var i = 0; i < items.length; i++) {
        if (items[i] && items[i].isValid) {
            validItems.push(items[i]);
        }
    }

    if (validItems.length === 0) {
        return;
    }

    app.select(validItems[0], SelectionOptions.REPLACE_WITH);

    for (var j = 1; j < validItems.length; j++) {
        app.select(validItems[j], SelectionOptions.ADD_TO);
    }
}

function formatMm(valueInMillimeters) {
    return Number(valueInMillimeters).toFixed(2) + "mm";
}

function formatSignedMm(valueInMillimeters) {
    var value = Number(valueInMillimeters);
    var sign = value > 0 ? "+" : "";
    return sign + value.toFixed(2) + "mm";
}
