//==============================================================
// Image resolution checker
// Checks placed raster images against hard-coded effective PPI rules.
//==============================================================

var NEWLINE = "\r";
var PPI_TOLERANCE = 0.5;
var XMP_NS_PHOTOSHOP = "http://ns.adobe.com/photoshop/1.0/";

var RESOLUTION_RULES = {
    monochrome: [600, 1200],
    grayscale: [600, 1200],
    color: [350, 600, 1200]
};

(function () {
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;

    try {
        var result = inspectDocument(doc);
        showReportDialog(result);
    } catch (e) {
        alert("エラーが発生しました: " + e);
    }
})();

function inspectDocument(doc) {
    var result = {
        totalImages: 0,
        checkedImages: 0,
        okImages: 0,
        issueImages: 0,
        skippedImages: 0,
        issueFrames: [],
        skippedFrames: [],
        issueEntries: [],
        skippedEntries: [],
        typeCounts: {
            monochrome: 0,
            grayscale: 0,
            color: 0,
            unknown: 0
        }
    };

    var images = getPlacedRasterImages(doc);

    for (var i = 0; i < images.length; i++) {
        var image = images[i];
        result.totalImages++;

        var frame = getGraphicFrame(image);
        var page = getGraphicPage(image, frame);
        var classification = classifyImage(image);
        var effectivePpi = getPpiPair(safeRead(image, "effectivePpi"));
        var actualPpi = getPpiPair(safeRead(image, "actualPpi"));

        incrementTypeCount(result.typeCounts, classification.kind);

        if (!classification.allowedValues) {
            result.skippedImages++;
            if (frame) {
                result.skippedFrames.push(frame);
            }
            result.skippedEntries.push(buildSkippedEntry(page, image, classification, effectivePpi, "画像種別を判定できませんでした。"));
            continue;
        }

        if (!effectivePpi) {
            result.skippedImages++;
            if (frame) {
                result.skippedFrames.push(frame);
            }
            result.skippedEntries.push(buildSkippedEntry(page, image, classification, effectivePpi, "実効解像度を取得できませんでした。"));
            continue;
        }

        result.checkedImages++;

        if (isAllowedPpiPair(effectivePpi, classification.allowedValues)) {
            result.okImages++;
            continue;
        }

        result.issueImages++;
        if (frame) {
            result.issueFrames.push(frame);
        }
        result.issueEntries.push(buildIssueEntry(page, image, classification, effectivePpi, actualPpi));
    }

    return result;
}

function classifyImage(image) {
    var colorMode = getPhotoshopColorMode(image);
    var imageTypeName = normalizeImageTypeName(safeRead(image, "imageTypeName"));

    if (colorMode === 0) {
        return buildClassification("monochrome", "モノクロ", RESOLUTION_RULES.monochrome, "XMP: Bitmap");
    }

    if (colorMode === 1) {
        return buildClassification("grayscale", "グレー", RESOLUTION_RULES.grayscale, "XMP: Gray scale");
    }

    if (colorMode === 3 || colorMode === 4 || colorMode === 9) {
        return buildClassification("color", "カラー", RESOLUTION_RULES.color, "XMP: Color");
    }

    if (colorMode === 2 || colorMode === 7 || colorMode === 8) {
        return buildClassification("unknown", "未判定", null, "XMP: 特殊カラーモード");
    }

    if (imageTypeName === "bitmap" || imageTypeName === "bilevel" || imageTypeName === "1bit") {
        return buildClassification("monochrome", "モノクロ", RESOLUTION_RULES.monochrome, "imageTypeName: " + imageTypeName);
    }

    return buildClassification("unknown", "未判定", null, "XMP の ColorMode を取得できませんでした");
}

function getPlacedRasterImages(doc) {
    var allGraphics = safeRead(doc, "allGraphics");
    if (!allGraphics) {
        return [];
    }

    var result = [];

    for (var i = 0; i < allGraphics.length; i++) {
        var graphic = allGraphics[i];

        if (isResolutionCheckTarget(graphic)) {
            result.push(graphic);
        }
    }

    return result;
}

function isResolutionCheckTarget(graphic) {
    if (!graphic || !isValidObject(graphic)) {
        return false;
    }

    if (getPpiPair(safeRead(graphic, "effectivePpi"))) {
        return true;
    }

    if (getPpiPair(safeRead(graphic, "actualPpi"))) {
        return true;
    }

    return false;
}

function buildClassification(kind, label, allowedValues, sourceLabel) {
    return {
        kind: kind,
        label: label,
        allowedValues: allowedValues,
        sourceLabel: sourceLabel
    };
}

function getPhotoshopColorMode(image) {
    var link = safeRead(image, "itemLink");

    if (!link || !isValidObject(link)) {
        return null;
    }

    var linkXmp = safeRead(link, "linkXmp");

    if (!linkXmp || !isValidObject(linkXmp)) {
        return null;
    }

    try {
        var value = linkXmp.getProperty(XMP_NS_PHOTOSHOP, "photoshop:ColorMode");
        if (value === null || typeof value === "undefined" || value === "") {
            return null;
        }

        var numeric = Number(value);
        return isNaN(numeric) ? null : numeric;
    } catch (e) {
        return null;
    }
}

function normalizeImageTypeName(value) {
    if (value === null || typeof value === "undefined") {
        return null;
    }

    var text = String(value).toLowerCase();

    if (text.indexOf("bitmap") >= 0) {
        return "bitmap";
    }
    if (text.indexOf("bilevel") >= 0) {
        return "bilevel";
    }
    if (text.indexOf("1-bit") >= 0 || text.indexOf("1bit") >= 0) {
        return "1bit";
    }

    return text;
}

function getPpiPair(value) {
    if (value === null || typeof value === "undefined") {
        return null;
    }

    var x;
    var y;

    try {
        if (typeof value.length !== "undefined") {
            x = Number(value[0]);
            y = Number(value.length > 1 ? value[1] : value[0]);
        } else {
            x = Number(value);
            y = Number(value);
        }
    } catch (e) {
        x = Number(value);
        y = Number(value);
    }

    if (isNaN(x) || isNaN(y)) {
        return null;
    }

    return {
        x: x,
        y: y
    };
}

function isAllowedPpiPair(ppiPair, allowedValues) {
    return isAllowedPpiValue(ppiPair.x, allowedValues) && isAllowedPpiValue(ppiPair.y, allowedValues);
}

function isAllowedPpiValue(value, allowedValues) {
    for (var i = 0; i < allowedValues.length; i++) {
        if (Math.abs(value - allowedValues[i]) <= PPI_TOLERANCE) {
            return true;
        }
    }

    return false;
}

function incrementTypeCount(typeCounts, kind) {
    if (typeCounts.hasOwnProperty(kind)) {
        typeCounts[kind]++;
    } else {
        typeCounts.unknown++;
    }
}

function buildIssueEntry(page, image, classification, effectivePpi, actualPpi) {
    var lines = [describeGraphic(page, image)];

    lines.push("  - 種別: " + classification.label + "（" + classification.sourceLabel + "）");
    lines.push("  - 実効解像度: " + formatPpiPair(effectivePpi));

    if (actualPpi) {
        lines.push("  - 元画像解像度: " + formatPpiPair(actualPpi));
    }

    lines.push("  - 許容値: " + formatAllowedValues(classification.allowedValues));

    return lines.join(NEWLINE);
}

function buildSkippedEntry(page, image, classification, effectivePpi, reason) {
    var lines = [describeGraphic(page, image)];

    lines.push("  - 種別: " + classification.label + "（" + classification.sourceLabel + "）");

    if (effectivePpi) {
        lines.push("  - 実効解像度: " + formatPpiPair(effectivePpi));
    }

    lines.push("  - " + reason);

    return lines.join(NEWLINE);
}

function describeGraphic(page, image) {
    var pageName = page && isValidObject(page) ? safeRead(page, "name") : null;
    if (!pageName) {
        pageName = "ページ不明";
    }
    return "ページ " + pageName + " / " + getGraphicName(image);
}

function getGraphicName(image) {
    var link = safeRead(image, "itemLink");

    try {
        if (link && link.isValid && link.name) {
            return link.name;
        }
    } catch (e) {}

    try {
        if (image.name) {
            return image.name;
        }
    } catch (e2) {}

    try {
        return "Image #" + image.id;
    } catch (e3) {}

    return "名称不明";
}

function buildReportText(result) {
    var lines = [];

    lines.push("画像解像度チェック");
    lines.push("");
    lines.push("対象: 配置されたラスター画像");
    lines.push("総画像数: " + result.totalImages);
    lines.push("チェック対象: " + result.checkedImages);
    lines.push("問題なし: " + result.okImages);
    lines.push("要確認: " + result.issueImages);

    if (result.skippedImages > 0) {
        lines.push("未判定: " + result.skippedImages);
    }

    lines.push("");
    lines.push("許容値");
    lines.push("  モノクロ: " + formatAllowedValues(RESOLUTION_RULES.monochrome));
    lines.push("  グレー: " + formatAllowedValues(RESOLUTION_RULES.grayscale));
    lines.push("  カラー: " + formatAllowedValues(RESOLUTION_RULES.color));
    lines.push("判定許容差: ±" + PPI_TOLERANCE + " ppi");

    lines.push("");
    lines.push("種別内訳");
    lines.push("  モノクロ: " + result.typeCounts.monochrome);
    lines.push("  グレー: " + result.typeCounts.grayscale);
    lines.push("  カラー: " + result.typeCounts.color);

    if (result.typeCounts.unknown > 0) {
        lines.push("  未判定: " + result.typeCounts.unknown);
    }

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

function showReportDialog(result) {
    var reportText = buildReportText(result);
    var win = new Window("dialog", "画像解像度チェック");
    var shouldSelectFlagged = false;

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

    if (result.issueFrames.length > 0 || result.skippedFrames.length > 0) {
        var selectButton = buttonGroup.add("button", undefined, "要確認フレームを選択");
        selectButton.onClick = function () {
            shouldSelectFlagged = true;
            win.close(1);
        };
    }

    buttonGroup.add("button", undefined, "閉じる", { name: "ok" });

    win.show();

    if (shouldSelectFlagged) {
        selectItems(uniqueValidItems(result.issueFrames.concat(result.skippedFrames)));
    }
}

function uniqueValidItems(items) {
    var result = [];
    var seen = {};

    for (var i = 0; i < items.length; i++) {
        var item = items[i];

        if (!item || !isValidObject(item)) {
            continue;
        }

        var key = getObjectKey(item);

        if (key === null || seen[key]) {
            continue;
        }

        seen[key] = true;
        result.push(item);
    }

    return result;
}

function getObjectKey(item) {
    try {
        return String(item.id);
    } catch (e) {
        return null;
    }
}

function selectItems(items) {
    if (items.length === 0) {
        return;
    }

    app.select(items[0], SelectionOptions.REPLACE_WITH);

    for (var i = 1; i < items.length; i++) {
        app.select(items[i], SelectionOptions.ADD_TO);
    }
}

function getGraphicFrame(graphic) {
    if (!graphic || !isValidObject(graphic)) {
        return null;
    }

    var current = graphic;

    for (var i = 0; i < 6; i++) {
        try {
            current = current.parent;
        } catch (e) {
            return null;
        }

        if (!current || !isValidObject(current)) {
            return null;
        }

        var bounds = safeRead(current, "geometricBounds");
        if (bounds && bounds.length === 4) {
            return current;
        }
    }

    return null;
}

function getGraphicPage(graphic, frame) {
    var graphicPage = safeRead(graphic, "parentPage");
    if (graphicPage && isValidObject(graphicPage)) {
        return graphicPage;
    }

    var framePage = safeRead(frame, "parentPage");
    if (framePage && isValidObject(framePage)) {
        return framePage;
    }

    return null;
}

function formatAllowedValues(values) {
    var parts = [];

    for (var i = 0; i < values.length; i++) {
        parts.push(String(values[i]));
    }

    return parts.join(" / ") + " ppi";
}

function formatPpiPair(ppiPair) {
    return formatPpiValue(ppiPair.x) + " x " + formatPpiValue(ppiPair.y) + " ppi";
}

function formatPpiValue(value) {
    return Number(value).toFixed(2);
}

function safeRead(target, propertyName) {
    if (!target || !hasProperty(target, propertyName)) {
        return null;
    }

    try {
        return target[propertyName];
    } catch (e) {
        return null;
    }
}

function hasProperty(target, propertyName) {
    var reflection = getReflection(target);

    if (!reflection) {
        return false;
    }

    try {
        return reflection.find(propertyName) !== null;
    } catch (e) {
        return false;
    }
}

function getReflection(target) {
    if (!target) {
        return null;
    }

    try {
        return target.reflect;
    } catch (e) {
        return null;
    }
}

function isValidObject(target) {
    try {
        return target && target.isValid !== false;
    } catch (e) {
        return !!target;
    }
}
