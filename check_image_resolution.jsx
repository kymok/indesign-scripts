//==============================================================
// Image resolution checker
// Checks placed raster images against hard-coded effective PPI rules.
//==============================================================

var NEWLINE = "\r";
var PPI_TOLERANCE = 0.5;
var XMP_NS_PHOTOSHOP = "http://ns.adobe.com/photoshop/1.0/";
var XMP_NS_TIFF = "http://ns.adobe.com/tiff/1.0/";
var TIFF_SCAN_BYTES = 262144;

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
    var tiffClassification = getTiffClassification(image);
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

    if (tiffClassification) {
        return tiffClassification;
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
    var linkXmp = getLinkXmp(image);
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

function getTiffClassification(image) {
    var xmpMetadata = getTiffMetadataFromXmp(image);
    var xmpClassification = classifyTiffMetadata(xmpMetadata, "TIFF XMP");

    if (xmpClassification) {
        return xmpClassification;
    }

    var fileMetadata = getTiffMetadataFromFile(image);
    return classifyTiffMetadata(fileMetadata, "TIFFヘッダ");
}

function getTiffMetadataFromXmp(image) {
    var linkXmp = getLinkXmp(image);

    if (!linkXmp || !isValidObject(linkXmp)) {
        return null;
    }

    var samplesPerPixel = getXmpNumericProperty(linkXmp, XMP_NS_TIFF, "tiff:SamplesPerPixel");
    var photometricInterpretation = getXmpNumericProperty(linkXmp, XMP_NS_TIFF, "tiff:PhotometricInterpretation");
    var bitsPerSample = getXmpBitsPerSample(linkXmp);

    if (samplesPerPixel === null && photometricInterpretation === null && !bitsPerSample) {
        return null;
    }

    return {
        samplesPerPixel: samplesPerPixel,
        photometricInterpretation: photometricInterpretation,
        bitsPerSample: bitsPerSample
    };
}

function getTiffMetadataFromFile(image) {
    var file = getLinkedFile(image);

    if (!file || !file.exists || !isTiffFilename(file.name)) {
        return null;
    }

    var originalEncoding = file.encoding;
    var binaryData = null;

    try {
        file.encoding = "BINARY";
        if (!file.open("r")) {
            return null;
        }

        binaryData = file.read(TIFF_SCAN_BYTES);
    } catch (e) {
        return null;
    } finally {
        try {
            file.close();
        } catch (e2) {}

        try {
            file.encoding = originalEncoding;
        } catch (e3) {}
    }

    if (!binaryData || binaryData.length < 8) {
        return null;
    }

    return parseTiffMetadata(binaryStringToBytes(binaryData));
}

function classifyTiffMetadata(metadata, sourceLabel) {
    if (!metadata) {
        return null;
    }

    var samplesPerPixel = asPositiveInteger(metadata.samplesPerPixel);
    var photometric = asPositiveInteger(metadata.photometricInterpretation);
    var bitDepth = getPrimaryBitDepth(metadata.bitsPerSample);
    var suffix = buildTiffMetadataSuffix(samplesPerPixel, photometric, bitDepth);

    if ((samplesPerPixel === 1 || photometric === 0 || photometric === 1) && bitDepth === 1) {
        return buildClassification("monochrome", "モノクロ", RESOLUTION_RULES.monochrome, sourceLabel + suffix);
    }

    if (samplesPerPixel === 1 || photometric === 0 || photometric === 1) {
        return buildClassification("grayscale", "グレー", RESOLUTION_RULES.grayscale, sourceLabel + suffix);
    }

    if (samplesPerPixel >= 3 || photometric === 2 || photometric === 3 || photometric === 5 || photometric === 6 || photometric === 8) {
        return buildClassification("color", "カラー", RESOLUTION_RULES.color, sourceLabel + suffix);
    }

    return null;
}

function buildTiffMetadataSuffix(samplesPerPixel, photometric, bitDepth) {
    var parts = [];

    if (bitDepth !== null) {
        parts.push("BitsPerSample=" + bitDepth);
    }

    if (samplesPerPixel !== null) {
        parts.push("SamplesPerPixel=" + samplesPerPixel);
    }

    if (photometric !== null) {
        parts.push("PhotometricInterpretation=" + photometric);
    }

    if (parts.length === 0) {
        return "";
    }

    return " (" + parts.join(", ") + ")";
}

function getPrimaryBitDepth(bitsPerSample) {
    if (!bitsPerSample || bitsPerSample.length === 0) {
        return null;
    }

    var first = asPositiveInteger(bitsPerSample[0]);

    if (first === null) {
        return null;
    }

    for (var i = 1; i < bitsPerSample.length; i++) {
        if (asPositiveInteger(bitsPerSample[i]) !== first) {
            return first;
        }
    }

    return first;
}

function asPositiveInteger(value) {
    var numeric = Number(value);

    if (isNaN(numeric) || numeric < 0) {
        return null;
    }

    return Math.round(numeric);
}

function getLinkXmp(image) {
    var link = safeRead(image, "itemLink");

    if (!link || !isValidObject(link)) {
        return null;
    }

    var linkXmp = safeRead(link, "linkXmp");
    return linkXmp && isValidObject(linkXmp) ? linkXmp : null;
}

function getXmpNumericProperty(linkXmp, namespaceUri, propertyPath) {
    if (!linkXmp) {
        return null;
    }

    try {
        var value = linkXmp.getProperty(namespaceUri, propertyPath);

        if (value === null || typeof value === "undefined" || value === "") {
            return null;
        }

        var numeric = Number(value);
        return isNaN(numeric) ? null : numeric;
    } catch (e) {
        return null;
    }
}

function getXmpBitsPerSample(linkXmp) {
    var directValue = getXmpNumericProperty(linkXmp, XMP_NS_TIFF, "tiff:BitsPerSample");
    if (directValue !== null) {
        return [directValue];
    }

    if (!hasMethod(linkXmp, "countContainer")) {
        return null;
    }

    var count = 0;

    try {
        count = Number(linkXmp.countContainer(XMP_NS_TIFF, "tiff:BitsPerSample"));
    } catch (e) {
        count = 0;
    }

    if (!count || isNaN(count)) {
        return null;
    }

    var bits = [];

    for (var i = 1; i <= count; i++) {
        var value = getXmpNumericProperty(linkXmp, XMP_NS_TIFF, "tiff:BitsPerSample[" + i + "]");
        if (value === null) {
            return null;
        }
        bits.push(value);
    }

    return bits.length > 0 ? bits : null;
}

function getLinkedFile(image) {
    var link = safeRead(image, "itemLink");
    if (!link || !isValidObject(link)) {
        return null;
    }

    var filePath = safeRead(link, "filePath");
    if (!filePath) {
        return null;
    }

    try {
        return File(filePath);
    } catch (e) {
        return null;
    }
}

function isTiffFilename(name) {
    if (!name) {
        return false;
    }

    return /\.(tif|tiff)$/i.test(String(name));
}

function parseTiffMetadata(bytes) {
    if (!bytes || bytes.length < 8) {
        return null;
    }

    var littleEndian;

    if (bytes[0] === 0x49 && bytes[1] === 0x49) {
        littleEndian = true;
    } else if (bytes[0] === 0x4D && bytes[1] === 0x4D) {
        littleEndian = false;
    } else {
        return null;
    }

    var magic = readUint16(bytes, 2, littleEndian);
    if (magic !== 42) {
        return null;
    }

    var ifdOffset = readUint32(bytes, 4, littleEndian);
    if (!isByteRangeAvailable(bytes, ifdOffset, 2)) {
        return null;
    }

    var entryCount = readUint16(bytes, ifdOffset, littleEndian);
    var metadata = {
        bitsPerSample: null,
        samplesPerPixel: null,
        photometricInterpretation: null
    };

    for (var i = 0; i < entryCount; i++) {
        var entryOffset = ifdOffset + 2 + (i * 12);
        if (!isByteRangeAvailable(bytes, entryOffset, 12)) {
            return null;
        }

        var tag = readUint16(bytes, entryOffset, littleEndian);
        var values = readTiffEntryValues(bytes, entryOffset, littleEndian);

        if (!values) {
            continue;
        }

        if (tag === 258) {
            metadata.bitsPerSample = values;
        } else if (tag === 262 && values.length > 0) {
            metadata.photometricInterpretation = values[0];
        } else if (tag === 277 && values.length > 0) {
            metadata.samplesPerPixel = values[0];
        }
    }

    if (!metadata.bitsPerSample && metadata.samplesPerPixel === null && metadata.photometricInterpretation === null) {
        return null;
    }

    return metadata;
}

function readTiffEntryValues(bytes, entryOffset, littleEndian) {
    var fieldType = readUint16(bytes, entryOffset + 2, littleEndian);
    var count = readUint32(bytes, entryOffset + 4, littleEndian);
    var typeSize = getTiffFieldTypeSize(fieldType);

    if (!typeSize || !count) {
        return null;
    }

    var totalSize = typeSize * count;
    var valueOffset = totalSize <= 4 ? entryOffset + 8 : readUint32(bytes, entryOffset + 8, littleEndian);

    if (!isByteRangeAvailable(bytes, valueOffset, totalSize)) {
        return null;
    }

    var values = [];

    for (var i = 0; i < count; i++) {
        var offset = valueOffset + (i * typeSize);

        if (fieldType === 1 || fieldType === 6 || fieldType === 7) {
            values.push(bytes[offset]);
        } else if (fieldType === 3) {
            values.push(readUint16(bytes, offset, littleEndian));
        } else if (fieldType === 4) {
            values.push(readUint32(bytes, offset, littleEndian));
        } else {
            return null;
        }
    }

    return values;
}

function getTiffFieldTypeSize(fieldType) {
    if (fieldType === 1 || fieldType === 2 || fieldType === 6 || fieldType === 7) {
        return 1;
    }

    if (fieldType === 3 || fieldType === 8) {
        return 2;
    }

    if (fieldType === 4 || fieldType === 9 || fieldType === 11) {
        return 4;
    }

    if (fieldType === 5 || fieldType === 10 || fieldType === 12) {
        return 8;
    }

    return 0;
}

function readUint16(bytes, offset, littleEndian) {
    if (!isByteRangeAvailable(bytes, offset, 2)) {
        return 0;
    }

    if (littleEndian) {
        return bytes[offset] + (bytes[offset + 1] * 256);
    }

    return (bytes[offset] * 256) + bytes[offset + 1];
}

function readUint32(bytes, offset, littleEndian) {
    if (!isByteRangeAvailable(bytes, offset, 4)) {
        return 0;
    }

    if (littleEndian) {
        return bytes[offset] +
            (bytes[offset + 1] * 256) +
            (bytes[offset + 2] * 65536) +
            (bytes[offset + 3] * 16777216);
    }

    return (bytes[offset] * 16777216) +
        (bytes[offset + 1] * 65536) +
        (bytes[offset + 2] * 256) +
        bytes[offset + 3];
}

function isByteRangeAvailable(bytes, offset, length) {
    return offset >= 0 && length >= 0 && (offset + length) <= bytes.length;
}

function binaryStringToBytes(binaryString) {
    var bytes = [];

    for (var i = 0; i < binaryString.length; i++) {
        bytes.push(binaryString.charCodeAt(i) & 0xFF);
    }

    return bytes;
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

function hasMethod(target, methodName) {
    var reflection = getReflection(target);

    if (!reflection) {
        return false;
    }

    try {
        return reflection.find(methodName) !== null;
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
