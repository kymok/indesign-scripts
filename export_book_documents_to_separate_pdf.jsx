/*
  Export each document in the active InDesign Book as a separate PDF.

  Usage:
  1. Save this file as: ExportBookDocsToSeparatePDF.jsx
  2. Put it in InDesign Scripts Panel folder
  3. Open the target book (.indb)
  4. Run the script from Scripts panel

  Notes:
  - PDF_PRESET_NAME, FILENAME_PREFIX, and output folder are dialog defaults.
  - If PDF_PRESET_NAME is empty, the dialog defaults to current PDF export settings.
  - Default output folder: next to the .indb file, named "<BookName>_PDF"
*/

(function () {
    // =========================
    // Configuration
    // =========================

    // Example:
    // var PDF_PRESET_NAME = "[High Quality Print]";
    // var PDF_PRESET_NAME = "[Press Quality]";
    // In localized environments, preset names may differ.
    var PDF_PRESET_NAME = "栄光_漫画あり";

    // Prefix added to each exported PDF filename. Example: "2026_"
    var FILENAME_PREFIX = "rev1_";

    // If true, try to update book numbering before export.
    var UPDATE_BOOK_NUMBERING = false;

    // If true, overwrite existing PDFs with the same name.
    var OVERWRITE_EXISTING = true;

    // =========================
    // Main
    // =========================

    if (app.books.length === 0) {
        alert("ブックが開かれていません。対象の .indb を開いてから実行してください。");
        return;
    }

    var book = app.activeBook;
    if (!book || !book.isValid) {
        alert("アクティブなブックを取得できませんでした。");
        return;
    }

    var bookFile = File(book.fullName);
    if (!bookFile.exists) {
        alert("ブックファイルが見つかりません。");
        return;
    }

    var bookName = stripExtension(decodeURI(bookFile.name));
    var defaultOutputFolderPath = bookFile.parent.fsName + "/" + sanitizeFileName(bookName) + "_PDF";
    var presetOptions = getPdfPresetOptions();

    var originalInteractionLevel = app.scriptPreferences.userInteractionLevel;
    var successCount = 0;
    var failureCount = 0;
    var logs = [];

    if (book.bookContents.length === 0) {
        alert("ブックにドキュメントが含まれていません。");
        return;
    }

    var exportSettings = showExportConfirmationDialog({
        bookName: decodeURI(bookFile.name),
        documentCount: book.bookContents.length,
        outputFolderPath: defaultOutputFolderPath,
        filenamePrefix: FILENAME_PREFIX,
        presetOptions: presetOptions,
        defaultPresetName: PDF_PRESET_NAME
    });

    if (!exportSettings) {
        return;
    }

    var outputFolder = Folder(exportSettings.outputFolderPath);
    if (!outputFolder.exists) {
        if (!outputFolder.create()) {
            alert("出力フォルダを作成できませんでした:\r" + outputFolder.fsName);
            return;
        }
    }

    var preset = getPdfPresetByName(exportSettings.presetName);
    if (exportSettings.presetName !== "" && (!preset || !preset.isValid)) {
        alert("選択された PDF 書き出しプリセットが見つかりません:\r" + exportSettings.presetName);
        return;
    }

    var filenamePrefix = exportSettings.filenamePrefix;
    var progressState = createProgressStateWindow({
        totalCount: book.bookContents.length,
        outputFolderPath: outputFolder.fsName,
        pdfSettingLabel: getPdfExportSettingLabel(preset),
        filenamePrefix: filenamePrefix
    });
    var wasAborted = false;

    try {
        progressState.show();
        progressState.update({
            phase: "開始準備中",
            currentFileName: "",
            processedCount: 0,
            successCount: successCount,
            failureCount: failureCount,
            lastMessage: "書き出しを開始します。"
        });

        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;

        if (UPDATE_BOOK_NUMBERING) {
            progressState.pump();
            try {
                book.updateAllNumbers();
                progressState.update({
                    phase: "開始準備中",
                    currentFileName: "",
                    processedCount: 0,
                    successCount: successCount,
                    failureCount: failureCount,
                    lastMessage: "ブック番号を更新しました。"
                });
            } catch (eUpdate) {
                logs.push("[WARN] ブック番号更新に失敗: " + eUpdate);
                progressState.update({
                    phase: "開始準備中",
                    currentFileName: "",
                    processedCount: 0,
                    successCount: successCount,
                    failureCount: failureCount,
                    lastMessage: "ブック番号更新に失敗しました。"
                });
            }
        }

        for (var i = 0; i < book.bookContents.length; i++) {
            var bc = book.bookContents[i];
            var docFile = File(bc.fullName);
            var currentFileName = decodeURI(docFile.name);

            progressState.update({
                phase: "書き出し待機中",
                currentFileName: currentFileName,
                processedCount: successCount + failureCount,
                successCount: successCount,
                failureCount: failureCount,
                lastMessage: "次のドキュメントを書き出します。"
            });
            progressState.pump();

            if (progressState.shouldAbort()) {
                wasAborted = true;
                logs.push("[ABORT] ユーザーが中断しました。");
                break;
            }

            if (!docFile.exists) {
                failureCount++;
                logs.push("[ERROR] ファイルが見つかりません: " + docFile.fsName);
                progressState.update({
                    phase: "エラー",
                    currentFileName: currentFileName,
                    processedCount: successCount + failureCount,
                    successCount: successCount,
                    failureCount: failureCount,
                    lastMessage: "ファイルが見つかりません。"
                });
                continue;
            }

            var doc = null;

            try {
                progressState.update({
                    phase: "ドキュメントを開いています",
                    currentFileName: currentFileName,
                    processedCount: successCount + failureCount,
                    successCount: successCount,
                    failureCount: failureCount,
                    lastMessage: "ドキュメントを開いています。"
                });
                doc = app.open(docFile, false);

                var pdfBaseName = filenamePrefix + stripExtension(decodeURI(docFile.name));
                var pdfName = sanitizeFileName(pdfBaseName) + ".pdf";
                var pdfFile = File(outputFolder.fsName + "/" + pdfName);

                if (pdfFile.exists && !OVERWRITE_EXISTING) {
                    failureCount++;
                    logs.push("[SKIP] 既存ファイルあり: " + pdfFile.fsName);
                    progressState.update({
                        phase: "スキップ",
                        currentFileName: currentFileName,
                        processedCount: successCount + failureCount,
                        successCount: successCount,
                        failureCount: failureCount,
                        lastMessage: "既存ファイルがあるためスキップしました。"
                    });
                    doc.close(SaveOptions.no);
                    doc = null;
                    continue;
                }

                progressState.update({
                    phase: "PDF 書き出し中",
                    currentFileName: currentFileName,
                    processedCount: successCount + failureCount,
                    successCount: successCount,
                    failureCount: failureCount,
                    lastMessage: pdfFile.name
                });
                if (preset) {
                    doc.exportFile(ExportFormat.pdfType, pdfFile, false, preset);
                } else {
                    doc.exportFile(ExportFormat.pdfType, pdfFile);
                }

                successCount++;
                logs.push("[OK] " + decodeURI(docFile.name) + " -> " + pdfFile.name);
                progressState.update({
                    phase: "完了",
                    currentFileName: currentFileName,
                    processedCount: successCount + failureCount,
                    successCount: successCount,
                    failureCount: failureCount,
                    lastMessage: "書き出し完了: " + pdfFile.name
                });
            } catch (eDoc) {
                failureCount++;
                logs.push("[ERROR] " + decodeURI(docFile.name) + ": " + eDoc);
                progressState.update({
                    phase: "エラー",
                    currentFileName: currentFileName,
                    processedCount: successCount + failureCount,
                    successCount: successCount,
                    failureCount: failureCount,
                    lastMessage: "エラー: " + eDoc
                });
            } finally {
                if (doc && doc.isValid) {
                    try {
                        doc.close(SaveOptions.no);
                    } catch (eClose) {
                        logs.push("[WARN] クローズ失敗: " + decodeURI(docFile.name) + ": " + eClose);
                    }
                }
            }
        }

        progressState.update({
            phase: wasAborted ? "中断" : "完了",
            currentFileName: "",
            processedCount: successCount + failureCount,
            successCount: successCount,
            failureCount: failureCount,
            lastMessage: wasAborted ? "ユーザーが処理を中断しました。" : "すべての処理が完了しました。"
        });

        alert(
            (wasAborted ? "PDF 書き出しを中断しました。" : "PDF 書き出しが完了しました。") + "\r\r" +
            "成功: " + successCount + "\r" +
            "失敗: " + failureCount + "\r" +
            "出力先: " + outputFolder.fsName + "\r\r" +
            "詳細ログ:\r" + logs.join("\r")
        );
    } catch (e) {
        alert("処理中にエラーが発生しました:\r" + e);
    } finally {
        app.scriptPreferences.userInteractionLevel = originalInteractionLevel;
        progressState.close();
    }

    // =========================
    // Helpers
    // =========================

    function stripExtension(filename) {
        return filename.replace(/\.[^\.]+$/, "");
    }

    function sanitizeFileName(name) {
        // Windows / macOS で問題になりやすい文字を置換
        return name.replace(/[\\\/\:\*\?\"\<\>\|]/g, "_");
    }

    function getPdfExportSettingLabel(preset) {
        if (preset && preset.isValid) {
            return 'プリセット "' + preset.name + '"';
        }

        var prefs = app.pdfExportPreferences;
        var details = ["現在の PDF 書き出し設定"];

        try {
            if (prefs.pageRange === PageRange.ALL_PAGES) {
                details.push("ページ範囲: すべて");
            } else {
                details.push("ページ範囲: " + prefs.pageRange);
            }
        } catch (ePageRange) {}

        try {
            details.push("書き出し単位: " + (prefs.exportReaderSpreads ? "見開き" : "ページ"));
        } catch (eSpreads) {}

        return details.join(" / ");
    }

    function getPdfPresetOptions() {
        var options = [{ label: "現在の PDF 書き出し設定", presetName: "" }];

        for (var i = 0; i < app.pdfExportPresets.length; i++) {
            var preset = app.pdfExportPresets[i];
            options.push({
                label: preset.name,
                presetName: preset.name
            });
        }

        return options;
    }

    function getPdfPresetByName(presetName) {
        if (presetName === "") {
            return null;
        }

        return app.pdfExportPresets.itemByName(presetName);
    }

    function trimString(value) {
        return value.replace(/^\s+|\s+$/g, "");
    }

    function createProgressStateWindow(options) {
        var state = {
            aborted: false,
            allowClose: false,
            window: new Window("palette", "PDF 書き出し状態")
        };

        state.window.orientation = "column";
        state.window.alignChildren = ["fill", "top"];
        state.window.spacing = 10;

        var summaryPanel = state.window.add("panel", undefined, "設定");
        summaryPanel.orientation = "column";
        summaryPanel.alignChildren = ["fill", "top"];
        summaryPanel.margins = 12;
        summaryPanel.add("statictext", undefined, "出力先: " + options.outputFolderPath);
        summaryPanel.add("statictext", undefined, "PDF 書き出し設定: " + options.pdfSettingLabel);
        summaryPanel.add("statictext", undefined, "ファイル名プレフィックス: " + (options.filenamePrefix !== "" ? options.filenamePrefix : "(なし)"));

        var statusPanel = state.window.add("panel", undefined, "進行状況");
        statusPanel.orientation = "column";
        statusPanel.alignChildren = ["fill", "top"];
        statusPanel.margins = 12;

        var phaseText = statusPanel.add("statictext", undefined, "状態: 待機中");
        var fileText = statusPanel.add("statictext", undefined, "対象: -");
        var countText = statusPanel.add("statictext", undefined, "進捗: 0 / " + options.totalCount + "  成功: 0  失敗: 0");
        var progressBar = statusPanel.add("progressbar", undefined, 0, options.totalCount);
        progressBar.preferredSize = [560, 18];
        var messageField = statusPanel.add("edittext", undefined, "", {
            multiline: true,
            readonly: true
        });
        messageField.preferredSize = [560, 70];

        var buttonGroup = state.window.add("group");
        buttonGroup.alignment = "right";
        var abortButton = buttonGroup.add("button", undefined, "中断");

        function safeRefresh() {
            try {
                if (state.window && state.window.visible) {
                    state.window.update();
                }
            } catch (eUpdateWindow) {}
        }

        state.window.onClose = function () {
            if (state.allowClose) {
                return true;
            }

            state.aborted = true;
            phaseText.text = "状態: 中断要求";
            messageField.text = "現在のファイル処理が終わり次第、中断します。";
            safeRefresh();
            return false;
        };

        abortButton.onClick = function () {
            state.aborted = true;
            phaseText.text = "状態: 中断要求";
            messageField.text = "現在のファイル処理が終わり次第、中断します。";
            safeRefresh();
        };

        state.show = function () {
            state.window.show();
        };

        state.update = function (info) {
            var processedCount = info.processedCount;
            if (processedCount < 0) {
                processedCount = 0;
            }
            if (processedCount > options.totalCount) {
                processedCount = options.totalCount;
            }

            phaseText.text = "状態: " + info.phase;
            fileText.text = "対象: " + (info.currentFileName !== "" ? info.currentFileName : "-");
            countText.text =
                "進捗: " + processedCount + " / " + options.totalCount +
                "  成功: " + info.successCount +
                "  失敗: " + info.failureCount;
            progressBar.value = processedCount;
            messageField.text = info.lastMessage || "";
            safeRefresh();
        };

        state.pump = function () {
            safeRefresh();
            $.sleep(50);
        };

        state.shouldAbort = function () {
            return state.aborted;
        };

        state.close = function () {
            try {
                if (state.window && state.window.visible) {
                    state.allowClose = true;
                    state.window.close();
                }
            } catch (eCloseWindow) {}
        };

        return state;
    }

    function showExportConfirmationDialog(options) {
        var dialog = new Window("dialog", "PDF 書き出し確認");
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.spacing = 10;

        dialog.add("statictext", undefined, "以下の内容で PDF 書き出しを開始しますか？");

        var targetPanel = dialog.add("panel", undefined, "対象");
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.margins = 12;
        targetPanel.add("statictext", undefined, "ブック: " + options.bookName);
        targetPanel.add("statictext", undefined, "ドキュメント数: " + options.documentCount);

        var presetPanel = dialog.add("panel", undefined, "PDF 書き出し設定");
        presetPanel.orientation = "column";
        presetPanel.alignChildren = ["fill", "top"];
        presetPanel.margins = 12;

        var presetDropdown = presetPanel.add("dropdownlist", undefined, []);
        for (var i = 0; i < options.presetOptions.length; i++) {
            var item = presetDropdown.add("item", options.presetOptions[i].label);
            item.presetName = options.presetOptions[i].presetName;
        }
        presetDropdown.selection = findPresetSelection(presetDropdown, options.defaultPresetName);

        var presetDetailsField = presetPanel.add("edittext", undefined, "", {
            multiline: true,
            readonly: true
        });
        presetDetailsField.preferredSize = [560, 55];

        var filePanel = dialog.add("panel", undefined, "出力");
        filePanel.orientation = "column";
        filePanel.alignChildren = ["fill", "top"];
        filePanel.margins = 12;

        filePanel.add("statictext", undefined, "ファイル名プレフィックス");
        var prefixField = filePanel.add("edittext", undefined, options.filenamePrefix);
        prefixField.characters = 40;

        filePanel.add("statictext", undefined, "出力先フォルダ");
        var outputPathField = filePanel.add("edittext", undefined, options.outputFolderPath);
        outputPathField.characters = 60;

        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "right";

        var cancelButton = buttonGroup.add("button", undefined, "キャンセル", { name: "cancel" });
        var okButton = buttonGroup.add("button", undefined, "書き出し開始", { name: "ok" });

        dialog.cancelElement = cancelButton;
        dialog.defaultElement = okButton;

        function updatePresetDetails() {
            var selectedItem = presetDropdown.selection;
            var selectedPreset = getPdfPresetByName(selectedItem ? selectedItem.presetName : "");
            presetDetailsField.text = getPdfExportSettingLabel(selectedPreset);
        }

        updatePresetDetails();
        presetDropdown.onChange = updatePresetDetails;

        var result = null;
        okButton.onClick = function () {
            var outputFolderPath = trimString(outputPathField.text);

            if (outputFolderPath === "") {
                alert("出力先フォルダを入力してください。");
                outputPathField.active = true;
                return;
            }

            result = {
                presetName: presetDropdown.selection ? presetDropdown.selection.presetName : "",
                filenamePrefix: prefixField.text,
                outputFolderPath: outputFolderPath
            };

            dialog.close(1);
        };

        return dialog.show() === 1 ? result : null;
    }

    function findPresetSelection(dropdown, presetName) {
        if (presetName !== "") {
            for (var i = 0; i < dropdown.items.length; i++) {
                if (dropdown.items[i].presetName === presetName) {
                    return dropdown.items[i];
                }
            }
        }

        return dropdown.items.length > 0 ? dropdown.items[0] : null;
    }
})();
