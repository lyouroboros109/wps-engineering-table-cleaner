/**
 * WPS 工程表清理助手 v1.8.0
 * - 支持当前工作表 / 整个工作簿
 * - 支持按指定列为空直接删行，并兼容纵向合并区域缩短重建
 * - 支持跨工作表精确查找并清空相同文字内容，不删除行列
 * - 修复纵向合并缩短后的内容核验误报
 * - 支持整个工作簿批量切换分页预览/普通视图
 * - 支持功能区按钮直接运行
 * - 支持删除多余列与调整打印范围
 * - 空白行、预览删除支持跨工作表激活后执行
 * - 条件删行固定保护中文大项，删除采用锚点实移核验
 * - 删除前在当前WPS中实际自检多种删行接口，只使用内容和合并结构均真实上移的方法
 * - 支持按实际表格边界重建打印区域与水平分页线
 * - 支持PDF导出文字截断、打印范围遗漏和分页穿越合并单元格检测
 */

function 工程表清理_快速模式开始(label) {
    var state = {
        screenUpdating: null,
        enableEvents: null,
        calculation: null,
        statusBar: null
    };
    try { state.screenUpdating = Application.ScreenUpdating; } catch (e1) {}
    try { state.enableEvents = Application.EnableEvents; } catch (e2) {}
    try { state.calculation = Application.Calculation; } catch (e3) {}
    try { state.statusBar = Application.StatusBar; } catch (e4) {}
    try { Application.ScreenUpdating = false; } catch (e5) {}
    try { Application.EnableEvents = false; } catch (e6) {}
    try { Application.Calculation = -4135; } catch (e7) {}
    try { Application.StatusBar = label || "工程表清理助手正在处理…"; } catch (e8) {}
    return state;
}

function 工程表清理_快速模式结束(state) {
    try { Application.CutCopyMode = false; } catch (e0) {}
    try { Application.StatusBar = false; } catch (e1) {}
    if (!state) return;
    try { if (state.calculation !== null) Application.Calculation = state.calculation; } catch (e2) {}
    try { if (state.enableEvents !== null) Application.EnableEvents = state.enableEvents; } catch (e3) {}
    try { if (state.screenUpdating !== null) Application.ScreenUpdating = state.screenUpdating; } catch (e4) {}
}

function 工程表清理_内部执行(action, scopeMode) {
    "use strict";

    var PREVIEW_SHEET_NAME = "工程表_删除预览";
    var PREVIEW_DATA_START_ROW = 10;
    var DUPLICATE_SHEET_NAME = "工程表_重复内容选择";
    var DUPLICATE_DATA_START_ROW = 8;
    var UNDO_META_SHEET_NAME = "__工程清理撤回信息__";
    var UNDO_BACKUP_PREFIX = "__ECT_UNDO_";
    var MAX_DUPLICATE_RECORDS = 50000;
    var MAX_TEXT_SEARCH_CELLS = 500000;
    var MAX_BORDER_SCAN_CELLS = 300000;
    var MAX_PDF_SCAN_CELLS = 350000;
    var MAX_PDF_ISSUES = 120;
    var BORDER_CORRECTIVE_PASSES = 3;
    var JS_YES_NO = 4;
    var JS_QUESTION = 32;
    var JS_EXCLAMATION = 48;
    var JS_INFORMATION = 64;
    var JS_RESULT_YES = 6;

    function normalizeText(value) {
        if (value === null || value === undefined) return "";
        return String(value).replace(/\u00a0/g, " ").replace(/[\r\n\t]+/g, " ").trim();
    }

    function clipPromptText(value, maxLength) {
        var text = normalizeText(value);
        var limit = Number(maxLength);
        if (!isFinite(limit) || limit < 1) limit = 120;
        limit = Math.floor(limit);
        if (text.length <= limit) return text;
        if (limit === 1) return "…";
        return text.substring(0, limit - 1) + "…";
    }

    function parseNumberSelection(text, maxValue) {
        var normalized = normalizeText(text)
            .replace(/[，；;、\s]+/g, ",")
            .replace(/[—–~～至]/g, "-");
        if (!normalized) return [];

        var upper = Number(maxValue);
        if (!isFinite(upper) || upper < 1) return [];
        upper = Math.floor(upper);

        var parts = normalized.split(",");
        var selected = {};
        var result = [];

        for (var i = 0; i < parts.length; i += 1) {
            var part = normalizeText(parts[i]);
            if (!part) continue;

            var rangeMatch = part.match(/^(\d+)\-(\d+)$/);
            if (rangeMatch) {
                var startValue = Number(rangeMatch[1]);
                var endValue = Number(rangeMatch[2]);
                if (startValue > endValue) {
                    var swap = startValue;
                    startValue = endValue;
                    endValue = swap;
                }
                for (
                    var value = startValue;
                    value <= endValue;
                    value += 1
                ) {
                    if (
                        value >= 1 &&
                        value <= upper &&
                        !selected[value]
                    ) {
                        selected[value] = true;
                        result.push(value);
                    }
                }
                continue;
            }

            var number = Number(part);
            if (
                isFinite(number) &&
                Math.floor(number) === number &&
                number >= 1 &&
                number <= upper &&
                !selected[number]
            ) {
                selected[number] = true;
                result.push(number);
            }
        }

        result.sort(function (a, b) {
            return a - b;
        });
        return result;
    }

    function isBlankValue(value) {
        return normalizeText(value) === "";
    }

    function columnToLetters(columnNumber) {
        var n = Number(columnNumber);
        if (!isFinite(n) || n < 1) return "";
        n = Math.floor(n);
        var result = "";
        while (n > 0) {
            var remainder = (n - 1) % 26;
            result = String.fromCharCode(65 + remainder) + result;
            n = Math.floor((n - 1) / 26);
        }
        return result;
    }

    function getTargetContext(allowPreviewSheet) {
        var workbook = Application.ActiveWorkbook;
        var sheet = Application.ActiveSheet;
        if (!workbook || !sheet) throw new Error("没有可操作的活动工作簿或工作表。");
        if (!allowPreviewSheet && isAssistantInternalSheetName(String(sheet.Name))) {
            throw new Error("当前位于插件内部工作表。请先切回需要处理的原工作表。");
        }
        return { workbook: workbook, sheet: sheet };
    }

    function getUsedBounds(sheet) {
        var used = sheet.UsedRange;
        var firstRow = Number(used.Row);
        var firstColumn = Number(used.Column);
        var rowCount = Number(used.Rows.Count);
        var columnCount = Number(used.Columns.Count);
        if (!isFinite(firstRow) || !isFinite(firstColumn) || !isFinite(rowCount) || !isFinite(columnCount)) {
            throw new Error("无法读取工作表“" + sheet.Name + "”的已用区域。");
        }
        return {
            firstRow: firstRow,
            firstColumn: firstColumn,
            rowCount: rowCount,
            columnCount: columnCount,
            lastRow: firstRow + rowCount - 1,
            lastColumn: firstColumn + columnCount - 1
        };
    }

    function cellHasFormula(cell) {
        try { return !!cell.HasFormula; } catch (e) { return false; }
    }

    function readMergedAwareValue(cell) {
        var value = null;
        try { value = cell.Value2; } catch (e) { return null; }
        if (!isBlankValue(value)) return value;
        try {
            if (cell.MergeCells) return cell.MergeArea.Cells.Item(1, 1).Value2;
        } catch (e2) {}
        return value;
    }

    function getWritableMergedTopLeft(cell) {
        try {
            if (cell.MergeCells) {
                return cell.MergeArea.Cells.Item(1, 1);
            }
        } catch (ignored) {}
        return cell;
    }

    function readOwnRowValue(cell, rowNumber) {
        try {
            if (cell.MergeCells) {
                var area = cell.MergeArea;
                if (Number(area.Row) !== Number(rowNumber)) return null;
                return area.Cells.Item(1, 1).Value2;
            }
        } catch (e) {}
        try { return cell.Value2; } catch (e2) { return null; }
    }

    function isTopLeftOfMerge(cell) {
        try {
            if (!cell.MergeCells) return true;
            var area = cell.MergeArea;
            return Number(cell.Row) === Number(area.Row) && Number(cell.Column) === Number(area.Column);
        } catch (e) { return true; }
    }

    function inspectRow(sheet, rowNumber, bounds) {
        for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
            var cell = sheet.Cells.Item(rowNumber, column);
            if (cellHasFormula(cell)) return { isBlank: false };
            /*
             * 行只要处于一个有文字的合并区域内，就不属于“完全空白行”。
             * 这样不会误删跨行标题、工程名称和分页标题的组成行。
             */
            if (!isBlankValue(readMergedAwareValue(cell))) return { isBlank: false };
        }
        return { isBlank: true };
    }

    function rowHasAnyContent(sheet, rowNumber, bounds) {
        for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
            var cell = sheet.Cells.Item(rowNumber, column);
            if (cellHasFormula(cell)) return true;
            if (!isBlankValue(readMergedAwareValue(cell))) return true;
        }
        return false;
    }

    function sheetHasAnyContent(sheet, bounds) {
        for (var row = bounds.firstRow; row <= bounds.lastRow; row += 1) {
            if (rowHasAnyContent(sheet, row, bounds)) return true;
        }
        return false;
    }

    function getActualContentRowBounds(sheet, bounds) {
        var firstContentRow = null;
        var lastContentRow = null;
        for (var row = bounds.firstRow; row <= bounds.lastRow; row += 1) {
            if (!rowHasAnyContent(sheet, row, bounds)) continue;
            if (firstContentRow === null) firstContentRow = row;
            lastContentRow = row;
        }
        if (firstContentRow === null) return null;
        return { firstRow: firstContentRow, lastRow: lastContentRow };
    }

    function rowIntersectsVerticalMerge(sheet, rowNumber, bounds) {
        for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
            try {
                var cell = sheet.Cells.Item(rowNumber, column);
                if (!cell.MergeCells) continue;
                if (Number(cell.MergeArea.Rows.Count) > 1) return true;
            } catch (error) {}
        }
        return false;
    }

    function summarizeRow(sheet, rowNumber, bounds, maxItems, maxLength) {
        var items = [];
        var hidden = 0;
        var itemLimit = maxItems || 14;
        var lengthLimit = maxLength || 700;
        for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
            var cell = sheet.Cells.Item(rowNumber, column);
            if (!isTopLeftOfMerge(cell)) continue;
            var text = normalizeText(readMergedAwareValue(cell));
            if (!text && cellHasFormula(cell)) text = "（公式）";
            if (!text) continue;
            if (items.length >= itemLimit) { hidden += 1; continue; }
            items.push(columnToLetters(column) + "：" + text);
        }
        var result = items.join(" ｜ ");
        if (hidden > 0) result += " ｜ 另有" + hidden + "项";
        if (result.length > lengthLimit) result = result.substring(0, lengthLimit - 1) + "…";
        return result;
    }

    function rowSignature(sheet, rowNumber, bounds) {
        var values = [];
        for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
            var cell = sheet.Cells.Item(rowNumber, column);
            if (!isTopLeftOfMerge(cell)) continue;
            var text = normalizeText(readMergedAwareValue(cell));
            var formula = "";
            try { if (cell.HasFormula) formula = normalizeText(cell.Formula); } catch (e) {}
            values.push(column + "=" + text + "|" + formula);
        }
        var textAll = values.join("\u001f");
        if (textAll.length > 3000) textAll = textAll.substring(0, 3000);
        return textAll;
    }

    function uniqueSortedRowsDescending(rows) {
        var map = {};
        var result = [];
        for (var i = 0; i < rows.length; i += 1) {
            var row = Number(rows[i]);
            if (!isFinite(row) || row < 1 || Math.floor(row) !== row || map[row]) continue;
            map[row] = true;
            result.push(row);
        }
        result.sort(function (a, b) { return b - a; });
        return result;
    }

    function countDeletedLessThan(values, target) {
        var count = 0;
        for (var i = 0; i < values.length; i += 1) if (values[i] < target) count += 1;
        return count;
    }

    function countDeletedInRange(values, startValue, endValue) {
        var count = 0;
        for (var i = 0; i < values.length; i += 1) {
            if (values[i] >= startValue && values[i] <= endValue) count += 1;
        }
        return count;
    }

    function saveMergeArea(area) {
        var topLeft = area.Cells.Item(1, 1);
        var saved = {
            row: Number(area.Row),
            column: Number(area.Column),
            rowCount: Number(area.Rows.Count),
            columnCount: Number(area.Columns.Count),
            value: null,
            formula: "",
            hasFormula: false,
            horizontalAlignment: null,
            verticalAlignment: null,
            wrapText: null,
            numberFormat: null
        };
        try { saved.value = topLeft.Value2; } catch (e) {}
        try { saved.hasFormula = !!topLeft.HasFormula; if (saved.hasFormula) saved.formula = topLeft.Formula; } catch (e2) {}
        try { saved.horizontalAlignment = area.HorizontalAlignment; } catch (e3) {}
        try { saved.verticalAlignment = area.VerticalAlignment; } catch (e4) {}
        try { saved.wrapText = area.WrapText; } catch (e5) {}
        try { saved.numberFormat = topLeft.NumberFormat; } catch (e6) {}
        return saved;
    }

    function collectAffectedRowMerges(sheet, rows, bounds) {
        var map = {};
        var result = [];
        for (var r = 0; r < rows.length; r += 1) {
            var row = rows[r];
            for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
                var cell = sheet.Cells.Item(row, column);
                try {
                    if (!cell.MergeCells) continue;
                    var area = cell.MergeArea;
                    if (Number(area.Rows.Count) <= 1) continue;
                    var key = area.Row + ":" + area.Column + ":" + area.Rows.Count + ":" + area.Columns.Count;
                    if (map[key]) continue;
                    map[key] = true;
                    result.push(saveMergeArea(area));
                } catch (e) {}
            }
        }
        return result;
    }

    function unmergeSavedAreas(sheet, merges) {
        for (var i = 0; i < merges.length; i += 1) {
            var item = merges[i];
            var address = columnToLetters(item.column) + item.row + ":" +
                columnToLetters(item.column + item.columnCount - 1) + (item.row + item.rowCount - 1);
            sheet.Range(address).UnMerge();
        }
    }

    function restoreRowMerges(sheet, merges, deletedRows) {
        var ascending = uniqueSortedRowsDescending(deletedRows).slice().sort(function (a, b) { return a - b; });
        for (var i = 0; i < merges.length; i += 1) {
            var item = merges[i];
            var oldEndRow = item.row + item.rowCount - 1;
            var removedInside = countDeletedInRange(ascending, item.row, oldEndRow);
            var remainingRows = item.rowCount - removedInside;
            if (remainingRows <= 0) continue;
            var newRow = item.row - countDeletedLessThan(ascending, item.row);
            var newEndRow = newRow + remainingRows - 1;
            var address = columnToLetters(item.column) + newRow + ":" +
                columnToLetters(item.column + item.columnCount - 1) + newEndRow;
            var area = sheet.Range(address);
            if (remainingRows > 1 || item.columnCount > 1) area.Merge();
            var topLeft = sheet.Cells.Item(newRow, item.column);
            try { if (item.hasFormula && item.formula) topLeft.Formula = item.formula; else topLeft.Value2 = item.value; } catch (e1) {}
            try { if (item.horizontalAlignment !== null) area.HorizontalAlignment = item.horizontalAlignment; } catch (e2) {}
            try { if (item.verticalAlignment !== null) area.VerticalAlignment = item.verticalAlignment; } catch (e3) {}
            try { if (item.wrapText !== null) area.WrapText = item.wrapText; } catch (e4) {}
            try { if (item.numberFormat !== null) topLeft.NumberFormat = item.numberFormat; } catch (e5) {}
        }
    }

    function activateWorksheet(workbook, sheetName) {
        if (!workbook) throw new Error("没有目标工作簿。");
        try { workbook.Activate(); } catch (workbookActivateError) {}
        var sheet = findWorksheet(workbook, sheetName);
        if (!sheet) throw new Error("找不到工作表：“" + sheetName + "”。");
        try {
            if (sheet.Visible === false || Number(sheet.Visible) === 0) {
                throw new Error("工作表已隐藏，无法安全处理。");
            }
        } catch (visibilityError) {
            if (visibilityError && visibilityError.message) throw visibilityError;
        }
        sheet.Activate();
        try { workbook.Activate(); sheet.Activate(); } catch (repeatActivateError) {}
        var activeBook = Application.ActiveWorkbook;
        var activeSheet = Application.ActiveSheet;
        if (!activeBook || String(activeBook.Name) !== String(workbook.Name)) {
            throw new Error("WPS 未能激活目标工作簿：“" + workbook.Name + "”。");
        }
        if (!activeSheet || String(activeSheet.Name) !== String(sheetName)) {
            throw new Error("WPS 未能激活工作表：“" + sheetName + "”。");
        }
        return Application.ActiveSheet;
    }

    /*
     * v0.9.0：执行前在临时工作表中测试多种删行接口。
     * 实际删除时按连续区块从下往上处理，并核对下方内容及合并结构是否真正上移。
     * 自检或核验失败时立即停止，不再仅凭接口无报错就提示成功。
     */

    function copyBasicCellFormat(sourceCell, destinationCell) {
        try { destinationCell.NumberFormat = sourceCell.NumberFormat; } catch (e1) {}
        try { destinationCell.HorizontalAlignment = sourceCell.HorizontalAlignment; } catch (e2) {}
        try { destinationCell.VerticalAlignment = sourceCell.VerticalAlignment; } catch (e3) {}
        try { destinationCell.WrapText = sourceCell.WrapText; } catch (e4) {}
        try { destinationCell.Orientation = sourceCell.Orientation; } catch (e5) {}
        try { destinationCell.IndentLevel = sourceCell.IndentLevel; } catch (e6) {}
        try { destinationCell.ShrinkToFit = sourceCell.ShrinkToFit; } catch (e7) {}
        try { destinationCell.NumberFormatLocal = sourceCell.NumberFormatLocal; } catch (e8) {}
        try {
            destinationCell.Font.Name = sourceCell.Font.Name;
            destinationCell.Font.Size = sourceCell.Font.Size;
            destinationCell.Font.Bold = sourceCell.Font.Bold;
            destinationCell.Font.Italic = sourceCell.Font.Italic;
            destinationCell.Font.Underline = sourceCell.Font.Underline;
            destinationCell.Font.Color = sourceCell.Font.Color;
        } catch (fontError) {}
        try { destinationCell.Interior.Color = sourceCell.Interior.Color; } catch (fillError) {}
        var borderIndexes = [7, 8, 9, 10, 11, 12];
        for (var i = 0; i < borderIndexes.length; i += 1) {
            try {
                var sourceBorder = sourceCell.Borders.Item(borderIndexes[i]);
                var destinationBorder = destinationCell.Borders.Item(borderIndexes[i]);
                destinationBorder.LineStyle = sourceBorder.LineStyle;
                destinationBorder.Weight = sourceBorder.Weight;
                destinationBorder.Color = sourceBorder.Color;
            } catch (borderError) {}
        }
    }

    function copyCellValueAndFormula(sourceCell, destinationCell) {
        if (cellHasFormula(sourceCell)) {
            try {
                destinationCell.FormulaR1C1 = sourceCell.FormulaR1C1;
                return;
            } catch (e1) {}
            try {
                destinationCell.Formula = sourceCell.Formula;
                return;
            } catch (e2) {}
        }
        try { destinationCell.Value2 = sourceCell.Value2; }
        catch (e3) { destinationCell.Value = sourceCell.Value; }
    }

    function copyRowReliable(sheet, sourceRow, destinationRow, bounds) {
        if (sourceRow === destinationRow) return;
        var sourceAddress = columnToLetters(bounds.firstColumn) + sourceRow + ":" +
            columnToLetters(bounds.lastColumn) + sourceRow;
        var destinationAddress = columnToLetters(bounds.firstColumn) + destinationRow + ":" +
            columnToLetters(bounds.lastColumn) + destinationRow;
        var expectedSignature = rowSignature(sheet, sourceRow, bounds);
        var copied = false;

        try {
            sheet.Range(sourceAddress).Copy(sheet.Range(destinationAddress));
            copied = rowSignature(sheet, destinationRow, bounds) === expectedSignature;
        } catch (rangeCopyError) {
            copied = false;
        }

        if (!copied) {
            for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
                var sourceCell = sheet.Cells.Item(sourceRow, column);
                var destinationCell = sheet.Cells.Item(destinationRow, column);
                try { sourceCell.Copy(destinationCell); } catch (cellCopyError) {}
                copyCellValueAndFormula(sourceCell, destinationCell);
                copyBasicCellFormat(sourceCell, destinationCell);
            }
            copied = rowSignature(sheet, destinationRow, bounds) === expectedSignature;
        }

        try { sheet.Rows.Item(destinationRow).RowHeight = sheet.Rows.Item(sourceRow).RowHeight; } catch (heightError) {}
        try { sheet.Rows.Item(destinationRow).Hidden = sheet.Rows.Item(sourceRow).Hidden; } catch (hiddenError) {}

        if (!copied) {
            throw new Error("第" + sourceRow + "行复制到第" + destinationRow + "行后内容核验不一致。");
        }
    }

    function clearTrailingRows(sheet, firstRow, lastRow, bounds) {
        if (firstRow > lastRow) return;
        var address = columnToLetters(bounds.firstColumn) + firstRow + ":" +
            columnToLetters(bounds.lastColumn) + lastRow;
        try { sheet.Range(address).Clear(); }
        catch (clearError) {
            for (var row = firstRow; row <= lastRow; row += 1) {
                for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
                    var cell = sheet.Cells.Item(row, column);
                    try { cell.Clear(); }
                    catch (cellClearError) {
                        try { cell.ClearContents(); } catch (ignored1) {}
                        try { cell.ClearFormats(); } catch (ignored2) {}
                    }
                }
            }
        }
        for (var rowNumber = firstRow; rowNumber <= lastRow; rowNumber += 1) {
            try { sheet.Rows.Item(rowNumber).Hidden = false; } catch (hiddenError) {}
            try { sheet.Rows.Item(rowNumber).RowHeight = 15; } catch (heightError) {}
        }
    }


    function groupRowsIntoBlocksDescending(rows) {
        var descending = uniqueSortedRowsDescending(rows);
        if (descending.length === 0) return [];
        var ascending = descending.slice().sort(function (a, b) { return a - b; });
        var blocks = [];
        var start = ascending[0];
        var end = ascending[0];
        for (var i = 1; i < ascending.length; i += 1) {
            if (ascending[i] === end + 1) {
                end = ascending[i];
            } else {
                blocks.push({ start: start, end: end });
                start = ascending[i];
                end = ascending[i];
            }
        }
        blocks.push({ start: start, end: end });
        blocks.sort(function (a, b) { return b.start - a.start; });
        return blocks;
    }

    function rowMergePattern(sheet, rowNumber, bounds) {
        var seen = {};
        var parts = [];
        for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
            try {
                var cell = sheet.Cells.Item(rowNumber, column);
                if (!cell.MergeCells) continue;
                var area = cell.MergeArea;
                var key = Number(area.Row) + ":" + Number(area.Column) + ":" +
                    Number(area.Rows.Count) + ":" + Number(area.Columns.Count);
                if (seen[key]) continue;
                seen[key] = true;
                parts.push(
                    (Number(area.Column) - bounds.firstColumn) + "," +
                    Number(area.Rows.Count) + "," + Number(area.Columns.Count)
                );
            } catch (error) {}
        }
        parts.sort();
        return parts.join("|");
    }

    function deletionMethodLabel(method) {
        if (method === "rangeShift") return "区域上移";
        if (method === "entireRow") return "整行删除";
        if (method === "rowsItem") return "行集合删除";
        if (method === "selectionShift") return "选区上移";
        return method || "未知";
    }

    function applyDeleteMethod(sheet, method, startRow, endRow, bounds) {
        var firstColumnLetter = columnToLetters(bounds.firstColumn);
        var lastColumnLetter = columnToLetters(bounds.lastColumn);
        var address = firstColumnLetter + startRow + ":" + lastColumnLetter + endRow;

        if (method === "rangeShift") {
            sheet.Range(address).Delete(-4162);
            return;
        }
        if (method === "selectionShift") {
            sheet.Range(address).Select();
            Application.Selection.Delete(-4162);
            return;
        }
        if (method === "entireRow") {
            for (var row1 = endRow; row1 >= startRow; row1 -= 1) {
                sheet.Range("A" + row1).EntireRow.Delete();
            }
            return;
        }
        if (method === "rowsItem") {
            for (var row2 = endRow; row2 >= startRow; row2 -= 1) {
                sheet.Rows.Item(row2).Delete();
            }
            return;
        }
        throw new Error("未知删除方式：" + method);
    }

    function makeTemporarySheetName(workbook, seed) {
        var base = "__工程清理自检" + seed;
        var name = base;
        var index = 1;
        while (findWorksheet(workbook, name)) {
            index += 1;
            name = base.substring(0, 25) + index;
        }
        return name;
    }

    function deleteTemporarySheet(sheet) {
        if (!sheet) return;
        var oldAlerts = Application.DisplayAlerts;
        try {
            Application.DisplayAlerts = false;
            sheet.Delete();
        } catch (error) {
            try {
                sheet.Name = "__自检残留_" + String(new Date().getTime()).slice(-6);
                sheet.Visible = false;
            } catch (ignored) {}
        } finally {
            try { Application.DisplayAlerts = oldAlerts; } catch (ignored2) {}
        }
    }

    function setDeletionSelfTestPattern(sheet, scenario) {
        sheet.Range("A1:G8").Clear();
        try { sheet.Range("A1:G8").UnMerge(); } catch (ignored) {}

        sheet.Range("A1:G1").Merge();
        sheet.Range("A1").Value2 = "工程表清理自检";

        if (scenario === "blank") {
            sheet.Range("A2:G2").Clear();
            sheet.Range("B3:C3").Merge();
            sheet.Range("E3:F3").Merge();
            sheet.Range("A3").Value2 = "ANCHOR-BLANK";
            sheet.Range("B3").Value2 = "合并名称";
            sheet.Range("E3").Value2 = 123;
            sheet.Range("G3").Value2 = "尾";
            sheet.Range("B4:C4").Merge();
            sheet.Range("E4:F4").Merge();
            sheet.Range("A4").Value2 = "TAIL-BLANK";
            sheet.Range("B4").Value2 = "下一行";
            sheet.Range("E4").Value2 = 456;
            return {
                deleteStart: 2,
                deleteEnd: 2,
                anchorRow: 3,
                expectedA: "ANCHOR-BLANK",
                expectedB: "合并名称"
            };
        }

        sheet.Range("B2:C2").Merge();
        sheet.Range("E2:F2").Merge();
        sheet.Range("A2").Value2 = "DELETE-MERGED";
        sheet.Range("B2").Value2 = "删除项";
        sheet.Range("E2").Value2 = 111;

        sheet.Range("B3:C3").Merge();
        sheet.Range("E3:F3").Merge();
        sheet.Range("A3").Value2 = "ANCHOR-MERGED";
        sheet.Range("B3").Value2 = "保留项";
        sheet.Range("E3").Value2 = 222;
        sheet.Range("G3").Value2 = "尾";

        sheet.Range("B4:C4").Merge();
        sheet.Range("E4:F4").Merge();
        sheet.Range("A4").Value2 = "TAIL-MERGED";
        sheet.Range("B4").Value2 = "下一行";
        sheet.Range("E4").Value2 = 333;

        return {
            deleteStart: 2,
            deleteEnd: 2,
            anchorRow: 3,
            expectedA: "ANCHOR-MERGED",
            expectedB: "保留项"
        };
    }

    function testDeletionMethodScenario(workbook, method, scenario, sequence) {
        var originalSheetName = getOriginalActiveSheetName(workbook);
        var testSheet = null;
        try {
            testSheet = workbook.Worksheets.Add();
            testSheet.Name = makeTemporarySheetName(workbook, sequence);
            var setup = setDeletionSelfTestPattern(testSheet, scenario);
            var testBounds = {
                firstRow: 1,
                firstColumn: 1,
                lastRow: 8,
                lastColumn: 7,
                rowCount: 8,
                columnCount: 7
            };
            var expectedSignature = rowSignature(testSheet, setup.anchorRow, testBounds);
            var expectedMergePattern = rowMergePattern(testSheet, setup.anchorRow, testBounds);

            applyDeleteMethod(
                testSheet,
                method,
                setup.deleteStart,
                setup.deleteEnd,
                testBounds
            );
            try { Application.Calculate(); } catch (ignored) {}

            var actualSignature = rowSignature(testSheet, setup.deleteStart, testBounds);
            var actualMergePattern = rowMergePattern(testSheet, setup.deleteStart, testBounds);
            var actualA = normalizeText(testSheet.Cells.Item(setup.deleteStart, 1).Value2);
            var actualB = normalizeText(readMergedAwareValue(testSheet.Cells.Item(setup.deleteStart, 2)));

            if (
                actualSignature !== expectedSignature ||
                actualMergePattern !== expectedMergePattern ||
                actualA !== setup.expectedA ||
                actualB !== setup.expectedB
            ) {
                throw new Error(
                    scenario + "场景核验失败，目标行没有按预期上移。"
                );
            }
            return { ok: true, message: "通过" };
        } catch (error) {
            return {
                ok: false,
                message: error && error.message ? error.message : String(error)
            };
        } finally {
            deleteTemporarySheet(testSheet);
            restoreActiveSheet(workbook, originalSheetName);
        }
    }

    function resolveWorkingDeletionMethod(workbook) {
        var methods = ["rangeShift", "selectionShift", "entireRow", "rowsItem"];
        var results = [];
        for (var i = 0; i < methods.length; i += 1) {
            var method = methods[i];
            var blankResult = testDeletionMethodScenario(
                workbook,
                method,
                "blank",
                String(i + 1) + "A"
            );
            var mergedResult = testDeletionMethodScenario(
                workbook,
                method,
                "merged",
                String(i + 1) + "B"
            );
            if (blankResult.ok && mergedResult.ok) return method;
            results.push(
                deletionMethodLabel(method) + "：" +
                (blankResult.ok ? "空白行通过" : blankResult.message) + "；" +
                (mergedResult.ok ? "合并行通过" : mergedResult.message)
            );
        }
        throw new Error(
            "当前WPS的删行接口自检未通过，插件已停止执行，因此不会再虚报成功。\n" +
            results.join("\n")
        );
    }

    function captureShiftedRowProperties(sheet, startRow, endRow) {
        var result = [];
        for (var row = startRow; row <= endRow; row += 1) {
            var item = { row: row, height: null, hidden: null };
            try { item.height = sheet.Rows.Item(row).RowHeight; } catch (e1) {}
            try { item.hidden = sheet.Rows.Item(row).Hidden; } catch (e2) {}
            result.push(item);
        }
        return result;
    }

    function restoreShiftedRowProperties(sheet, saved, shiftCount) {
        for (var i = 0; i < saved.length; i += 1) {
            var targetRow = saved[i].row - shiftCount;
            if (targetRow < 1) continue;
            try {
                if (saved[i].height !== null) {
                    sheet.Rows.Item(targetRow).RowHeight = saved[i].height;
                }
            } catch (e1) {}
            try {
                if (saved[i].hidden !== null) {
                    sheet.Rows.Item(targetRow).Hidden = saved[i].hidden;
                }
            } catch (e2) {}
        }
    }

    function compactRowsOnSheet(workbook, sheetName, rows, deletionMethod) {
        var requestedRows = uniqueSortedRowsDescending(rows);
        if (requestedRows.length === 0) {
            return { deleted: 0, deletedRows: [], failedRows: [], method: deletionMethod || "" };
        }

        var sheet = activateWorksheet(workbook, sheetName);
        var bounds = getUsedBounds(sheet);
        var contentRows = getActualContentRowBounds(sheet, bounds);
        if (!contentRows) {
            return { deleted: 0, deletedRows: [], failedRows: [], method: deletionMethod || "" };
        }

        var validRows = [];
        var failedRows = [];
        var validMap = {};
        for (var i = 0; i < requestedRows.length; i += 1) {
            var row = requestedRows[i];
            if (row < contentRows.firstRow || row > contentRows.lastRow) {
                failedRows.push({ row: row, message: "行号超出实际内容范围。" });
                continue;
            }
            if (rowIntersectsVerticalMerge(sheet, row, bounds)) {
                failedRows.push({ row: row, message: "与纵向合并单元格相交，已跳过。" });
                continue;
            }
            if (!validMap[row]) {
                validMap[row] = true;
                validRows.push(row);
            }
        }

        if (validRows.length === 0) {
            return {
                deleted: 0,
                deletedRows: [],
                failedRows: failedRows,
                method: deletionMethod || ""
            };
        }

        var method = deletionMethod || resolveWorkingDeletionMethod(workbook);
        var blocks = groupRowsIntoBlocksDescending(validRows);
        var deletedRows = [];
        var currentLastContentRow = contentRows.lastRow;

        for (var blockIndex = 0; blockIndex < blocks.length; blockIndex += 1) {
            var block = blocks[blockIndex];
            var blockCount = block.end - block.start + 1;
            var hasAnchor = block.end < currentLastContentRow;
            var expectedAnchorSignature = "";
            var expectedAnchorMerge = "";
            var expectedLastSignature = "";
            var shiftedProperties = [];

            if (hasAnchor) {
                expectedAnchorSignature = rowSignature(sheet, block.end + 1, bounds);
                expectedAnchorMerge = rowMergePattern(sheet, block.end + 1, bounds);
            } else {
                expectedLastSignature = rowSignature(sheet, block.start, bounds);
            }

            if (method === "rangeShift" || method === "selectionShift") {
                shiftedProperties = captureShiftedRowProperties(
                    sheet,
                    block.end + 1,
                    currentLastContentRow
                );
            }

            applyDeleteMethod(sheet, method, block.start, block.end, bounds);
            try { Application.Calculate(); } catch (ignored) {}

            if (method === "rangeShift" || method === "selectionShift") {
                restoreShiftedRowProperties(sheet, shiftedProperties, blockCount);
            }

            if (hasAnchor) {
                var actualAnchorSignature = rowSignature(sheet, block.start, bounds);
                var actualAnchorMerge = rowMergePattern(sheet, block.start, bounds);
                if (
                    actualAnchorSignature !== expectedAnchorSignature ||
                    actualAnchorMerge !== expectedAnchorMerge
                ) {
                    throw new Error(
                        "工作表“" + sheetName + "”第" + block.start + "至" +
                        block.end + "行删除后，上移内容核验失败。"
                    );
                }
            } else {
                if (rowHasAnyContent(sheet, block.start, bounds)) {
                    var actualLastSignature = rowSignature(sheet, block.start, bounds);
                    if (actualLastSignature === expectedLastSignature) {
                        throw new Error(
                            "工作表“" + sheetName + "”末尾第" + block.start +
                            "至" + block.end + "行没有实际删除。"
                        );
                    }
                }
            }

            for (var deletedRow = block.start; deletedRow <= block.end; deletedRow += 1) {
                deletedRows.push(deletedRow);
            }
            currentLastContentRow -= blockCount;
        }

        return {
            deleted: deletedRows.length,
            deletedRows: deletedRows,
            failedRows: failedRows,
            method: method
        };
    }




    /*
     * v1.8.0 稳定重写引擎：不依赖 WPS 的 Delete 接口。
     * 先保存合并结构，全部拆分，再把保留行逐行写到目标位置，最后重建合并区域。
     * 预览表能写入说明 Value2/Formula 写入可靠；本引擎只使用同一套可核验写入方式。
     */
    function collectContentMergeAreas(sheet, bounds, firstRow, lastRow) {
        var seen = {};
        var result = [];
        for (var row = firstRow; row <= lastRow; row += 1) {
            for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
                try {
                    var cell = sheet.Cells.Item(row, column);
                    if (!cell.MergeCells) continue;
                    var area = cell.MergeArea;
                    var areaRow = Number(area.Row);
                    var areaEnd = areaRow + Number(area.Rows.Count) - 1;
                    if (areaEnd < firstRow || areaRow > lastRow) continue;
                    var key = areaRow + ":" + Number(area.Column) + ":" + Number(area.Rows.Count) + ":" + Number(area.Columns.Count);
                    if (seen[key]) continue;
                    seen[key] = true;
                    result.push(saveMergeArea(area));
                } catch (error) {}
            }
        }
        return result;
    }

    function rawCellFormula(cell) {
        try {
            if (cell.HasFormula) return normalizeText(cell.FormulaR1C1 || cell.Formula || "");
        } catch (error) {}
        return "";
    }

    function rawRowSignature(sheet, rowNumber, bounds) {
        var values = [];
        for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
            var cell = sheet.Cells.Item(rowNumber, column);
            var value = "";
            try { value = normalizeText(cell.Value2); } catch (error1) {}
            values.push(column + "=" + value + "|" + rawCellFormula(cell));
        }
        return values.join("\u001f");
    }

    function copyUnmergedRow(sheet, sourceRow, destinationRow, bounds) {
        if (sourceRow === destinationRow) return;
        var firstLetter = columnToLetters(bounds.firstColumn);
        var lastLetter = columnToLetters(bounds.lastColumn);
        var sourceRange = sheet.Range(firstLetter + sourceRow + ":" + lastLetter + sourceRow);
        var destinationRange = sheet.Range(firstLetter + destinationRow + ":" + lastLetter + destinationRow);
        var expected = rawRowSignature(sheet, sourceRow, bounds);
        var copied = false;

        try {
            destinationRange.Clear();
            sourceRange.Copy(destinationRange);
            copied = rawRowSignature(sheet, destinationRow, bounds) === expected;
        } catch (copyError) {
            copied = false;
        }

        if (!copied) {
            try { destinationRange.Clear(); } catch (clearError) {}
            for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
                var sourceCell = sheet.Cells.Item(sourceRow, column);
                var destinationCell = sheet.Cells.Item(destinationRow, column);
                try { destinationCell.Clear(); } catch (ignored0) {}
                copyCellValueAndFormula(sourceCell, destinationCell);
                copyBasicCellFormat(sourceCell, destinationCell);
            }
            copied = rawRowSignature(sheet, destinationRow, bounds) === expected;
        }

        try { sheet.Rows.Item(destinationRow).RowHeight = sheet.Rows.Item(sourceRow).RowHeight; } catch (heightError) {}
        try { sheet.Rows.Item(destinationRow).Hidden = sheet.Rows.Item(sourceRow).Hidden; } catch (hiddenError) {}
        if (!copied) throw new Error("第" + sourceRow + "行写入第" + destinationRow + "行后核验失败。");
    }

    function rebuildContentMerges(sheet, merges, removedAscending) {
        var oldAlerts = Application.DisplayAlerts;
        var expected = [];
        try { Application.DisplayAlerts = false; } catch (ignored0) {}
        try {
            for (var i = 0; i < merges.length; i += 1) {
                var item = merges[i];
                var oldEndRow = item.row + item.rowCount - 1;
                var removedInside = countDeletedInRange(removedAscending, item.row, oldEndRow);
                var remainingRows = item.rowCount - removedInside;
                if (remainingRows <= 0) continue;
                var newRow = item.row - countDeletedLessThan(removedAscending, item.row);
                var newEndRow = newRow + remainingRows - 1;
                var newEndColumn = item.column + item.columnCount - 1;
                var address = columnToLetters(item.column) + newRow + ":" + columnToLetters(newEndColumn) + newEndRow;
                var area = sheet.Range(address);
                if (remainingRows > 1 || item.columnCount > 1) area.Merge();
                var topLeft = sheet.Cells.Item(newRow, item.column);
                try {
                    if (item.hasFormula && item.formula) topLeft.Formula = item.formula;
                    else topLeft.Value2 = item.value;
                } catch (valueError) {}
                try { if (item.horizontalAlignment !== null) area.HorizontalAlignment = item.horizontalAlignment; } catch (e1) {}
                try { if (item.verticalAlignment !== null) area.VerticalAlignment = item.verticalAlignment; } catch (e2) {}
                try { if (item.wrapText !== null) area.WrapText = item.wrapText; } catch (e3) {}
                try { if (item.numberFormat !== null) topLeft.NumberFormat = item.numberFormat; } catch (e4) {}
                expected.push({row:newRow,column:item.column,rowCount:remainingRows,columnCount:item.columnCount});
            }
        } finally {
            try { Application.DisplayAlerts = oldAlerts; } catch (ignored1) {}
        }
        return expected;
    }

    function verifyRebuiltMerges(sheet, expected) {
        for (var i = 0; i < expected.length; i += 1) {
            var item = expected[i];
            var shouldMerge = item.rowCount > 1 || item.columnCount > 1;
            var cell = sheet.Cells.Item(item.row, item.column);
            if (!shouldMerge) {
                if (cell.MergeCells) throw new Error("单格区域被错误合并：" + columnToLetters(item.column) + item.row);
                continue;
            }
            if (!cell.MergeCells) throw new Error("合并区域没有重建：" + columnToLetters(item.column) + item.row);
            var area = cell.MergeArea;
            if (Number(area.Row)!==item.row || Number(area.Column)!==item.column || Number(area.Rows.Count)!==item.rowCount || Number(area.Columns.Count)!==item.columnCount) {
                throw new Error("合并区域尺寸核验失败：" + columnToLetters(item.column) + item.row);
            }
        }
    }


    function buildRetainedSegments(firstRow, lastRow, removeMap) {
        var segments = [];
        var start = null;
        for (var row = firstRow; row <= lastRow; row += 1) {
            if (removeMap[row]) {
                if (start !== null) {
                    segments.push({ start: start, end: row - 1 });
                    start = null;
                }
            } else if (start === null) start = row;
        }
        if (start !== null) segments.push({ start: start, end: lastRow });
        return segments;
    }

    function copyUnmergedBlock(sheet, sourceStart, sourceEnd, destinationStart, bounds) {
        var rowCount = sourceEnd - sourceStart + 1;
        if (rowCount <= 0 || sourceStart === destinationStart) return;
        var firstLetter = columnToLetters(bounds.firstColumn);
        var lastLetter = columnToLetters(bounds.lastColumn);
        var sourceRange = sheet.Range(firstLetter + sourceStart + ":" + lastLetter + sourceEnd);
        var destinationEnd = destinationStart + rowCount - 1;
        var destinationRange = sheet.Range(firstLetter + destinationStart + ":" + lastLetter + destinationEnd);
        var sourceFirstSignature = rawRowSignature(sheet, sourceStart, bounds);
        var sourceLastSignature = rawRowSignature(sheet, sourceEnd, bounds);
        sourceRange.Copy(destinationRange);
        if (rawRowSignature(sheet, destinationStart, bounds) !== sourceFirstSignature || rawRowSignature(sheet, destinationEnd, bounds) !== sourceLastSignature) {
            throw new Error("连续行块复制核验失败：第" + sourceStart + "-" + sourceEnd + "行。");
        }
    }

    function shouldVerifyRetainedIndex(index, count, segments) {
        if (count <= 3000) return true;
        if (index === 0 || index === count - 1 || index % 200 === 0) return true;
        return false;
    }

    function rewriteRowsOnActiveSheet(rows) {
        var sheet = Application.ActiveSheet;
        if (!sheet) throw new Error("没有活动工作表。");
        var bounds = getUsedBounds(sheet);
        var contentRows = getActualContentRowBounds(sheet, bounds);
        if (!contentRows) return { deleted: 0, deletedRows: [], failedRows: [] };

        var requested = uniqueSortedRowsDescending(rows);
        var removeMap = {};
        var validRows = [];
        var failedRows = [];
        for (var i = 0; i < requested.length; i += 1) {
            var row = requested[i];
            if (row < contentRows.firstRow || row > contentRows.lastRow) {
                failedRows.push({ row: row, message: "超出实际内容范围。" });
                continue;
            }
            if (!removeMap[row]) { removeMap[row] = true; validRows.push(row); }
        }
        if (!validRows.length) return { deleted: 0, deletedRows: [], failedRows: failedRows };
        validRows.sort(function (a, b) { return a - b; });

        var merges = collectContentMergeAreas(
            sheet,
            bounds,
            contentRows.firstRow,
            contentRows.lastRow
        );
        var mergeAffectedRows = {};
        for (var mergeIndex = 0; mergeIndex < merges.length; mergeIndex += 1) {
            var mergeItem = merges[mergeIndex];
            var mergeEnd = mergeItem.row + mergeItem.rowCount - 1;
            if (countDeletedInRange(validRows, mergeItem.row, mergeEnd) <= 0) continue;
            for (var mergeRow = mergeItem.row; mergeRow <= mergeEnd; mergeRow += 1) {
                mergeAffectedRows[mergeRow] = true;
            }
        }

        var retained = [];
        var rowProperties = [];
        var expectedSamples = {};
        for (var sourceRow = contentRows.firstRow; sourceRow <= contentRows.lastRow; sourceRow += 1) {
            if (removeMap[sourceRow]) continue;
            var retainedIndex = retained.length;
            retained.push(sourceRow);
            if (
                !mergeAffectedRows[sourceRow] &&
                shouldVerifyRetainedIndex(
                    retainedIndex,
                    contentRows.lastRow - contentRows.firstRow + 1 - validRows.length,
                    null
                )
            ) {
                /*
                 * 使用 FormulaR1C1 的原始行签名，避免公式随行移动后 A1 引用文本改变而误报。
                 * 与被删除行相交的纵向合并区域由 verifyRebuiltMerges 单独核验。
                 */
                expectedSamples[retainedIndex] = rawRowSignature(sheet, sourceRow, bounds);
            }
            var props = { height: null, hidden: null };
            try { props.height = sheet.Rows.Item(sourceRow).RowHeight; } catch (e1) {}
            try { props.hidden = sheet.Rows.Item(sourceRow).Hidden; } catch (e2) {}
            rowProperties.push(props);
        }
        if (!retained.length) throw new Error("不能移除工作表中的全部实际内容行。");

        var segments = buildRetainedSegments(contentRows.firstRow, contentRows.lastRow, removeMap);
        for (var m = 0; m < merges.length; m += 1) {
            try {
                var mergeAddress = columnToLetters(merges[m].column) + merges[m].row + ":" + columnToLetters(merges[m].column + merges[m].columnCount - 1) + (merges[m].row + merges[m].rowCount - 1);
                sheet.Range(mergeAddress).UnMerge();
            } catch (unmergeError) {
                throw new Error("拆分合并区域失败：" + (unmergeError && unmergeError.message ? unmergeError.message : String(unmergeError)));
            }
        }

        var destinationRow = contentRows.firstRow;
        for (var s = 0; s < segments.length; s += 1) {
            var segment = segments[s];
            copyUnmergedBlock(sheet, segment.start, segment.end, destinationRow, bounds);
            destinationRow += segment.end - segment.start + 1;
        }
        for (var r = 0; r < retained.length; r += 1) {
            var targetRow = contentRows.firstRow + r;
            try { if (rowProperties[r].height !== null) sheet.Rows.Item(targetRow).RowHeight = rowProperties[r].height; } catch (e3) {}
            try { if (rowProperties[r].hidden !== null) sheet.Rows.Item(targetRow).Hidden = rowProperties[r].hidden; } catch (e4) {}
        }

        clearTrailingRows(sheet, destinationRow, contentRows.lastRow, bounds);
        var expectedMerges = rebuildContentMerges(sheet, merges, validRows);
        verifyRebuiltMerges(sheet, expectedMerges);
        try { Application.Calculate(); } catch (calculateError) {}

        var expectedLastRow = contentRows.lastRow - validRows.length;
        var actualRows = getActualContentRowBounds(sheet, bounds);
        if (!actualRows || actualRows.lastRow !== expectedLastRow) {
            throw new Error("内容上移核验失败：预计最后内容行为" + expectedLastRow + "，实际为" + (actualRows ? actualRows.lastRow : "无内容") + "。");
        }
        for (var sampleIndex in expectedSamples) {
            if (!expectedSamples.hasOwnProperty(sampleIndex)) continue;
            var index = Number(sampleIndex);
            var verifyRow = contentRows.firstRow + index;
            if (rawRowSignature(sheet, verifyRow, bounds) !== expectedSamples[sampleIndex]) {
                throw new Error("第" + verifyRow + "行内容抽样核验失败。");
            }
        }
        for (var trailing = expectedLastRow + 1; trailing <= contentRows.lastRow; trailing += 1) {
            if (rowHasAnyContent(sheet, trailing, bounds)) throw new Error("尾部第" + trailing + "行未清空。");
        }
        return { deleted: validRows.length, deletedRows: validRows.slice(), failedRows: failedRows, method: "rewrite-block" };
    }


    function runStableRewriteSelfTest(workbook) {
        var originalSheetName = getOriginalActiveSheetName(workbook);
        var testSheet = null;
        var oldAlerts = Application.DisplayAlerts;
        try {
            workbook.Activate();
            testSheet = workbook.Worksheets.Add();
            testSheet.Name = makeTemporarySheetName(workbook, "RW");
            testSheet.Activate();
            testSheet.Range("A1:F8").Clear();
            try { testSheet.Range("A1:F8").UnMerge(); } catch (ignored0) {}
            testSheet.Range("A1:F1").Merge();
            testSheet.Range("A1").Value2 = "稳定重写自检";
            testSheet.Range("A2").Value2 = "头部";
            testSheet.Range("B2:C2").Merge();
            testSheet.Range("B2").Value2 = "保留一";
            testSheet.Range("A3:F4").Clear();
            testSheet.Range("A5:C5").Merge();
            testSheet.Range("D5:F5").Merge();
            testSheet.Range("A5").Value2 = "合计";
            testSheet.Range("D5").Value2 = 12345;
            testSheet.Range("A6").Value2 = "尾部";
            testSheet.Range("B6:C6").Merge();
            testSheet.Range("B6").Value2 = "保留二";
            testSheet.Range("A7:A8").Merge();
            testSheet.Range("A7").Value2 = "纵向合并";
            testSheet.Range("B7").Value2 = "删除本行";
            testSheet.Range("B8").Value2 = "保留三";

            var result = rewriteRowsOnActiveSheet([3, 4, 7]);
            if (result.deleted !== 3) throw new Error("预计移除3行，实际为" + result.deleted + "行。");
            if (normalizeText(testSheet.Range("A3").Value2) !== "合计") throw new Error("合计行没有上移到第3行。");
            if (normalizeText(testSheet.Range("D3").Value2) !== "12345") throw new Error("合计金额没有上移。");
            if (normalizeText(testSheet.Range("A4").Value2) !== "尾部") throw new Error("尾部行没有上移到第4行。");
            if (!testSheet.Range("A3:C3").MergeCells || !testSheet.Range("D3:F3").MergeCells) throw new Error("横向合并结构没有正确重建。");
            if (normalizeText(testSheet.Range("A5").Value2) !== "纵向合并" || normalizeText(testSheet.Range("B5").Value2) !== "保留三") throw new Error("纵向合并区域缩短后内容没有正确保留。");
            if (testSheet.Range("A5").MergeCells) throw new Error("纵向合并缩短为单格后仍处于合并状态。");
            return true;
        } catch (error) {
            throw new Error("稳定重写引擎自检失败，原表尚未处理：" + (error && error.message ? error.message : String(error)));
        } finally {
            try { Application.DisplayAlerts = false; } catch (ignored1) {}
            if (testSheet) {
                try { testSheet.Delete(); }
                catch (deleteError) {
                    try { testSheet.Name = "__重写自检残留"; testSheet.Visible = false; } catch (ignored2) {}
                }
            }
            try { Application.DisplayAlerts = oldAlerts; } catch (ignored3) {}
            restoreActiveSheet(workbook, originalSheetName);
        }
    }

    function deleteRowsOnSheet(workbook, sheetName, rows, preferredColumns, deletionMethod) {
        activateWorksheet(workbook, sheetName);
        return rewriteRowsOnActiveSheet(rows);
    }



    function updateProgress(label, current, total) {
        var now = Number(current);
        var count = Number(total);
        if (
            count > 8 &&
            now !== 1 &&
            now !== count &&
            now % 5 !== 0
        ) {
            return;
        }

        try {
            Application.StatusBar =
                "工程表清理助手｜" +
                label + "（" +
                now + "/" + count + "）";
        } catch (ignored) {}
    }


    function isAssistantInternalSheetName(name) {
        var text = String(name || "");
        return text === PREVIEW_SHEET_NAME ||
            text === DUPLICATE_SHEET_NAME ||
            text === UNDO_META_SHEET_NAME ||
            text.indexOf(UNDO_BACKUP_PREFIX) === 0 ||
            text.indexOf("__重写自检") === 0;
    }

    function worksheetIndexByName(workbook, sheetName) {
        for (var i = 1; i <= workbook.Worksheets.Count; i += 1) {
            if (String(workbook.Worksheets.Item(i).Name) === String(sheetName)) return i;
        }
        return 0;
    }

    function clearUndoPointInternal(workbook) {
        var oldAlerts = Application.DisplayAlerts;
        try { Application.DisplayAlerts = false; } catch (ignored0) {}
        try {
            for (var i = workbook.Worksheets.Count; i >= 1; i -= 1) {
                var sheet = workbook.Worksheets.Item(i);
                var name = String(sheet.Name);
                if (name === UNDO_META_SHEET_NAME || name.indexOf(UNDO_BACKUP_PREFIX) === 0) {
                    try { sheet.Visible = true; } catch (ignored1) {}
                    try { sheet.Delete(); } catch (deleteError) {}
                }
            }
        } finally {
            try { Application.DisplayAlerts = oldAlerts; } catch (ignored2) {}
        }
    }

    function makeUndoBackupName(workbook, number) {
        var base = UNDO_BACKUP_PREFIX + ("000" + number).slice(-3);
        var name = base;
        var suffix = 1;
        while (findWorksheet(workbook, name)) {
            suffix += 1;
            name = base.substring(0, 25) + "_" + suffix;
        }
        return name;
    }

    function createUndoPoint(workbook, sheetNames, actionLabel) {
        var originalSheetName = getOriginalActiveSheetName(workbook);
        var names = [];
        var seen = {};
        for (var i = 0; i < sheetNames.length; i += 1) {
            var name = String(sheetNames[i]);
            if (!name || seen[name] || isAssistantInternalSheetName(name)) continue;
            if (!findWorksheet(workbook, name)) continue;
            seen[name] = true;
            names.push(name);
        }
        if (!names.length) return 0;

        clearUndoPointInternal(workbook);
        var meta = workbook.Worksheets.Add();
        meta.Name = UNDO_META_SHEET_NAME;
        var metaHeaders = ["原工作表", "备份工作表", "原索引", "原可见性", "操作", "时间"];
        for (var headerIndex = 0; headerIndex < metaHeaders.length; headerIndex += 1) {
            meta.Cells.Item(1, headerIndex + 1).Value2 = metaHeaders[headerIndex];
        }
        var created = [];
        try {
            for (var n = 0; n < names.length; n += 1) {
                updateProgress("建立撤回点", n + 1, names.length);
                var source = findWorksheet(workbook, names[n]);
                if (!source) continue;
                var originalIndex = worksheetIndexByName(workbook, names[n]);
                var originalVisible = true;
                try { originalVisible = source.Visible; } catch (ignored0) {}
                source.Copy(null, workbook.Worksheets.Item(workbook.Worksheets.Count));
                var backup = Application.ActiveSheet;
                var backupName = makeUndoBackupName(workbook, n + 1);
                backup.Name = backupName;
                var row = n + 2;
                meta.Cells.Item(row, 1).Value2 = names[n];
                meta.Cells.Item(row, 2).Value2 = backupName;
                meta.Cells.Item(row, 3).Value2 = originalIndex;
                meta.Cells.Item(row, 4).Value2 = originalVisible;
                meta.Cells.Item(row, 5).Value2 = actionLabel;
                meta.Cells.Item(row, 6).Value2 = new Date().toLocaleString();
                created.push(backupName);
                try { backup.Visible = false; } catch (ignored1) {}
            }
            try { meta.Visible = false; } catch (ignored2) {}
            restoreActiveSheet(workbook, originalSheetName);
            return created.length;
        } catch (error) {
            clearUndoPointInternal(workbook);
            restoreActiveSheet(workbook, originalSheetName);
            throw new Error("撤回点创建失败，原操作尚未开始：" + (error && error.message ? error.message : String(error)));
        }
    }

    function undoLastOperation() {
        var context = getTargetContext(true);
        var workbook = context.workbook;
        var meta = findWorksheet(workbook, UNDO_META_SHEET_NAME);
        if (!meta) {
            MsgBox("当前工作簿没有可撤回的插件操作。", JS_INFORMATION, "没有撤回点");
            return;
        }
        var bounds = getUsedBounds(meta);
        var records = [];
        for (var row = 2; row <= bounds.lastRow; row += 1) {
            var originalName = normalizeText(meta.Cells.Item(row, 1).Value2);
            var backupName = normalizeText(meta.Cells.Item(row, 2).Value2);
            if (!originalName || !backupName) continue;
            records.push({
                originalName: originalName,
                backupName: backupName,
                originalIndex: Number(meta.Cells.Item(row, 3).Value2),
                originalVisible: meta.Cells.Item(row, 4).Value2,
                actionLabel: normalizeText(meta.Cells.Item(row, 5).Value2)
            });
        }
        if (!records.length) {
            clearUndoPointInternal(workbook);
            MsgBox("撤回点没有有效数据。", JS_EXCLAMATION, "无法撤回");
            return;
        }
        var label = records[0].actionLabel || "上一步插件操作";
        if (MsgBox("将撤回：“" + label + "”。\n涉及" + records.length + "张工作表。\n\n是否继续？", JS_YES_NO + JS_QUESTION, "撤回上一步") !== JS_RESULT_YES) return;

        var oldAlerts = Application.DisplayAlerts;
        var restored = 0;
        var failures = [];
        try { Application.DisplayAlerts = false; } catch (ignored0) {}
        try {
            records.sort(function (a, b) { return a.originalIndex - b.originalIndex; });
            for (var i = 0; i < records.length; i += 1) {
                updateProgress("撤回上一步", i + 1, records.length);
                var record = records[i];
                try {
                    var backup = findWorksheet(workbook, record.backupName);
                    if (!backup) throw new Error("找不到备份工作表。");
                    try { backup.Visible = true; } catch (ignored1) {}
                    var original = findWorksheet(workbook, record.originalName);
                    if (original) original.Delete();
                    backup.Name = record.originalName;
                    try { backup.Visible = record.originalVisible; } catch (ignored2) {}
                    try {
                        var before = null;
                        if (record.originalIndex >= 1 && record.originalIndex <= workbook.Worksheets.Count) {
                            before = workbook.Worksheets.Item(record.originalIndex);
                        }
                        if (before && String(before.Name) !== record.originalName) backup.Move(before, null);
                    } catch (moveError) {}
                    restored += 1;
                } catch (restoreError) {
                    failures.push("“" + record.originalName + "”：" + (restoreError && restoreError.message ? restoreError.message : String(restoreError)));
                }
            }
            try { meta.Delete(); } catch (metaDeleteError) {}
        } finally {
            try { Application.DisplayAlerts = oldAlerts; } catch (ignored3) {}
        }
        var message = "已撤回并恢复" + restored + "张工作表。";
        if (failures.length) {
            message += "\n\n未恢复：\n" + failures.slice(0, 12).join("\n");
            MsgBox(message, JS_EXCLAMATION, "部分撤回");
        } else {
            MsgBox(message, JS_INFORMATION, "撤回完成");
        }
    }

    function clearUndoPointCommand() {
        var context = getTargetContext(true);
        var workbook = context.workbook;
        if (!findWorksheet(workbook, UNDO_META_SHEET_NAME)) {
            MsgBox("当前工作簿没有撤回点。", JS_INFORMATION, "无需清理");
            return;
        }
        if (MsgBox("将永久清除插件的一步撤回备份。\n是否继续？", JS_YES_NO + JS_QUESTION, "清除撤回点") !== JS_RESULT_YES) return;
        clearUndoPointInternal(workbook);
        MsgBox("撤回点已清除。", JS_INFORMATION, "清理完成");
    }

    function rowTextItems(sheet, rowNumber, bounds) {
        var items = [];
        for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
            var cell = sheet.Cells.Item(rowNumber, column);
            if (!isTopLeftOfMerge(cell)) continue;
            var text = normalizeText(readMergedAwareValue(cell));
            if (text) items.push({ column: column, text: text, cell: cell });
        }
        return items;
    }

    function likelyRepeatedTitle(
        sheet,
        rowNumber,
        bounds
    ) {
        var items =
            rowTextItems(
                sheet,
                rowNumber,
                bounds
            );
        var tableWidth =
            bounds.lastColumn -
            bounds.firstColumn + 1;

        for (
            var i = 0;
            i < items.length;
            i += 1
        ) {
            var text = items[i].text;

            if (
                /第\s*\d+\s*页/.test(text) ||
                /(工程名称|项目名称)/.test(text)
            ) {
                continue;
            }

            var mergeWidth = 1;
            var mergeHeight = 1;
            try {
                if (items[i].cell.MergeCells) {
                    mergeWidth = Number(
                        items[i].cell
                            .MergeArea
                            .Columns.Count
                    );
                    mergeHeight = Number(
                        items[i].cell
                            .MergeArea
                            .Rows.Count
                    );
                }
            } catch (ignored0) {}

            /*
             * 结构规则：
             * 宽合并且有文字的行，优先视为表名。
             * 因此换成其他表格名称也能识别，不依赖具体名称。
             */
            if (
                text.length >= 4 &&
                mergeWidth >= Math.max(
                    3,
                    Math.floor(
                        tableWidth * 0.45
                    )
                ) &&
                mergeHeight <= 3
            ) {
                return text;
            }

            /*
             * 非合并或合并信息读取失败时保留兼容判断。
             */
            if (
                (
                    /(表|清单)$/.test(text) ||
                    /(汇总表|概预算表|明细表|费用表|计算表)/.test(text)
                ) &&
                (
                    mergeWidth >= 3 ||
                    text.length >= 6
                )
            ) {
                return text;
            }
        }
        return "";
    }


    function rowHeaderKeywordScore(sheet, rowNumber, bounds) {
        var text = rowTextItems(sheet, rowNumber, bounds).map(function (item) { return item.text; }).join("|");
        var keywords = ["序号", "编号", "编码", "名称", "单位", "数量", "单价", "合计", "工程量", "价值", "人工费", "材料费", "报价金额", "备注", "项目名称", "工程名称"];
        var score = 0;
        for (var i = 0; i < keywords.length; i += 1) if (text.indexOf(keywords[i]) >= 0) score += 1;
        return score;
    }

    function rowLooksLikeData(sheet, rowNumber, bounds) {
        var items = rowTextItems(sheet, rowNumber, bounds);
        if (!items.length) return false;
        var joined = items.map(function (item) { return item.text; }).join("|");
        if (/第\s*\d+\s*页/.test(joined) || /(工程名称|项目名称)/.test(joined)) return false;
        if (rowHeaderKeywordScore(sheet, rowNumber, bounds) >= 2) return false;
        var numeric = 0;
        var serial = false;
        for (var i = 0; i < items.length; i += 1) {
            var text = items[i].text.replace(/,/g, "");
            if (/^-?\d+(?:\.\d+)?$/.test(text)) numeric += 1;
            if (i < 3 && /^\d+(?:\.\d+)*$/.test(text)) serial = true;
        }
        return (serial && items.length >= 2) || numeric >= 2;
    }

    function findEarlierSameTitle(
        sheet,
        title,
        beforeRow,
        bounds
    ) {
        var key = duplicateContentKey(title);
        for (
            var row = bounds.firstRow;
            row < beforeRow;
            row += 1
        ) {
            if (
                duplicateContentKey(
                    likelyRepeatedTitle(
                        sheet,
                        row,
                        bounds
                    )
                ) === key
            ) {
                return row;
            }
        }
        return 0;
    }

    function rowLooksLikeStrongData(
        sheet,
        rowNumber,
        bounds
    ) {
        if (
            detectChineseMajorRow(
                sheet,
                rowNumber,
                bounds
            )
        ) {
            return true;
        }

        if (
            rowLooksLikeData(
                sheet,
                rowNumber,
                bounds
            )
        ) {
            return true;
        }

        var items = rowTextItems(
            sheet,
            rowNumber,
            bounds
        );
        if (!items.length) return false;

        for (
            var i = 0;
            i < items.length && i < 3;
            i += 1
        ) {
            if (
                /^\d+(?:\.\d+)*$/.test(
                    items[i].text
                ) &&
                items.length >= 2
            ) {
                return true;
            }
        }
        return false;
    }

    function normalizedRowContentKey(
        sheet,
        rowNumber,
        bounds
    ) {
        return rowTextItems(
            sheet,
            rowNumber,
            bounds
        ).map(function (item) {
            return duplicateContentKey(
                item.text
            );
        }).join("|");
    }

    function rowIsRepeatableContinuationMarker(
        sheet,
        rowNumber,
        bounds
    ) {
        var items =
            rowTextItems(
                sheet,
                rowNumber,
                bounds
            );
        if (!items.length) return false;

        if (
            detectChineseMajorRow(
                sheet,
                rowNumber,
                bounds
            )
        ) {
            return true;
        }

        /*
         * 普通续页分类行通常文字较少，且没有两项以上数字金额。
         */
        if (items.length > 3) return false;

        var numeric = 0;
        for (
            var i = 0;
            i < items.length;
            i += 1
        ) {
            if (
                /^-?\d+(?:\.\d+)?$/.test(
                    items[i].text
                        .replace(/,/g, "")
                )
            ) {
                numeric += 1;
            }
        }
        return numeric < 2;
    }

    function hasEarlierSameRowContent(
        sheet,
        rowNumber,
        beforeRow,
        bounds
    ) {
        var key =
            normalizedRowContentKey(
                sheet,
                rowNumber,
                bounds
            );
        if (!key) return false;

        for (
            var row = bounds.firstRow;
            row < beforeRow;
            row += 1
        ) {
            if (
                normalizedRowContentKey(
                    sheet,
                    row,
                    bounds
                ) === key
            ) {
                return true;
            }
        }
        return false;
    }

    function detectRepeatedHeaderBlock(
        sheet,
        candidateRow,
        title,
        bounds,
        contentRows
    ) {
        if (
            !findEarlierSameTitle(
                sheet,
                title,
                candidateRow,
                bounds
            )
        ) {
            return null;
        }

        var endRow = candidateRow;
        var sawContext = false;
        var headerScore = 0;
        var maxEnd = Math.min(
            contentRows.lastRow,
            candidateRow + 14
        );

        for (
            var row = candidateRow + 1;
            row <= maxEnd;
            row += 1
        ) {
            if (
                rowLooksLikeStrongData(
                    sheet,
                    row,
                    bounds
                )
            ) {
                break;
            }

            var items =
                rowTextItems(
                    sheet,
                    row,
                    bounds
                );
            var joined = items.map(
                function (item) {
                    return item.text;
                }
            ).join("|");
            var context =
                /第\s*\d+\s*页/.test(joined) ||
                /(工程名称|项目名称)/.test(joined);
            var score =
                rowHeaderKeywordScore(
                    sheet,
                    row,
                    bounds
                );

            if (context) sawContext = true;
            headerScore += score;

            if (
                !items.length ||
                context ||
                score > 0
            ) {
                endRow = row;
                continue;
            }

            if (
                headerScore > 0 &&
                row <= candidateRow + 7
            ) {
                endRow = row;
                continue;
            }
            break;
        }

        if (
            !sawContext &&
            headerScore < 1
        ) {
            return null;
        }

        /*
         * 页首常重复“材料类别、机械类别、主材类别”等续页标记。
         * 如果表头后的一至两行与前页某行完全相同，则一并删除后面的重复项，
         * 避免清理后出现两条相同分类行。
         */
        for (
            var extra = endRow + 1;
            extra <= Math.min(
                contentRows.lastRow,
                endRow + 2
            );
            extra += 1
        ) {
            if (
                !rowIsRepeatableContinuationMarker(
                    sheet,
                    extra,
                    bounds
                ) ||
                !hasEarlierSameRowContent(
                    sheet,
                    extra,
                    candidateRow,
                    bounds
                )
            ) {
                break;
            }
            endRow = extra;
        }

        return {
            startRow: candidateRow,
            endRow: endRow,
            title: title
        };
    }


    function findRepeatedPaginationHeaderBlocks(sheet) {
        var bounds = getUsedBounds(sheet);
        var contentRows =
            getActualContentRowBounds(
                sheet,
                bounds
            );
        if (!contentRows) return [];

        var firstByTitle = {};
        var blocks = [];

        for (
            var row = contentRows.firstRow;
            row <= contentRows.lastRow;
            row += 1
        ) {
            var title = likelyRepeatedTitle(
                sheet,
                row,
                bounds
            );
            if (!title) continue;

            var key = duplicateContentKey(title);
            if (!firstByTitle.hasOwnProperty(key)) {
                firstByTitle[key] = row;
                continue;
            }

            var block = detectRepeatedHeaderBlock(
                sheet,
                row,
                title,
                bounds,
                contentRows
            );
            if (!block) continue;

            if (
                blocks.length &&
                block.startRow <=
                    blocks[blocks.length - 1].endRow
            ) {
                continue;
            }

            blocks.push(block);
            row = Math.max(row, block.endRow);
        }
        return blocks;
    }

    function flattenHeaderBlocks(blocks) {
        var rows = [];
        var seen = {};
        for (
            var i = 0;
            i < blocks.length;
            i += 1
        ) {
            for (
                var row = blocks[i].startRow;
                row <= blocks[i].endRow;
                row += 1
            ) {
                if (!seen[row]) {
                    seen[row] = true;
                    rows.push(row);
                }
            }
        }
        rows.sort(function (a, b) {
            return a - b;
        });
        return rows;
    }

    function detectRepeatedPaginationHeaderRows(sheet) {
        return flattenHeaderBlocks(
            findRepeatedPaginationHeaderBlocks(
                sheet
            )
        );
    }

    function countStrongDataRowsAfterBlock(
        sheet,
        block,
        bounds,
        contentRows
    ) {
        var count = 0;
        for (
            var row = block.endRow + 1;
            row <= contentRows.lastRow;
            row += 1
        ) {
            if (
                rowLooksLikeStrongData(
                    sheet,
                    row,
                    bounds
                )
            ) {
                count += 1;
            }
        }
        return count;
    }

    function detectOrphanHeaderRows(sheet) {
        var bounds = getUsedBounds(sheet);
        var contentRows =
            getActualContentRowBounds(
                sheet,
                bounds
            );
        if (!contentRows) return [];

        var blocks =
            findRepeatedPaginationHeaderBlocks(
                sheet
            );
        if (!blocks.length) return [];

        /*
         * 自动模式只考虑最后一个重复表头。
         * 中间分页表头无论文字是否相同，均不自动删除。
         */
        var block =
            blocks[blocks.length - 1];
        var headerRows =
            block.endRow -
            block.startRow + 1;
        var tailDataRows =
            countStrongDataRowsAfterBlock(
                sheet,
                block,
                bounds,
                contentRows
            );

        /*
         * 仅当尾页内容很少时自动合并。
         * 阈值同时考虑固定安全值和表头自身高度。
         */
        var safeLimit = Math.max(
            12,
            headerRows * 3
        );
        if (
            tailDataRows <= 0 ||
            tailDataRows > safeLimit
        ) {
            return [];
        }

        var rows = [];
        for (
            var row = block.startRow;
            row <= block.endRow;
            row += 1
        ) {
            rows.push(row);
        }
        return rows;
    }


    function cleanupOrphanHeadersOnActiveSheet() {
        var sheet = Application.ActiveSheet;
        var rows =
            detectOrphanHeaderRows(sheet);
        if (!rows.length) return 0;

        var result =
            rewriteRowsOnActiveSheet(rows);
        return Number(
            result.deleted || 0
        );
    }



    function clearStaleTailOnSheet(sheet) {
        var used = getUsedBounds(sheet);
        var actual =
            getActualContentBounds(sheet);
        if (!actual) return 0;

        var trailing =
            getTrailingBlankTemplateInfo(
                sheet,
                used,
                actual
            );
        if (!trailing.count) return 0;

        var startRow =
            actual.lastRow + 1;
        var lastRow =
            trailing.lastRow;
        var firstColumn = Math.min(
            used.firstColumn,
            actual.firstColumn
        );
        var lastColumn = Math.max(
            used.lastColumn,
            actual.lastColumn
        );

        var address =
            columnToLetters(firstColumn) +
            startRow + ":" +
            columnToLetters(lastColumn) +
            lastRow;
        var area = sheet.Range(address);

        /*
         * 空白模板区中的合并区域会让 Clear 保留合并结构，
         * 导致下次仍被判定为可见模板。因此先拆分再清理。
         */
        try {
            area.UnMerge();
        } catch (ignored0) {}

        try {
            area.Clear();
        } catch (clearError) {
            try {
                area.ClearContents();
            } catch (ignored1) {}
            try {
                area.ClearFormats();
            } catch (ignored2) {}
        }

        /*
         * 对部分 WPS 版本再显式清除边框和填充。
         */
        try {
            area.Borders.LineStyle = -4142;
        } catch (ignored3) {}
        try {
            area.Interior.Pattern = -4142;
        } catch (ignored4) {}

        for (
            var row = startRow;
            row <= lastRow;
            row += 1
        ) {
            try {
                sheet.Rows.Item(row).Hidden =
                    false;
            } catch (ignored5) {}
            try {
                sheet.Rows.Item(row).RowHeight =
                    15;
            } catch (ignored6) {}
        }

        return trailing.count;
    }


    function parseSinglePrintArea(value) {
        var text = normalizeText(value);
        if (!text) return null;
        var match = text.match(/\$?([A-Z]+)\$?(\d+)\s*:\s*\$?([A-Z]+)\$?(\d+)/i);
        if (!match) return null;
        return {
            firstColumn: lettersToColumn(match[1]),
            firstRow: Number(match[2]),
            lastColumn: lettersToColumn(match[3]),
            lastRow: Number(match[4])
        };
    }

    function getLogicalPrintBottomRow(
        sheet,
        actual
    ) {
        /*
         * 只有已经实际删除的行才会改变底部。
         * 不再因为“疑似表头”提前缩短打印区域。
         */
        return Number(actual.lastRow);
    }


    function addHorizontalBreakBeforeRow(
        sheet,
        rowNumber,
        firstColumn
    ) {
        var row = Number(rowNumber);
        var column = Number(firstColumn);
        if (!isFinite(row) || row < 2) return false;
        if (!isFinite(column) || column < 1) column = 1;

        try {
            sheet.HPageBreaks.Add(
                sheet.Cells.Item(row, column)
            );
            return true;
        } catch (ignored0) {}

        try {
            sheet.Rows.Item(row).PageBreak =
                -4135;
            return true;
        } catch (ignored1) {}

        try {
            sheet.Cells.Item(
                row,
                column
            ).PageBreak = -4135;
            return true;
        } catch (ignored2) {}
        return false;
    }

    function removeManualBreaksInsideHeaderBlock(
        sheet,
        block
    ) {
        try {
            var breaks = sheet.HPageBreaks;
            for (
                var i = Number(breaks.Count);
                i >= 1;
                i -= 1
            ) {
                var item = breaks.Item(i);
                var locationRow = Number(
                    item.Location.Row
                );
                if (
                    locationRow > block.startRow &&
                    locationRow <= block.endRow + 1
                ) {
                    try {
                        item.Delete();
                    } catch (ignored0) {}
                }
            }
        } catch (ignored1) {}
    }

    function alignInternalPageBreaksToTables(
        sheet,
        actual
    ) {
        var blocks =
            findRepeatedPaginationHeaderBlocks(
                sheet
            );
        var added = 0;

        for (
            var i = 0;
            i < blocks.length;
            i += 1
        ) {
            if (
                blocks[i].startRow <=
                actual.firstRow
            ) {
                continue;
            }

            removeManualBreaksInsideHeaderBlock(
                sheet,
                blocks[i]
            );

            if (
                addHorizontalBreakBeforeRow(
                    sheet,
                    blocks[i].startRow,
                    actual.firstColumn
                )
            ) {
                added += 1;
            }
        }
        return added;
    }

    function getVisibleTableBounds(
        sheet,
        actual
    ) {
        var used = getUsedBounds(sheet);
        var result = {
            firstRow: actual.firstRow,
            lastRow: actual.lastRow,
            firstColumn: actual.firstColumn,
            lastColumn: actual.lastColumn
        };

        var firstRow = Math.max(
            used.firstRow,
            actual.firstRow - 4
        );
        var lastRow = Math.min(
            used.lastRow,
            actual.lastRow + 500
        );
        var firstColumn = Math.max(
            used.firstColumn,
            actual.firstColumn - 2
        );
        var lastColumn = Math.min(
            used.lastColumn,
            actual.lastColumn + 8
        );

        var scanCells =
            (lastRow - firstRow + 1) *
            (lastColumn - firstColumn + 1);
        if (
            scanCells >
            MAX_BORDER_SCAN_CELLS
        ) {
            return result;
        }

        for (
            var row = firstRow;
            row <= lastRow;
            row += 1
        ) {
            for (
                var column = firstColumn;
                column <= lastColumn;
                column += 1
            ) {
                var cell =
                    sheet.Cells.Item(
                        row,
                        column
                    );

                if (
                    !cellHasAnyVisibleBorder(
                        cell
                    )
                ) {
                    continue;
                }

                if (row < result.firstRow) {
                    result.firstRow = row;
                }
                if (row > result.lastRow) {
                    result.lastRow = row;
                }
                if (
                    column <
                    result.firstColumn
                ) {
                    result.firstColumn =
                        column;
                }
                if (
                    column >
                    result.lastColumn
                ) {
                    result.lastColumn =
                        column;
                }
            }
        }

        return result;
    }

    function getBorderWeightRankFromRef(
        borderRef
    ) {
        if (!borderRef) return 0;
        try {
            return borderWeightRank(
                Number(
                    borderRef.border.Weight
                )
            );
        } catch (ignored) {
            return 0;
        }
    }

    function rowHasStrongBottomBoundary(
        sheet,
        rowNumber,
        bounds
    ) {
        var span =
            getRowVisibleBorderSpan(
                sheet,
                rowNumber,
                bounds
            );
        if (!span.hasBorder) {
            return false;
        }

        var width =
            span.lastColumn -
            span.firstColumn + 1;
        var strongCount = 0;
        var visibleCount = 0;
        var leftStrong = false;
        var rightStrong = false;

        for (
            var column =
                span.firstColumn;
            column <=
                span.lastColumn;
            column += 1
        ) {
            var cell =
                sheet.Cells.Item(
                    rowNumber,
                    column
                );
            var bottom =
                getVisibleBorder(
                    cell,
                    9
                );
            if (!bottom) continue;

            visibleCount += 1;
            var isStrong =
                getBorderWeightRankFromRef(
                    bottom
                ) >= 3;

            if (isStrong) {
                strongCount += 1;
                if (
                    column ===
                    span.firstColumn
                ) {
                    leftStrong = true;
                }
                if (
                    column ===
                    span.lastColumn
                ) {
                    rightStrong = true;
                }
            }
        }

        if (!visibleCount) return false;

        return (
            strongCount >=
                Math.max(
                    2,
                    Math.ceil(width * 0.62)
                ) ||
            (
                leftStrong &&
                rightStrong &&
                strongCount >=
                    Math.max(
                        2,
                        Math.ceil(
                            width * 0.42
                        )
                    )
            )
        );
    }

    function rowHasContentOrBorder(
        sheet,
        rowNumber,
        bounds
    ) {
        if (
            rowNumber < bounds.firstRow ||
            rowNumber > bounds.lastRow
        ) {
            return false;
        }

        for (
            var column = bounds.firstColumn;
            column <= bounds.lastColumn;
            column += 1
        ) {
            var cell =
                sheet.Cells.Item(
                    rowNumber,
                    column
                );

            if (
                cellHasFormula(cell) ||
                !isBlankValue(
                    readMergedAwareValue(
                        cell
                    )
                ) ||
                cellHasAnyVisibleBorder(
                    cell
                )
            ) {
                return true;
            }
        }
        return false;
    }

    function detectDesiredPageBreakRows(
        sheet,
        bounds
    ) {
        var result = [];
        var seen = {};
        var titleRows = [];

        /*
         * 分页优先按后续宽合并表名定位。表名行即新页起点，
         * 不再把表内小计粗线、分类粗线误判为分页线。
         */
        for (
            var row = bounds.firstRow;
            row <= bounds.lastRow;
            row += 1
        ) {
            if (
                rowIsWideTableTitle(
                    sheet,
                    row,
                    bounds
                )
            ) {
                titleRows.push(row);
            }
        }

        if (titleRows.length >= 2) {
            for (
                var t = 1;
                t < titleRows.length;
                t += 1
            ) {
                if (!seen[titleRows[t]]) {
                    seen[titleRows[t]] = true;
                    result.push(titleRows[t]);
                }
            }
        } else {
            var blocks =
                detectBorderTableBlocks(
                    sheet,
                    bounds
                );
            for (
                var i = 1;
                i < blocks.length;
                i += 1
            ) {
                var breakRow =
                    Number(
                        blocks[i].firstRow
                    );
                if (
                    isFinite(breakRow) &&
                    breakRow > bounds.firstRow &&
                    breakRow <= bounds.lastRow &&
                    !seen[breakRow]
                ) {
                    seen[breakRow] = true;
                    result.push(breakRow);
                }
            }
        }

        result.sort(function (a, b) {
            return a - b;
        });
        if (result.length > 1025) {
            result = result.slice(0, 1025);
        }
        return result;
    }


    function clearAllPageBreaks(
        sheet
    ) {
        try {
            sheet.ResetAllPageBreaks();
            return true;
        } catch (ignored0) {}

        try {
            var horizontal =
                sheet.HPageBreaks;
            for (
                var h =
                    Number(
                        horizontal.Count
                    );
                h >= 1;
                h -= 1
            ) {
                try {
                    horizontal
                        .Item(h)
                        .Delete();
                } catch (ignored1) {}
            }
        } catch (ignored2) {}

        try {
            var vertical =
                sheet.VPageBreaks;
            for (
                var v =
                    Number(
                        vertical.Count
                    );
                v >= 1;
                v -= 1
            ) {
                try {
                    vertical
                        .Item(v)
                        .Delete();
                } catch (ignored3) {}
            }
        } catch (ignored4) {}

        return false;
    }

    function refreshPageBreakDisplay(
        sheet
    ) {
        try {
            sheet.DisplayPageBreaks =
                false;
        } catch (ignored0) {}
        try {
            sheet.DisplayPageBreaks =
                true;
        } catch (ignored1) {}
    }

    function rebuildPrintLayoutOnSheet(
        sheet
    ) {
        var actual =
            getActualContentBounds(
                sheet
            );
        if (!actual) {
            return {
                address: "",
                breakRows: [],
                bottomRow: 0,
                firstColumn: 0,
                lastColumn: 0,
                tableBlocks: 0
            };
        }

        var bounds =
            getVisibleTableBounds(
                sheet,
                actual
            );
        var blocks =
            detectBorderTableBlocks(
                sheet,
                bounds
            );

        if (blocks.length) {
            bounds.firstRow =
                Math.min(
                    bounds.firstRow,
                    blocks[0].firstRow
                );
            bounds.lastRow =
                Math.max(
                    actual.lastRow,
                    blocks[
                        blocks.length - 1
                    ].lastRow
                );

            for (
                var b = 0;
                b < blocks.length;
                b += 1
            ) {
                bounds.firstColumn =
                    Math.min(
                        bounds.firstColumn,
                        blocks[b].firstColumn
                    );
                bounds.lastColumn =
                    Math.max(
                        bounds.lastColumn,
                        blocks[b].lastColumn
                    );
            }
        }

        var address =
            "$" +
            columnToLetters(
                bounds.firstColumn
            ) +
            "$" + bounds.firstRow +
            ":$" +
            columnToLetters(
                bounds.lastColumn
            ) +
            "$" + bounds.lastRow;

        sheet.PageSetup.PrintArea =
            address;

        /*
         * 不强制FitToPagesWide=1。强制按页数缩放会让WPS持续
         * 重算自动分页，分页预览中的蓝线难以手动拖动。
         */
        var breakRows =
            detectDesiredPageBreakRows(
                sheet,
                bounds
            );

        clearAllPageBreaks(sheet);

        var added = [];
        for (
            var i = 0;
            i < breakRows.length;
            i += 1
        ) {
            if (
                addHorizontalBreakBeforeRow(
                    sheet,
                    breakRows[i],
                    bounds.firstColumn
                )
            ) {
                added.push(breakRows[i]);
            }
        }

        refreshPageBreakDisplay(sheet);

        return {
            address: address,
            breakRows: added,
            bottomRow: bounds.lastRow,
            firstColumn: bounds.firstColumn,
            lastColumn: bounds.lastColumn,
            tableBlocks: blocks.length
        };
    }


    function alignPrintBottomOnSheet(
        sheet
    ) {
        var actual =
            getActualContentBounds(
                sheet
            );
        if (!actual) return "";

        var bounds =
            getVisibleTableBounds(
                sheet,
                actual
            );
        var blocks =
            detectBorderTableBlocks(
                sheet,
                bounds
            );

        if (blocks.length) {
            bounds.firstRow =
                Math.min(
                    bounds.firstRow,
                    blocks[0].firstRow
                );
            bounds.lastRow =
                Math.max(
                    actual.lastRow,
                    blocks[
                        blocks.length - 1
                    ].lastRow
                );
            for (
                var i = 0;
                i < blocks.length;
                i += 1
            ) {
                bounds.firstColumn =
                    Math.min(
                        bounds.firstColumn,
                        blocks[i].firstColumn
                    );
                bounds.lastColumn =
                    Math.max(
                        bounds.lastColumn,
                        blocks[i].lastColumn
                    );
            }
        }

        var address =
            "$" +
            columnToLetters(
                bounds.firstColumn
            ) +
            "$" + bounds.firstRow +
            ":$" +
            columnToLetters(
                bounds.lastColumn
            ) +
            "$" + bounds.lastRow;

        /*
         * 普通删行/删列后只收缩PrintArea，不清除分页符。
         * 用户手动拖动好的分页线会被保留。
         */
        sheet.PageSetup.PrintArea =
            address;
        refreshPageBreakDisplay(sheet);
        return address;
    }


    function postProcessAfterRowChange(sheet) {
        var result = {
            orphanRows: 0,
            tailRows: 0,
            printArea: "",
            warnings: []
        };

        /*
         * 表头清理属于有语义的独立操作，不能作为所有删行功能的自动后处理。
         * 否则换用其他表格或正常多页表时可能误删中间页表头。
         */
        try {
            result.tailRows =
                clearStaleTailOnSheet(sheet);
        } catch (tailError) {
            result.warnings.push(
                "清理尾部格式失败：" +
                (
                    tailError &&
                    tailError.message
                        ? tailError.message
                        : String(tailError)
                )
            );
        }

        try {
            result.printArea =
                alignPrintBottomOnSheet(sheet);
        } catch (printError) {
            result.warnings.push(
                "调整打印底边失败：" +
                (
                    printError &&
                    printError.message
                        ? printError.message
                        : String(printError)
                )
            );
        }
        return result;
    }


    function alignTableBottomByScope(mode) {
        var context =
            getTargetContext(false);
        var workbook =
            context.workbook;
        var originalSheetName =
            getOriginalActiveSheetName(
                workbook
            );
        var sheetNames =
            getScopeSheetNames(
                workbook,
                context.sheet,
                mode
            );

        if (
            MsgBox(
                "将重新建立打印布局：" +
                "\n1. PrintArea按实际内容和连续表格块重设；" +
                "\n2. 清除旧的手动水平、垂直分页符；" +
                "\n3. 在后续宽合并表名行之前重新放置水平分页线；" +
                "\n4. 不强制一页宽，保留现有缩放方式。" +
                "\n\n处理范围：" +
                (
                    mode === "workbook"
                        ? "整个工作簿（" +
                          sheetNames.length +
                          "张）"
                        : "当前工作表"
                ) +
                "\n\n是否继续？",
                JS_YES_NO + JS_QUESTION,
                "重建打印范围与分页线"
            ) !== JS_RESULT_YES
        ) {
            return;
        }

        var completed = 0;
        var breakCount = 0;
        var blockCount = 0;
        var failures = [];
        var examples = [];

        for (
            var i = 0;
            i < sheetNames.length;
            i += 1
        ) {
            updateProgress(
                "重建打印范围与分页线",
                i + 1,
                sheetNames.length
            );

            try {
                var sheet =
                    activateWorksheet(
                        workbook,
                        sheetNames[i]
                    );
                clearStaleTailOnSheet(
                    sheet
                );
                var result =
                    rebuildPrintLayoutOnSheet(
                        sheet
                    );

                breakCount +=
                    result.breakRows.length;
                blockCount +=
                    Number(
                        result.tableBlocks || 0
                    );
                completed += 1;

                if (examples.length < 8) {
                    examples.push(
                        "“" + sheetNames[i] +
                        "”：表格块" +
                        result.tableBlocks +
                        "个，底行" +
                        result.bottomRow +
                        "，分页线" +
                        result.breakRows.length +
                        "条"
                    );
                }
            } catch (error) {
                failures.push(
                    "“" + sheetNames[i] +
                    "”：" +
                    (
                        error &&
                        error.message
                            ? error.message
                            : String(error)
                    )
                );
            }
        }

        restoreActiveSheet(
            workbook,
            originalSheetName
        );

        var message =
            "已重建" + completed +
            "张工作表，识别" +
            blockCount +
            "个连续表格块，重新放置水平分页线" +
            breakCount + "条。" +
            "\n没有改变现有打印缩放方式。";

        if (examples.length) {
            message +=
                "\n\n" +
                examples.join("\n");
        }

        if (failures.length) {
            message +=
                "\n\n失败：\n" +
                failures
                    .slice(0, 12)
                    .join("\n");
            MsgBox(
                message,
                JS_EXCLAMATION,
                completed
                    ? "部分完成"
                    : "处理失败"
            );
        } else {
            MsgBox(
                message,
                JS_INFORMATION,
                "打印布局重建完成"
            );
        }
    }


    function getSelectedRowSpanForSheet(sheet) {
        try {
            if (
                String(Application.ActiveSheet.Name) !==
                String(sheet.Name)
            ) {
                return null;
            }
            var selection = Application.Selection;
            if (!selection) return null;
            var firstRow = Number(selection.Row);
            var rowCount = Number(selection.Rows.Count);
            if (
                !isFinite(firstRow) ||
                !isFinite(rowCount) ||
                rowCount < 1
            ) {
                return null;
            }
            return {
                firstRow: firstRow,
                lastRow: firstRow + rowCount - 1,
                rowCount: rowCount
            };
        } catch (ignored) {
            return null;
        }
    }

    function rowIsHeaderLike(
        sheet,
        rowNumber,
        bounds
    ) {
        if (
            likelyRepeatedTitle(
                sheet,
                rowNumber,
                bounds
            )
        ) {
            return true;
        }

        var items = rowTextItems(
            sheet,
            rowNumber,
            bounds
        );
        if (!items.length) return true;

        var joined = items.map(
            function (item) {
                return item.text;
            }
        ).join("|");

        if (
            /第\s*\d+\s*页/.test(joined) ||
            /(工程名称|项目名称)/.test(joined)
        ) {
            return true;
        }

        return (
            rowHeaderKeywordScore(
                sheet,
                rowNumber,
                bounds
            ) >= 1
        );
    }

    function detectSelectedHeaderRows(sheet) {
        var span =
            getSelectedRowSpanForSheet(sheet);
        if (!span) return [];

        var bounds = getUsedBounds(sheet);
        var contentRows =
            getActualContentRowBounds(
                sheet,
                bounds
            );
        if (!contentRows) return [];

        var firstRow = Math.max(
            span.firstRow,
            contentRows.firstRow
        );
        var lastRow = Math.min(
            span.lastRow,
            contentRows.lastRow
        );
        if (firstRow > lastRow) return [];

        if (span.rowCount > 1) {
            var exactRows = [];
            for (
                var exact = firstRow;
                exact <= lastRow;
                exact += 1
            ) {
                exactRows.push(exact);
            }
            return exactRows;
        }

        var startRow = firstRow;
        var searchTop = Math.max(
            contentRows.firstRow,
            firstRow - 5
        );
        for (
            var up = firstRow;
            up >= searchTop;
            up -= 1
        ) {
            if (
                likelyRepeatedTitle(
                    sheet,
                    up,
                    bounds
                )
            ) {
                startRow = up;
                break;
            }
        }

        if (
            !rowIsHeaderLike(
                sheet,
                firstRow,
                bounds
            ) &&
            startRow === firstRow
        ) {
            return [];
        }

        var maxEnd = Math.min(
            contentRows.lastRow,
            startRow + 14
        );
        var endRow = startRow;
        var sawHeader = false;

        for (
            var row = startRow;
            row <= maxEnd;
            row += 1
        ) {
            if (
                row > startRow &&
                rowLooksLikeData(
                    sheet,
                    row,
                    bounds
                )
            ) {
                break;
            }

            if (
                rowIsHeaderLike(
                    sheet,
                    row,
                    bounds
                )
            ) {
                sawHeader = true;
                endRow = row;
                continue;
            }

            var items = rowTextItems(
                sheet,
                row,
                bounds
            );
            if (!items.length) {
                endRow = row;
                continue;
            }

            if (sawHeader) break;
            endRow = row;
        }

        var rows = [];
        for (
            var removeRow = startRow;
            removeRow <= endRow;
            removeRow += 1
        ) {
            rows.push(removeRow);
        }
        return rows;
    }

    function describeRows(rows) {
        if (!rows || !rows.length) return "";
        if (rows.length === 1) {
            return "第" + rows[0] + "行";
        }
        return (
            "第" + rows[0] +
            "—" +
            rows[rows.length - 1] +
            "行"
        );
    }

    function cleanupOrphanHeadersByScope(mode) {
        var context = getTargetContext(false);
        var workbook = context.workbook;
        var originalSheetName =
            getOriginalActiveSheetName(workbook);
        var sheetNames = getScopeSheetNames(
            workbook,
            context.sheet,
            mode
        );
        var plans = [];
        var total = 0;
        var manualSelection = false;

        if (mode === "sheet") {
            var selectedRows =
                detectSelectedHeaderRows(
                    context.sheet
                );
            if (selectedRows.length) {
                plans.push({
                    sheetName:
                        String(context.sheet.Name),
                    rows: selectedRows,
                    manual: true
                });
                total = selectedRows.length;
                manualSelection = true;
            }
        }

        if (!manualSelection) {
            for (
                var i = 0;
                i < sheetNames.length;
                i += 1
            ) {
                try {
                    var sheet = activateWorksheet(
                        workbook,
                        sheetNames[i]
                    );
                    var rows =
                        detectOrphanHeaderRows(
                            sheet
                        );
                    if (rows.length) {
                        plans.push({
                            sheetName:
                                sheetNames[i],
                            rows: rows,
                            manual: false
                        });
                        total += rows.length;
                    }
                } catch (ignored) {}
            }
        }

        restoreActiveSheet(
            workbook,
            originalSheetName
        );

        if (!total) {
            MsgBox(
                mode === "sheet"
                    ? "没有识别到可安全合并的短尾页表头。" +
                      "\n可选中表名行中的任意单元格，" +
                      "或直接选中需要删除的多行，" +
                      "再点击【当前工作表／按选择】。"
                    : "没有检测到可安全合并的短尾页表头。",
                JS_INFORMATION,
                "无需清理"
            );
            return;
        }

        var confirmText =
            manualSelection
                ? "已按当前选择识别：" +
                  describeRows(plans[0].rows) +
                  "。\n将删除这些行并把后续内容上移。" +
                  "\n\n是否继续？"
                : "检测到" +
                  plans.length +
                  "张工作表存在可安全合并的短尾页表头，共" +
                  total +
                  "行。\n\n是否删除并合并到上一页？";

        if (
            MsgBox(
                confirmText,
                JS_YES_NO + JS_QUESTION,
                "清理重复分页表头"
            ) !== JS_RESULT_YES
        ) {
            return;
        }

        createUndoPoint(
            workbook,
            plans.map(function (item) {
                return item.sheetName;
            }),
            "清理重复分页表头"
        );

        var deleted = 0;
        var failures = [];

        for (
            var p = 0;
            p < plans.length;
            p += 1
        ) {
            updateProgress(
                "清理重复分页表头",
                p + 1,
                plans.length
            );
            try {
                var active = activateWorksheet(
                    workbook,
                    plans[p].sheetName
                );
                var result =
                    rewriteRowsOnActiveSheet(
                        plans[p].rows
                    );
                deleted += result.deleted;
                postProcessAfterRowChange(active);

                for (
                    var f = 0;
                    f < result.failedRows.length;
                    f += 1
                ) {
                    failures.push(
                        "“" +
                        plans[p].sheetName +
                        "”第" +
                        result.failedRows[f].row +
                        "行：" +
                        result.failedRows[f].message
                    );
                }
            } catch (error) {
                failures.push(
                    "“" +
                    plans[p].sheetName +
                    "”：" +
                    (
                        error && error.message
                            ? error.message
                            : String(error)
                    )
                );
            }
        }

        restoreActiveSheet(
            workbook,
            originalSheetName
        );

        var message =
            "已清理重复分页表头" +
            deleted +
            "行。";
        if (failures.length) {
            message +=
                "\n失败：" +
                failures.length +
                "项。";
            MsgBox(
                message,
                JS_EXCLAMATION,
                deleted
                    ? "部分完成"
                    : "处理失败"
            );
        } else {
            MsgBox(
                message,
                JS_INFORMATION,
                "清理完成"
            );
        }
    }

    function sheetNameMatches(name, tokens) {
        var lower = String(name).toLowerCase();
        for (var i = 0; i < tokens.length; i += 1) {
            var token = tokens[i];
            if (token.charAt(0) === "=") {
                if (lower === token.substring(1).toLowerCase()) return true;
            } else if (lower.indexOf(token.toLowerCase()) >= 0) return true;
        }
        return false;
    }

    function promptBatchDeleteWorksheets() {
        var context = getTargetContext(false);
        var workbook = context.workbook;
        var originalSheetName = getOriginalActiveSheetName(workbook);
        var input = normalizeText(InputBox("请输入要删除的工作表名称关键词。\n多个关键词用逗号分隔；名称中包含关键词即匹配。\n精确匹配请在名称前加=。", "批量删除工作表", ""));
        if (!input) return;
        var tokens = splitKeywords(input);
        if (!tokens.length) return;
        var matched = [];
        for (var i = 1; i <= workbook.Worksheets.Count; i += 1) {
            var sheet = workbook.Worksheets.Item(i);
            var name = String(sheet.Name);
            if (isAssistantInternalSheetName(name)) continue;
            if (sheetNameMatches(name, tokens)) matched.push(name);
        }
        if (!matched.length) {
            MsgBox("没有匹配的工作表。", JS_INFORMATION, "没有匹配项");
            return;
        }
        var remaining = 0;
        for (var j = 1; j <= workbook.Worksheets.Count; j += 1) if (!isAssistantInternalSheetName(workbook.Worksheets.Item(j).Name)) remaining += 1;
        if (matched.length >= remaining) throw new Error("不能删除工作簿中的全部普通工作表。");
        var preview = matched.slice(0, 20).join("\n");
        if (matched.length > 20) preview += "\n……另有" + (matched.length - 20) + "张";
        if (MsgBox("将删除" + matched.length + "张工作表：\n\n" + preview + "\n\n是否继续？", JS_YES_NO + JS_QUESTION, "批量删除工作表") !== JS_RESULT_YES) return;
        createUndoPoint(workbook, matched, "批量删除工作表");
        var oldAlerts = Application.DisplayAlerts;
        var deleted = 0;
        var failures = [];
        try { Application.DisplayAlerts = false; } catch (ignored0) {}
        try {
            for (var d = 0; d < matched.length; d += 1) {
                updateProgress("批量删除工作表", d + 1, matched.length);
                try {
                    var target = findWorksheet(workbook, matched[d]);
                    if (target) { target.Delete(); deleted += 1; }
                } catch (error) {
                    failures.push("“" + matched[d] + "”：" + (error && error.message ? error.message : String(error)));
                }
            }
        } finally {
            try { Application.DisplayAlerts = oldAlerts; } catch (ignored1) {}
        }
        if (findWorksheet(workbook, originalSheetName)) restoreActiveSheet(workbook, originalSheetName);
        else {
            for (var k = 1; k <= workbook.Worksheets.Count; k += 1) {
                if (!isAssistantInternalSheetName(workbook.Worksheets.Item(k).Name)) { workbook.Worksheets.Item(k).Activate(); break; }
            }
        }
        var message = "已删除" + deleted + "张工作表。";
        if (failures.length) {
            message += "\n\n失败：\n" + failures.slice(0, 12).join("\n");
            MsgBox(message, JS_EXCLAMATION, "部分完成");
        } else MsgBox(message, JS_INFORMATION, "删除完成");
    }

    function findWorksheet(workbook, sheetName) {
        for (var i = 1; i <= workbook.Worksheets.Count; i += 1) {
            var sheet = workbook.Worksheets.Item(i);
            if (String(sheet.Name) === String(sheetName)) return sheet;
        }
        return null;
    }

    function getScopeSheets(workbook, activeSheet, mode) {
        if (mode === "sheet") return [activeSheet];
        var sheets = [];
        for (var i = 1; i <= workbook.Worksheets.Count; i += 1) {
            var sheet = workbook.Worksheets.Item(i);
            if (String(sheet.Name) === PREVIEW_SHEET_NAME) continue;
            sheets.push(sheet);
        }
        return sheets;
    }

    function getScopeSheetNames(workbook, activeSheet, mode) {
        if (mode === "sheet") return [String(activeSheet.Name)];
        var names = [];
        for (var i = 1; i <= workbook.Worksheets.Count; i += 1) {
            var sheet = workbook.Worksheets.Item(i);
            var name = String(sheet.Name);
            if (isAssistantInternalSheetName(name)) continue;
            names.push(name);
        }
        return names;
    }

    function promptScope() {
        var answer = InputBox(
            "请选择处理范围：\n\n1  当前工作表\n2  整个工作簿\n\n请输入1或2：",
            "选择处理范围",
            "2"
        );
        answer = normalizeText(answer);
        if (answer === "1") return "sheet";
        if (answer === "2") return "workbook";
        return "";
    }

    function splitKeywords(text) {
        var normalized = normalizeText(text);
        if (!normalized) return [];
        var raw = normalized.split(/[，,；;、|]+/);
        var result = [];
        var seen = {};
        for (var i = 0; i < raw.length; i += 1) {
            var item = normalizeText(raw[i]);
            if (!item || seen[item.toLowerCase()]) continue;
            seen[item.toLowerCase()] = true;
            result.push(item);
        }
        return result;
    }

    function getOriginalActiveSheetName(workbook) {
        try {
            if (Application.ActiveWorkbook === workbook && Application.ActiveSheet) return String(Application.ActiveSheet.Name);
        } catch (e) {}
        return "";
    }

    function restoreActiveSheet(workbook, sheetName) {
        if (!sheetName) return;
        try {
            var sheet = findWorksheet(workbook, sheetName);
            if (sheet) sheet.Activate();
        } catch (e) {}
    }

    function findHeaderCells(sheet, keyword, bounds) {
        var normalizedKeyword = normalizeText(keyword).toLowerCase();
        if (!normalizedKeyword) return [];
        var matches = [];
        var maxHeaderRow = bounds.lastRow;
        for (var row = bounds.firstRow; row <= maxHeaderRow; row += 1) {
            for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
                var text = normalizeText(readMergedAwareValue(sheet.Cells.Item(row, column)));
                if (text && text.toLowerCase().indexOf(normalizedKeyword) >= 0) {
                    matches.push({ row: row, column: column, text: text, address: columnToLetters(column) + row });
                    if (matches.length >= 80) return matches;
                }
            }
        }
        return matches;
    }

    function chooseHeaderInteractive(matches, keyword, sheetName) {
        if (matches.length === 1) return matches[0];
        var lines = [];
        var displayCount = Math.min(matches.length, 30);
        for (var i = 0; i < displayCount; i += 1) {
            lines.push((i + 1) + ". " + matches[i].address + "  " + matches[i].text);
        }
        if (matches.length > displayCount) lines.push("……另有 " + (matches.length - displayCount) + " 个匹配未列出");
        var answer = InputBox(
            "工作表：“" + sheetName + "”\n找到多个包含“" + keyword + "”的单元格：\n\n" +
            lines.join("\n") + "\n\n请输入要使用的序号（1-" + displayCount + "）：",
            "选择判断列",
            "1"
        );
        var index = Number(answer);
        if (!isFinite(index) || Math.floor(index) !== index || index < 1 || index > displayCount) return null;
        return matches[index - 1];
    }

    function chooseHeaderAutomatically(matches, keyword, bounds) {
        if (matches.length === 0) return null;
        var key = normalizeText(keyword).toLowerCase();
        var best = matches[0];
        var bestScore = -999999;
        for (var i = 0; i < matches.length; i += 1) {
            var item = matches[i];
            var text = normalizeText(item.text).toLowerCase();
            var score = 0;
            if (text === key) score += 1000;
            if (text.indexOf(key) === 0) score += 400;
            if (text.indexOf("金额") >= 0) score += 120;
            if (text.indexOf("费用") >= 0) score += 40;
            score += Math.max(0, 80 - (item.row - bounds.firstRow));
            if (item.column > best.column) score += 1;
            if (score > bestScore) {
                bestScore = score;
                best = item;
            }
        }
        return best;
    }

    function deletePreviewSheetIfExists(workbook, askFirst) {
        var preview = findWorksheet(workbook, PREVIEW_SHEET_NAME);
        if (!preview) return true;
        if (askFirst) {
            var replace = MsgBox(
                "当前工作簿已经存在“" + PREVIEW_SHEET_NAME + "”。\n是否删除旧预览并重新生成？",
                JS_YES_NO + JS_QUESTION,
                "替换删除预览"
            );
            if (replace !== JS_RESULT_YES) return false;
        }
        var oldAlerts = Application.DisplayAlerts;
        try {
            Application.DisplayAlerts = false;
            preview.Delete();
        } finally {
            Application.DisplayAlerts = oldAlerts;
        }
        return true;
    }

    function createPreviewSheet(workbook, scopeLabel, keyword, candidates, selectedHeaders, protectionSettings, protectionStats) {
        if (!deletePreviewSheetIfExists(workbook, true)) return null;
        var preview = workbook.Worksheets.Add();
        preview.Name = PREVIEW_SHEET_NAME;
        preview.Range("A1:K1").Merge();
        preview.Range("A1").Value2 = "工程表清理助手——指定空白列删除预览";
        preview.Range("A1").Font.Bold = true;
        preview.Range("A1").Font.Size = 16;
        preview.Range("A1").HorizontalAlignment = -4108;

        preview.Range("A2").Value2 = "处理范围";
        preview.Range("B2").Value2 = scopeLabel;
        preview.Range("A3").Value2 = "查找关键词";
        preview.Range("B3").Value2 = keyword;
        preview.Range("A4").Value2 = "智能保护";
        preview.Range("B4:K4").Merge();
        preview.Range("B4").Value2 = protectionSettings;
        preview.Range("B4").WrapText = true;
        preview.Range("A5").Value2 = "判断列";
        preview.Range("B5:K5").Merge();
        preview.Range("B5").Value2 = selectedHeaders.join("；");
        preview.Range("B5").WrapText = true;
        preview.Range("A6").Value2 = "默认结果";
        preview.Range("B6:K6").Merge();
        preview.Range("B6").Value2 = "默认删除" + protectionStats.deleteCount + "行；默认保留" + protectionStats.keepCount + "行。保留原因会逐行显示，A列仍可下拉改成“删除/保留”。";
        preview.Range("B6").WrapText = true;
        preview.Range("A7").Value2 = "操作方法";
        preview.Range("B7:K7").Merge();
        preview.Range("B7").Value2 = "确认A列后点击【执行选择表删除】。执行前会重新核对并自检删行接口；完成后本选择表会自动关闭并返回原表。";
        preview.Range("B7").WrapText = true;
        preview.Range("A8").Value2 = "重要提醒";
        preview.Range("B8:K8").Merge();
        preview.Range("B8").Value2 = "删除通常不能使用Ctrl+Z恢复。请先保存原文件副本。";
        preview.Range("B8").WrapText = true;

        var headers = ["操作", "工作表", "原行号", "判断列", "判断列值", "该行内容摘要", "保护原因", "原表位置", "行校验码", "执行状态", "判断列号"];
        for (var h = 0; h < headers.length; h += 1) preview.Cells.Item(9, h + 1).Value2 = headers[h];
        preview.Range("A9:K9").Font.Bold = true;
        preview.Range("A9:K9").HorizontalAlignment = -4108;

        for (var i = 0; i < candidates.length; i += 1) {
            var outputRow = PREVIEW_DATA_START_ROW + i;
            var item = candidates[i];
            preview.Cells.Item(outputRow, 1).Value2 = item.defaultAction;
            preview.Cells.Item(outputRow, 2).Value2 = item.sheetName;
            preview.Cells.Item(outputRow, 3).Value2 = item.row;
            preview.Cells.Item(outputRow, 4).Value2 = item.headerAddress + " / " + item.headerText;
            preview.Cells.Item(outputRow, 5).Value2 = "（空白）";
            preview.Cells.Item(outputRow, 6).Value2 = item.summary;
            preview.Cells.Item(outputRow, 7).Value2 = item.protectionReason || "";
            preview.Cells.Item(outputRow, 8).Value2 = item.sheetName + "!" + columnToLetters(item.headerColumn) + item.row;
            preview.Cells.Item(outputRow, 9).Value2 = item.signature;
            preview.Cells.Item(outputRow, 10).Value2 = "待执行";
            preview.Cells.Item(outputRow, 11).Value2 = item.headerColumn;
        }

        var lastOutputRow = PREVIEW_DATA_START_ROW + candidates.length - 1;
        if (candidates.length > 0) {
            preview.Range("A9:K" + lastOutputRow).Borders.LineStyle = 1;
            preview.Range("A9:K" + lastOutputRow).VerticalAlignment = -4108;
            preview.Range("F10:G" + lastOutputRow).WrapText = true;
            preview.Range("A10:A" + lastOutputRow).Font.Bold = true;
            try {
                preview.Range("A10:A" + lastOutputRow).Validation.Delete();
                preview.Range("A10:A" + lastOutputRow).Validation.Add(3, 1, 1, "删除,保留");
                preview.Range("A10:A" + lastOutputRow).Validation.IgnoreBlank = true;
                preview.Range("A10:A" + lastOutputRow).Validation.InCellDropdown = true;
            } catch (validationError) {}
            try { preview.Range("A9:K" + lastOutputRow).AutoFilter(); } catch (filterError) {}
        }

        preview.Columns.Item(1).ColumnWidth = 11;
        preview.Columns.Item(2).ColumnWidth = 30;
        preview.Columns.Item(3).ColumnWidth = 10;
        preview.Columns.Item(4).ColumnWidth = 25;
        preview.Columns.Item(5).ColumnWidth = 12;
        preview.Columns.Item(6).ColumnWidth = 62;
        preview.Columns.Item(7).ColumnWidth = 34;
        preview.Columns.Item(8).ColumnWidth = 36;
        preview.Columns.Item(10).ColumnWidth = 20;
        preview.Columns.Item(9).Hidden = true;
        preview.Columns.Item(11).Hidden = true;
        preview.Rows.Item(4).RowHeight = 42;
        preview.Rows.Item(5).RowHeight = 42;
        preview.Rows.Item(6).RowHeight = 34;
        preview.Rows.Item(7).RowHeight = 34;

        preview.Activate();
        preview.Range("A10").Select();
        try { Application.ActiveWindow.FreezePanes = true; } catch (freezeError) {}
        return preview;
    }

    function isDeleteMark(value) {
        var text = normalizeText(value).toLowerCase();
        return text === "删除" || text === "是" || text === "y" || text === "yes" || text === "1" || text === "√" || text === "勾选";
    }

    function setPreviewMarks(markText) {
        var context = getTargetContext(true);
        var preview = findWorksheet(context.workbook, PREVIEW_SHEET_NAME);
        if (!preview) {
            MsgBox("当前工作簿没有删除预览。", JS_EXCLAMATION, "找不到预览");
            return;
        }
        var bounds = getUsedBounds(preview);
        var changed = 0;
        for (var row = PREVIEW_DATA_START_ROW; row <= bounds.lastRow; row += 1) {
            var status = normalizeText(preview.Cells.Item(row, 10).Value2);
            if (status && status !== "待执行") continue;
            preview.Cells.Item(row, 1).Value2 = markText;
            changed += 1;
        }
        MsgBox("已将" + changed + "个候选行设置为“" + markText + "”。", JS_INFORMATION, "预览选择");
    }

    function locatePreviewRow() {
        var context = getTargetContext(true);
        var workbook = context.workbook;
        var preview = findWorksheet(workbook, PREVIEW_SHEET_NAME);
        if (!preview || String(Application.ActiveSheet.Name) !== PREVIEW_SHEET_NAME) {
            MsgBox("请先切换到“" + PREVIEW_SHEET_NAME + "”，选中一条候选记录。", JS_EXCLAMATION, "定位原行");
            return;
        }
        var row = Number(Application.ActiveCell.Row);
        if (row < PREVIEW_DATA_START_ROW) {
            MsgBox("请选中预览表中的候选记录行。", JS_EXCLAMATION, "定位原行");
            return;
        }
        var sheetName = normalizeText(preview.Cells.Item(row, 2).Value2);
        var originalRow = Number(preview.Cells.Item(row, 3).Value2);
        var headerColumn = Number(preview.Cells.Item(row, 11).Value2);
        var sheet = findWorksheet(workbook, sheetName);
        if (!sheet || !isFinite(originalRow) || originalRow < 1) {
            MsgBox("无法定位该记录对应的原工作表行。", JS_EXCLAMATION, "定位失败");
            return;
        }
        sheet.Activate();
        if (isFinite(headerColumn) && headerColumn >= 1) sheet.Cells.Item(originalRow, headerColumn).Select();
        else sheet.Rows.Item(originalRow).Select();
    }


    function isVisibleBorderLineStyle(value) {
        if (
            value === null ||
            value === undefined ||
            value === ""
        ) {
            return false;
        }

        var number = Number(value);
        if (
            isFinite(number) &&
            (
                number === 0 ||
                number === -4142
            )
        ) {
            return false;
        }
        return true;
    }

    function cellHasVisibleTailFormatting(cell) {
        try {
            if (cell.MergeCells) return true;
        } catch (ignored0) {}

        try {
            var aggregateLineStyle =
                cell.Borders.LineStyle;
            if (
                isVisibleBorderLineStyle(
                    aggregateLineStyle
                )
            ) {
                return true;
            }
        } catch (ignored1) {}

        /*
         * WPS 的 Borders.LineStyle 在混合边框时可能返回空值，
         * 因此逐项检查四条外边及内部边框。
         */
        var borderIndexes = [
            7, 8, 9, 10, 11, 12
        ];
        for (
            var i = 0;
            i < borderIndexes.length;
            i += 1
        ) {
            try {
                var lineStyle =
                    cell.Borders.Item(
                        borderIndexes[i]
                    ).LineStyle;
                if (
                    isVisibleBorderLineStyle(
                        lineStyle
                    )
                ) {
                    return true;
                }
            } catch (ignored2) {}
        }

        try {
            var pattern =
                Number(cell.Interior.Pattern);
            if (
                isFinite(pattern) &&
                pattern !== 0 &&
                pattern !== -4142
            ) {
                return true;
            }
        } catch (ignored3) {}

        try {
            var colorIndex =
                Number(cell.Interior.ColorIndex);
            if (
                isFinite(colorIndex) &&
                colorIndex !== -4142 &&
                colorIndex !== -4105 &&
                colorIndex !== 0
            ) {
                return true;
            }
        } catch (ignored4) {}

        return false;
    }

    function rowHasVisibleTailFormatting(
        sheet,
        rowNumber,
        firstColumn,
        lastColumn
    ) {
        for (
            var column = firstColumn;
            column <= lastColumn;
            column += 1
        ) {
            if (
                cellHasVisibleTailFormatting(
                    sheet.Cells.Item(
                        rowNumber,
                        column
                    )
                )
            ) {
                return true;
            }
        }
        return false;
    }

    function getTrailingBlankTemplateInfo(
        sheet,
        bounds,
        actualBounds
    ) {
        if (!actualBounds) {
            return {
                firstRow: 0,
                lastRow: 0,
                count: 0,
                usedTailCount: 0,
                printTailCount: 0,
                visibleTailCount: 0
            };
        }

        var contentLastRow =
            Number(actualBounds.lastRow);
        var printArea = null;
        var printTailEnd = contentLastRow;

        try {
            printArea = parseSinglePrintArea(
                sheet.PageSetup.PrintArea
            );
            if (
                printArea &&
                isFinite(printArea.lastRow)
            ) {
                printTailEnd =
                    Number(printArea.lastRow);
            }
        } catch (ignored0) {}

        var maxCandidateRow = Math.max(
            Number(bounds.lastRow),
            printTailEnd,
            contentLastRow
        );

        /*
         * 防止异常 UsedRange 或打印区域指向整张工作表，
         * 一次最多检查实际内容下方五千行。
         */
        maxCandidateRow = Math.min(
            maxCandidateRow,
            contentLastRow + 5000
        );

        var firstScanColumn =
            Number(actualBounds.firstColumn);
        var lastScanColumn =
            Number(actualBounds.lastColumn);

        if (printArea) {
            if (
                isFinite(printArea.firstColumn)
            ) {
                firstScanColumn = Math.min(
                    firstScanColumn,
                    Number(
                        printArea.firstColumn
                    )
                );
            }
            if (
                isFinite(printArea.lastColumn)
            ) {
                lastScanColumn = Math.max(
                    lastScanColumn,
                    Number(
                        printArea.lastColumn
                    )
                );
            }
        }

        var lastVisibleTemplateRow =
            contentLastRow;
        var visibleTailCount = 0;

        for (
            var row = contentLastRow + 1;
            row <= maxCandidateRow;
            row += 1
        ) {
            if (
                rowHasVisibleTailFormatting(
                    sheet,
                    row,
                    firstScanColumn,
                    lastScanColumn
                )
            ) {
                lastVisibleTemplateRow = row;
                visibleTailCount += 1;
            }
        }

        /*
         * 打印区域底边仍在实际内容下方时，即使单元格已无边框，
         * 也仍需把蓝色底线收回。
         */
        var effectivePrintTailEnd =
            printTailEnd > contentLastRow
                ? Math.min(
                    printTailEnd,
                    maxCandidateRow
                )
                : contentLastRow;

        var lastTemplateRow = Math.max(
            lastVisibleTemplateRow,
            effectivePrintTailEnd
        );

        /*
         * 关键修复：
         * 不再因为 UsedRange.lastRow 仍然很大，就把全部尾行再次计入。
         * UsedRange 只决定扫描上限；是否需要清理，由当前可见格式和
         * 当前打印区域共同决定。
         */
        return {
            firstRow:
                lastTemplateRow >
                contentLastRow
                    ? contentLastRow + 1
                    : 0,
            lastRow: lastTemplateRow,
            count: Math.max(
                0,
                lastTemplateRow -
                    contentLastRow
            ),
            usedTailCount: Math.max(
                0,
                Number(bounds.lastRow) -
                    contentLastRow
            ),
            printTailCount: Math.max(
                0,
                effectivePrintTailEnd -
                    contentLastRow
            ),
            visibleTailCount:
                visibleTailCount
        };
    }

    function buildBlankRowCleanupPlan(sheet) {
        var bounds = getUsedBounds(sheet);
        var actualBounds =
            getActualContentBounds(sheet);

        if (!actualBounds) {
            return {
                bounds: bounds,
                actualBounds: null,
                contentRows: null,
                internalRows: [],
                trailing: {
                    firstRow: 0,
                    lastRow: 0,
                    count: 0,
                    usedTailCount: 0,
                    printTailCount: 0,
                    visibleTailCount: 0
                },
                total: 0
            };
        }

        var contentRows = {
            firstRow:
                actualBounds.firstRow,
            lastRow:
                actualBounds.lastRow
        };

        /*
         * UsedRange 可能因旧格式扩大到很多列。
         * 内部空行检测只检查实际表格所在列，避免大表重复扫描空列。
         */
        var contentColumnBounds = {
            firstRow: bounds.firstRow,
            lastRow: bounds.lastRow,
            firstColumn:
                actualBounds.firstColumn,
            lastColumn:
                actualBounds.lastColumn
        };

        var internalRows = [];
        for (
            var row = contentRows.firstRow;
            row <= contentRows.lastRow;
            row += 1
        ) {
            if (
                inspectRow(
                    sheet,
                    row,
                    contentColumnBounds
                ).isBlank &&
                !rowIntersectsVerticalMerge(
                    sheet,
                    row,
                    contentColumnBounds
                )
            ) {
                internalRows.push(row);
            }
        }

        var trailing =
            getTrailingBlankTemplateInfo(
                sheet,
                bounds,
                actualBounds
            );

        return {
            bounds: bounds,
            actualBounds: actualBounds,
            contentRows: contentRows,
            internalRows: internalRows,
            trailing: trailing,
            total:
                internalRows.length +
                trailing.count
        };
    }


    function deleteBlankRowsByScope(mode) {
        var context = getTargetContext(false);
        var workbook = context.workbook;
        var originalSheetName =
            getOriginalActiveSheetName(workbook);
        var sheetNames = getScopeSheetNames(
            workbook,
            context.sheet,
            mode
        );

        var scanPlans = [];
        var internalTotal = 0;
        var trailingTotal = 0;
        var scanFailures = [];

        for (
            var i = 0;
            i < sheetNames.length;
            i += 1
        ) {
            updateProgress(
                "扫描完全空白行",
                i + 1,
                sheetNames.length
            );
            try {
                var sheet = activateWorksheet(
                    workbook,
                    sheetNames[i]
                );
                var plan =
                    buildBlankRowCleanupPlan(
                        sheet
                    );

                if (plan.total > 0) {
                    scanPlans.push({
                        sheetName:
                            sheetNames[i],
                        internalCount:
                            plan.internalRows.length,
                        trailingCount:
                            plan.trailing.count
                    });
                    internalTotal +=
                        plan.internalRows.length;
                    trailingTotal +=
                        plan.trailing.count;
                }
            } catch (error) {
                scanFailures.push(
                    "“" + sheetNames[i] +
                    "”：" +
                    (
                        error &&
                        error.message
                            ? error.message
                            : String(error)
                    )
                );
            }
        }

        restoreActiveSheet(
            workbook,
            originalSheetName
        );

        var total =
            internalTotal + trailingTotal;
        if (!total) {
            MsgBox(
                "所选范围没有需要清理的完全空白行。" +
                (
                    scanFailures.length
                        ? "\n\n未检查：\n" +
                          scanFailures
                              .slice(0, 12)
                              .join("\n")
                        : ""
                ),
                scanFailures.length
                    ? JS_EXCLAMATION
                    : JS_INFORMATION,
                "删除完全空白行"
            );
            return;
        }

        var detail =
            "内部断行：" +
            internalTotal + "行" +
            "\n尾部空白模板行：" +
            trailingTotal + "行";

        if (
            MsgBox(
                "处理范围：" +
                (
                    mode === "workbook"
                        ? "整个工作簿"
                        : "当前工作表"
                ) +
                "\n涉及" +
                scanPlans.length +
                "张工作表。" +
                "\n" + detail +
                "\n\n尾部空白行会清除边框、格式，" +
                "并把底部蓝线收回最后实际内容行。" +
                "\n\n是否继续？",
                JS_YES_NO + JS_QUESTION,
                "删除完全空白行"
            ) !== JS_RESULT_YES
        ) {
            return;
        }

        createUndoPoint(
            workbook,
            scanPlans.map(function (item) {
                return item.sheetName;
            }),
            "删除完全空白行"
        );

        /*
         * 只有存在内部断行时才需要执行重写引擎自检。
         * 单纯清理尾部空白模板行只需 Clear + 收缩 PrintArea。
         */
        if (internalTotal > 0) {
            runStableRewriteSelfTest(
                workbook
            );
        }

        var internalDeleted = 0;
        var tailCleared = 0;
        var orphanDeleted = 0;
        var failures =
            scanFailures.slice();

        for (
            var d = 0;
            d < scanPlans.length;
            d += 1
        ) {
            var sheetName =
                scanPlans[d].sheetName;
            updateProgress(
                "删除完全空白行",
                d + 1,
                scanPlans.length
            );

            try {
                var activeSheet =
                    activateWorksheet(
                        workbook,
                        sheetName
                    );
                var currentPlan =
                    buildBlankRowCleanupPlan(
                        activeSheet
                    );

                if (
                    currentPlan.internalRows
                        .length
                ) {
                    var result =
                        rewriteRowsOnActiveSheet(
                            currentPlan
                                .internalRows
                        );
                    internalDeleted +=
                        result.deleted;

                    for (
                        var f = 0;
                        f <
                            result.failedRows
                                .length;
                        f += 1
                    ) {
                        failures.push(
                            "“" + sheetName +
                            "”第" +
                            result.failedRows[f]
                                .row +
                            "行：" +
                            result.failedRows[f]
                                .message
                        );
                    }
                }

                /*
                 * 无论是否存在内部断行，都执行尾部清理。
                 * 这是 v1.3.0 漏掉截图中第26—33行的根本修复。
                 */
                var post =
                    postProcessAfterRowChange(
                        activeSheet
                    );
                orphanDeleted +=
                    Number(
                        post.orphanRows || 0
                    );
                tailCleared +=
                    Number(
                        post.tailRows || 0
                    );
            } catch (error2) {
                failures.push(
                    "“" + sheetName +
                    "”：" +
                    (
                        error2 &&
                        error2.message
                            ? error2.message
                            : String(error2)
                    )
                );
            }
        }

        restoreActiveSheet(
            workbook,
            originalSheetName
        );

        var message =
            "内部空白断行实际上移：" +
            internalDeleted + "行。" +
            "\n尾部空白模板实际清理：" +
            tailCleared + "行。";

        if (orphanDeleted) {
            message +=
                "\n另清理尾页残留表头：" +
                orphanDeleted + "行。";
        }

        if (failures.length) {
            message +=
                "\n\n未完成：\n" +
                failures
                    .slice(0, 12)
                    .join("\n");
            MsgBox(
                message,
                JS_EXCLAMATION,
                internalDeleted ||
                tailCleared
                    ? "部分完成"
                    : "处理失败"
            );
        } else {
            MsgBox(
                message,
                JS_INFORMATION,
                "清理完成"
            );
        }
    }


    function findBestHeaderNearRow(sheet, keywords, bounds, preferredRow, excludedColumn) {
        var best = null;
        var bestScore = -999999;
        for (var k = 0; k < keywords.length; k += 1) {
            var matches = findHeaderCells(sheet, keywords[k], bounds);
            for (var i = 0; i < matches.length; i += 1) {
                var item = matches[i];
                if (excludedColumn && Number(item.column) === Number(excludedColumn)) continue;
                var text = normalizeText(item.text).toLowerCase();
                var key = normalizeText(keywords[k]).toLowerCase();
                var score = 0;
                if (text === key) score += 1000;
                if (text.indexOf(key) === 0) score += 300;
                score -= Math.abs(Number(item.row) - Number(preferredRow)) * 20;
                if (Number(item.row) === Number(preferredRow)) score += 500;
                if (score > bestScore) { bestScore = score; best = item; }
            }
        }
        return best;
    }

    function expandRepeatedHeaders(matches, selected) {
        var result = [];
        var selectedText = normalizeText(selected.text).toLowerCase();
        for (var i = 0; i < matches.length; i += 1) {
            if (Number(matches[i].column) === Number(selected.column) && normalizeText(matches[i].text).toLowerCase() === selectedText) result.push(matches[i]);
        }
        if (result.length === 0) result.push(selected);
        result.sort(function (a, b) { return a.row - b.row; });
        return result;
    }

    function detectChineseMajorRow(sheet, row, bounds) {
        var majorText = "";
        var majorColumn = 0;
        var contentTexts = [];
        for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
            var text = normalizeText(readMergedAwareValue(sheet.Cells.Item(row, column)));
            if (!text) continue;
            if (!majorText && parseMajorNumber(text)) {
                majorText = text;
                majorColumn = column;
                continue;
            }
            if (column !== majorColumn && text !== "序号" && text !== "编号" && text !== "编码") contentTexts.push(text);
        }
        if (!majorText || contentTexts.length === 0) return "";
        return "中文大项：" + majorText + " " + contentTexts.slice(0, 2).join(" ");
    }

    function determineProtection(sheet, row, bounds) {
        return detectChineseMajorRow(sheet, row, bounds);
    }

    function buildBlankColumnPreviewByScope(mode) {
        var context = getTargetContext(false);
        var workbook = context.workbook;
        var originalSheetName = getOriginalActiveSheetName(workbook);
        var keyword = normalizeText(InputBox("请输入判断列的标题。该列为空且本行有其他内容时，将删除整行。\n例如：金额、费用金额、合计。", "按指定列为空删行", "金额"));
        if (!keyword) return;
        var oldPreview = findWorksheet(workbook, PREVIEW_SHEET_NAME);
        if (oldPreview) {
            var oldAlerts = Application.DisplayAlerts;
            try { Application.DisplayAlerts = false; oldPreview.Delete(); } catch (ignored0) {}
            finally { try { Application.DisplayAlerts = oldAlerts; } catch (ignored1) {} }
        }
        var sheetNames = getScopeSheetNames(workbook, context.sheet, mode);
        var plans = [];
        var total = 0;
        var protectedTotal = 0;
        var failures = [];
        for (var i = 0; i < sheetNames.length; i += 1) {
            updateProgress("扫描指定列为空行", i + 1, sheetNames.length);
            try {
                var sheet = activateWorksheet(workbook, sheetNames[i]);
                var bounds = getUsedBounds(sheet);
                var contentRows = getActualContentRowBounds(sheet, bounds);
                if (!contentRows) continue;
                var matches = findHeaderCells(sheet, keyword, bounds);
                if (!matches.length) continue;
                var selected = mode === "sheet" ? chooseHeaderInteractive(matches, keyword, sheetNames[i]) : chooseHeaderAutomatically(matches, keyword, bounds);
                if (!selected) continue;
                var headers = expandRepeatedHeaders(matches, selected);
                var rowMap = {};
                var protectedCount = 0;
                for (var h = 0; h < headers.length; h += 1) {
                    var header = headers[h];
                    var blockEnd = h + 1 < headers.length ? headers[h + 1].row - 1 : contentRows.lastRow;
                    if (blockEnd > contentRows.lastRow) blockEnd = contentRows.lastRow;
                    for (var row = header.row + 1; row <= blockEnd; row += 1) {
                        var target = sheet.Cells.Item(row, header.column);
                        if (cellHasFormula(target) || !isBlankValue(readMergedAwareValue(target))) continue;
                        if (!rowHasAnyContent(sheet, row, bounds)) continue;
                        if (detectChineseMajorRow(sheet, row, bounds)) { protectedCount += 1; continue; }
                        rowMap[row] = true;
                    }
                }
                var rows = [];
                for (var rowText in rowMap) if (rowMap.hasOwnProperty(rowText)) rows.push(Number(rowText));
                rows = uniqueSortedRowsDescending(rows);
                if (rows.length) {
                    plans.push({ sheetName: sheetNames[i], rows: rows });
                    total += rows.length;
                }
                protectedTotal += protectedCount;
            } catch (error) {
                failures.push("“" + sheetNames[i] + "”：" + (error && error.message ? error.message : String(error)));
            }
        }
        restoreActiveSheet(workbook, originalSheetName);
        if (!total) {
            MsgBox("没有发现“" + keyword + "”列为空且本行有内容的可删除行。\n中文大项保护：" + protectedTotal + "行。" + (failures.length ? "\n\n未检查：\n" + failures.slice(0, 12).join("\n") : ""), failures.length ? JS_EXCLAMATION : JS_INFORMATION, "没有可删除行");
            return;
        }
        if (MsgBox("处理范围：" + (mode === "workbook" ? "整个工作簿" : "当前工作表") + "\n判断列：“" + keyword + "”\n检测到" + total + "个待删除行，涉及" + plans.length + "张工作表。\n中文大项保留：" + protectedTotal + "行。\n操作会自动清理尾页残留表头并对齐底部蓝线。\n\n是否继续？", JS_YES_NO + JS_QUESTION, "按指定列为空删行") !== JS_RESULT_YES) return;
        createUndoPoint(workbook, plans.map(function (item) { return item.sheetName; }), "按指定列为空删行");
        runStableRewriteSelfTest(workbook);
        var deleted = 0;
        var orphanDeleted = 0;
        for (var d = 0; d < plans.length; d += 1) {
            updateProgress("按指定列为空删行", d + 1, plans.length);
            try {
                var active = activateWorksheet(workbook, plans[d].sheetName);
                var result = rewriteRowsOnActiveSheet(plans[d].rows);
                deleted += result.deleted;
                for (var fr = 0; fr < result.failedRows.length; fr += 1) failures.push("“" + plans[d].sheetName + "”第" + result.failedRows[fr].row + "行：" + result.failedRows[fr].message);
                var post = postProcessAfterRowChange(active);
                orphanDeleted += post.orphanRows;
            } catch (error2) {
                failures.push("“" + plans[d].sheetName + "”：" + (error2 && error2.message ? error2.message : String(error2)));
            }
        }
        restoreActiveSheet(workbook, originalSheetName);
        var message = "实际上移并核验：" + deleted + "行。";
        if (orphanDeleted) message += "\n另清理尾页残留表头：" + orphanDeleted + "行。";
        if (failures.length) {
            message += "\n\n未完成：\n" + failures.slice(0, 18).join("\n");
            MsgBox(message, JS_EXCLAMATION, deleted ? "部分完成" : "处理失败");
        } else MsgBox(message, JS_INFORMATION, "处理完成");
    }


    function executePreviewDelete() {
        MsgBox("v1.8.0 已取消删除预览表。请直接使用【按指定列为空删行】。", JS_INFORMATION, "功能已简化");
    }


    function chineseNumber(number) {
        var n = Number(number);
        if (!isFinite(n) || Math.floor(n) !== n || n <= 0 || n > 999) throw new Error("中文序号仅支持1到999。");
        var digits = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
        if (n < 10) return digits[n];
        if (n < 20) return "十" + (n % 10 === 0 ? "" : digits[n % 10]);
        if (n < 100) {
            var tens = Math.floor(n / 10);
            var ones = n % 10;
            return digits[tens] + "十" + (ones === 0 ? "" : digits[ones]);
        }
        var hundreds = Math.floor(n / 100);
        var rest = n % 100;
        if (rest === 0) return digits[hundreds] + "百";
        if (rest < 10) return digits[hundreds] + "百零" + digits[rest];
        var restTens = Math.floor(rest / 10);
        var restOnes = rest % 10;
        return digits[hundreds] + "百" + digits[restTens] + "十" + (restOnes === 0 ? "" : digits[restOnes]);
    }

    function parseMajorNumber(value) {
        var text = normalizeText(value);
        var match = text.match(/^([零〇一二三四五六七八九十百]+)\s*([、.．。)]?)$/);
        if (!match) return null;
        return { numeral: match[1], suffix: match[2] || "" };
    }

    function findSerialHeaders(sheet, bounds) {
        var candidates = [];
        for (var row = bounds.firstRow; row <= bounds.lastRow; row += 1) {
            for (var column = bounds.firstColumn; column <= bounds.lastColumn; column += 1) {
                var text = normalizeText(readMergedAwareValue(sheet.Cells.Item(row, column)));
                if (!text) continue;
                if (text.indexOf("序号") >= 0 || text.indexOf("编号") >= 0 || text === "编码") {
                    candidates.push({ row: row, column: column, text: text, address: columnToLetters(column) + row });
                }
            }
        }
        candidates.sort(function (a, b) { return a.row === b.row ? a.column - b.column : a.row - b.row; });
        return candidates;
    }

    function firstTextToRight(sheet, row, serialColumn, bounds) {
        for (var column = serialColumn + 1; column <= bounds.lastColumn; column += 1) {
            var cell = sheet.Cells.Item(row, column);
            if (!isTopLeftOfMerge(cell)) continue;
            var text = normalizeText(readMergedAwareValue(cell));
            if (!text) continue;
            if (/^[\d.,，%％+\-()（）\s]+$/.test(text)) continue;
            return text;
        }
        return "";
    }


    function renumberMajorItemsByScope(mode) {
        var context = getTargetContext(false);
        var workbook = context.workbook;
        var originalSheetName = getOriginalActiveSheetName(workbook);
        var sheetNames = getScopeSheetNames(workbook, context.sheet, mode);
        var plans = [];
        var totalChanges = 0;
        var skipped = [];

        for (var i = 0; i < sheetNames.length; i += 1) {
            var sheetName = sheetNames[i];

            try {
                var sheet = activateWorksheet(workbook, sheetName);
                var bounds = getUsedBounds(sheet);
                var contentRows = getActualContentRowBounds(sheet, bounds);
                if (!contentRows) continue;

                var serialHeaders = findSerialHeaders(sheet, bounds);
                if (serialHeaders.length === 0) {
                    skipped.push(sheetName + "（无序号列）");
                    continue;
                }

                var serialHeader = chooseHeaderAutomatically(
                    serialHeaders,
                    "序号",
                    bounds
                );
                var items = [];

                for (
                    var row = Number(serialHeader.row) + 1;
                    row <= contentRows.lastRow;
                    row += 1
                ) {
                    var serialCell = sheet.Cells.Item(row, serialHeader.column);
                    if (cellHasFormula(serialCell)) continue;

                    var serialText = normalizeText(
                        readMergedAwareValue(serialCell)
                    );
                    var parsed = parseMajorNumber(serialText);
                    if (!parsed) continue;

                    var nameText = firstTextToRight(
                        sheet,
                        row,
                        Number(serialHeader.column),
                        bounds
                    );
                    if (!nameText) continue;

                    items.push({
                        row: row,
                        value: serialText,
                        suffix: parsed.suffix,
                        content: nameText,
                        key: normalizeText(nameText)
                            .replace(/\s+/g, "")
                            .toLowerCase()
                    });
                }

                if (items.length === 0) {
                    skipped.push(sheetName + "（没有中文大项）");
                    continue;
                }

                /*
                 * 同一张工作表中：
                 * - 名称第一次出现时取得下一个中文序号；
                 * - 同名内容在后续分页再次出现时沿用同一序号。
                 */
                var numberByName = {};
                var nextNumber = 1;
                var changes = [];

                for (var j = 0; j < items.length; j += 1) {
                    var item = items[j];
                    if (!numberByName[item.key]) {
                        numberByName[item.key] = nextNumber;
                        nextNumber += 1;
                    }

                    var nextValue =
                        chineseNumber(numberByName[item.key]) +
                        item.suffix;

                    if (item.value !== nextValue) {
                        changes.push({
                            row: item.row,
                            from: item.value,
                            to: nextValue,
                            content: item.content
                        });
                    }
                }

                if (changes.length) {
                    plans.push({
                        sheetName: sheetName,
                        serialColumn: serialHeader.column,
                        changes: changes
                    });
                    totalChanges += changes.length;
                }
            } catch (error) {
                skipped.push(
                    sheetName + "（" +
                    (error && error.message ? error.message : String(error)) +
                    "）"
                );
            }
        }

        restoreActiveSheet(workbook, originalSheetName);

        if (totalChanges === 0) {
            var noChange =
                "没有需要修复的中文大项序号。\n" +
                "规则：名称相同则序号相同；名称第一次出现时才取得下一个序号。";
            if (skipped.length) {
                noChange += "\n\n跳过：" + skipped.slice(0, 12).join("、");
            }
            MsgBox(noChange, JS_INFORMATION, "中文大项无需修改");
            return;
        }

        if (MsgBox(
            "检测到" + totalChanges + "处中文大项需要修正，涉及" + plans.length + "张工作表。\n" +
            "同名内容将使用相同中文序号。\n\n是否继续？",
            JS_YES_NO + JS_QUESTION,
            "中文大项自动顺位"
        ) !== JS_RESULT_YES) return;

        createUndoPoint(workbook, plans.map(function (item) { return item.sheetName; }), "中文大项自动顺位");

        var changedCount = 0;
        var failures = [];

        for (var a = 0; a < plans.length; a += 1) {
            try {
                var activeSheet = activateWorksheet(
                    workbook,
                    plans[a].sheetName
                );

                for (var b = 0; b < plans[a].changes.length; b += 1) {
                    var changeItem = plans[a].changes[b];
                    activeSheet.Cells.Item(
                        changeItem.row,
                        plans[a].serialColumn
                    ).Value2 = changeItem.to;

                    var written = normalizeText(
                        activeSheet.Cells.Item(
                            changeItem.row,
                            plans[a].serialColumn
                        ).Value2
                    );
                    if (written !== changeItem.to) {
                        throw new Error(
                            "第" + changeItem.row +
                            "行写入后核验失败。"
                        );
                    }
                    changedCount += 1;
                }
            } catch (writeError) {
                failures.push(
                    "“" + plans[a].sheetName + "”：" +
                    (writeError && writeError.message
                        ? writeError.message
                        : String(writeError))
                );
            }
        }

        restoreActiveSheet(workbook, originalSheetName);

        var message =
            "已完成中文大项顺位，共修改" +
            changedCount + "处。\n同名内容已使用相同序号。";
        if (skipped.length) {
            message += "\n跳过：" + skipped.slice(0, 10).join("、");
        }

        if (failures.length) {
            message += "\n\n失败：\n" + failures.slice(0, 12).join("\n");
            MsgBox(message, JS_EXCLAMATION, "部分完成");
        } else {
            MsgBox(message, JS_INFORMATION, "操作完成");
        }
    }


    function parseNumericSerial(value) {
        var text = normalizeText(value).replace(/[．。]/g, ".").replace(/\s+/g, "");
        if (!/^\d+(?:\.\d+)*$/.test(text)) return null;
        var parts = text.split(".");
        return { text: text, parts: parts };
    }

    function renumberNumericItemsByScope(mode) {
        var context = getTargetContext(false);
        var workbook = context.workbook;
        var originalSheetName =
            getOriginalActiveSheetName(workbook);
        var sheetNames = getScopeSheetNames(
            workbook,
            context.sheet,
            mode
        );
        var plans = [];
        var skipped = [];
        var failures = [];
        var total = 0;

        for (
            var i = 0;
            i < sheetNames.length;
            i += 1
        ) {
            try {
                var sheet = activateWorksheet(
                    workbook,
                    sheetNames[i]
                );
                var bounds = getUsedBounds(sheet);
                var contentRows =
                    getActualContentRowBounds(
                        sheet,
                        bounds
                    );
                if (!contentRows) continue;

                var headers =
                    findSerialHeaders(
                        sheet,
                        bounds
                    );
                if (!headers.length) {
                    skipped.push(
                        sheetNames[i] +
                        "（无序号列）"
                    );
                    continue;
                }

                var header =
                    chooseHeaderAutomatically(
                        headers,
                        "序号",
                        bounds
                    );
                var headerCell =
                    getWritableMergedTopLeft(
                        sheet.Cells.Item(
                            header.row,
                            header.column
                        )
                    );
                var serialColumn =
                    Number(headerCell.Column);

                var mainMap = {};
                var nextMain = 0;
                var changes = [];

                for (
                    var row =
                        Number(header.row) + 1;
                    row <=
                        contentRows.lastRow;
                    row += 1
                ) {
                    var sourceCell =
                        sheet.Cells.Item(
                            row,
                            serialColumn
                        );
                    if (
                        cellHasFormula(
                            getWritableMergedTopLeft(
                                sourceCell
                            )
                        )
                    ) {
                        continue;
                    }

                    var parsed =
                        parseNumericSerial(
                            readMergedAwareValue(
                                sourceCell
                            )
                        );
                    if (!parsed) continue;

                    var oldMain = String(
                        Number(parsed.parts[0])
                    );
                    if (
                        !mainMap.hasOwnProperty(
                            oldMain
                        )
                    ) {
                        nextMain += 1;
                        mainMap[oldMain] =
                            nextMain;
                    }

                    var expected = String(
                        mainMap[oldMain]
                    );
                    if (
                        parsed.parts.length > 1
                    ) {
                        expected += "." +
                            parsed.parts
                                .slice(1)
                                .join(".");
                    }

                    if (
                        parsed.text !== expected
                    ) {
                        changes.push({
                            row: row,
                            column:
                                serialColumn,
                            value: expected,
                            original:
                                parsed.text
                        });
                    }
                }

                if (changes.length) {
                    plans.push({
                        sheetName:
                            sheetNames[i],
                        changes: changes
                    });
                    total += changes.length;
                }
            } catch (error) {
                failures.push(
                    "“" + sheetNames[i] +
                    "”：" +
                    (
                        error &&
                        error.message
                            ? error.message
                            : String(error)
                    )
                );
            }
        }

        restoreActiveSheet(
            workbook,
            originalSheetName
        );

        if (!total) {
            MsgBox(
                "没有需要修正的数字序号。",
                JS_INFORMATION,
                "无需修改"
            );
            return;
        }

        if (
            MsgBox(
                "检测到" + total +
                "处数字序号需要修正，涉及" +
                plans.length +
                "张工作表。" +
                "\n首个主项将从1开始，" +
                "小数后缀保留。" +
                "\n\n是否继续？",
                JS_YES_NO + JS_QUESTION,
                "数字序号顺位"
            ) !== JS_RESULT_YES
        ) {
            return;
        }

        createUndoPoint(
            workbook,
            plans.map(function (item) {
                return item.sheetName;
            }),
            "数字序号顺位"
        );

        var changed = 0;

        for (
            var p = 0;
            p < plans.length;
            p += 1
        ) {
            var active = null;
            try {
                active = activateWorksheet(
                    workbook,
                    plans[p].sheetName
                );
            } catch (activateError) {
                failures.push(
                    "“" + plans[p].sheetName +
                    "”：" +
                    (
                        activateError &&
                        activateError.message
                            ? activateError.message
                            : String(activateError)
                    )
                );
                continue;
            }

            for (
                var c = 0;
                c <
                    plans[p].changes.length;
                c += 1
            ) {
                var item =
                    plans[p].changes[c];
                try {
                    var originalCell =
                        active.Cells.Item(
                            item.row,
                            item.column
                        );
                    var writableCell =
                        getWritableMergedTopLeft(
                            originalCell
                        );
                    var oldNumberFormat =
                        null;
                    try {
                        oldNumberFormat =
                            writableCell
                                .NumberFormat;
                    } catch (ignored0) {}

                    var written = false;

                    try {
                        writableCell.Value2 =
                            item.value;
                        written =
                            normalizeText(
                                readMergedAwareValue(
                                    originalCell
                                )
                            ) === item.value;
                    } catch (ignored1) {}

                    if (!written) {
                        try {
                            writableCell
                                .ClearContents();
                        } catch (ignored2) {}
                        try {
                            writableCell.Value =
                                item.value;
                            written =
                                normalizeText(
                                    readMergedAwareValue(
                                        originalCell
                                    )
                                ) === item.value;
                        } catch (ignored3) {}
                    }

                    if (
                        oldNumberFormat !== null
                    ) {
                        try {
                            writableCell
                                .NumberFormat =
                                oldNumberFormat;
                        } catch (ignored4) {}
                    }

                    if (!written) {
                        throw new Error(
                            "第" + item.row +
                            "行写入后仍为“" +
                            normalizeText(
                                readMergedAwareValue(
                                    originalCell
                                )
                            ) +
                            "”，预计“" +
                            item.value + "”。"
                        );
                    }

                    changed += 1;
                } catch (writeError) {
                    failures.push(
                        "“" +
                        plans[p].sheetName +
                        "”第" + item.row +
                        "行：" +
                        (
                            writeError &&
                            writeError.message
                                ? writeError.message
                                : String(writeError)
                        )
                    );
                }
            }
        }

        restoreActiveSheet(
            workbook,
            originalSheetName
        );

        var message =
            "数字序号顺位完成，共修改" +
            changed + "处。" +
            "\n首个主项从1开始，" +
            "小数后缀原样保留。";

        if (skipped.length) {
            message +=
                "\n跳过：" +
                skipped
                    .slice(0, 10)
                    .join("、");
        }

        if (failures.length) {
            message +=
                "\n\n失败：" +
                failures.length + "项。";
            MsgBox(
                message,
                JS_EXCLAMATION,
                changed
                    ? "部分完成"
                    : "处理失败"
            );
        } else {
            MsgBox(
                message,
                JS_INFORMATION,
                "操作完成"
            );
        }
    }


    function deleteSelectedBlankCellRows() {
        var context = getTargetContext(false);
        var workbook = context.workbook;
        var sheet = context.sheet;
        var selection = Application.Selection;
        if (!selection) throw new Error("请先选中一个或多个空白单元格。");
        var bounds = getUsedBounds(sheet);
        var rowMap = {};
        var selectedCount = 0;
        var areaCount = 1;
        var areas = null;
        try { areas = selection.Areas; areaCount = Number(areas.Count); } catch (ignored) {}
        for (var a = 1; a <= areaCount; a += 1) {
            var area = areas ? areas.Item(a) : selection;
            var rc = Number(area.Rows.Count), cc = Number(area.Columns.Count);
            selectedCount += rc * cc;
            if (selectedCount > 20000) throw new Error("选区过大，请只选需要删除行中的空白单元格。");
            for (var r = 1; r <= rc; r += 1) for (var c = 1; c <= cc; c += 1) {
                var cell = area.Cells.Item(r, c);
                if (cellHasFormula(cell) || !isBlankValue(readMergedAwareValue(cell))) continue;
                rowMap[Number(cell.Row)] = true;
            }
        }
        var rows = [];
        var protectedCount = 0;
        for (var rowText in rowMap) if (rowMap.hasOwnProperty(rowText)) {
            var rowNumber = Number(rowText);
            if (detectChineseMajorRow(sheet, rowNumber, bounds)) { protectedCount += 1; continue; }
            rows.push(rowNumber);
        }
        rows = uniqueSortedRowsDescending(rows);
        if (!rows.length) {
            MsgBox("选区中没有可移除的空白单元格所在行。\n中文大项保护：" + protectedCount + "行。", JS_INFORMATION, "没有可删除行");
            return;
        }
        if (MsgBox("将直接移除所选空白单元格所在的" + rows.length + "行。\n中文大项保护：" + protectedCount + "行。\n操作会清理尾部残留并对齐底部蓝线。\n\n是否继续？", JS_YES_NO + JS_QUESTION, "删除所选空白格所在行") !== JS_RESULT_YES) return;
        createUndoPoint(workbook, [String(sheet.Name)], "删除所选空白格所在行");
        runStableRewriteSelfTest(workbook);
        var active = activateWorksheet(workbook, String(sheet.Name));
        var result = rewriteRowsOnActiveSheet(rows);
        var post = postProcessAfterRowChange(active);
        var message = "实际上移并核验：" + result.deleted + "行。";
        if (post.orphanRows) message += "\n另清理尾页残留表头：" + post.orphanRows + "行。";
        if (result.failedRows.length) {
            var lines = [];
            for (var i = 0; i < result.failedRows.length; i += 1) lines.push("第" + result.failedRows[i].row + "行：" + result.failedRows[i].message);
            message += "\n\n未完成：\n" + lines.slice(0, 15).join("\n");
            MsgBox(message, JS_EXCLAMATION, result.deleted ? "部分完成" : "处理失败");
        } else MsgBox(message, JS_INFORMATION, "处理完成");
    }

    function deleteRowsBySelectedTextByScope(mode) {
        var context = getTargetContext(false);
        var workbook = context.workbook;
        var originalSheetName = getOriginalActiveSheetName(workbook);
        var selectedCell = Application.ActiveCell;
        if (!selectedCell) throw new Error("请先选中包含目标文字的单元格。");
        var selectedText = normalizeText(readMergedAwareValue(selectedCell));
        if (!selectedText) throw new Error("选中的单元格没有文字。");
        var targetColumn = Number(selectedCell.Column);
        var sheetNames = getScopeSheetNames(workbook, context.sheet, mode);
        var plans = [];
        var total = 0;
        var failures = [];
        for (var i = 0; i < sheetNames.length; i += 1) {
            try {
                var sheet = activateWorksheet(workbook, sheetNames[i]);
                var bounds = getUsedBounds(sheet);
                var contentRows = getActualContentRowBounds(sheet, bounds);
                if (!contentRows || targetColumn > bounds.lastColumn) continue;
                var rows = [];
                for (var row = contentRows.firstRow; row <= contentRows.lastRow; row += 1) {
                    var cell = sheet.Cells.Item(row, targetColumn);
                    if (cellHasFormula(cell)) continue;
                    if (normalizeText(readMergedAwareValue(cell)) === selectedText) rows.push(row);
                }
                if (rows.length) { plans.push({ sheetName: sheetNames[i], rows: rows }); total += rows.length; }
            } catch (error) { failures.push("“" + sheetNames[i] + "”：" + (error && error.message ? error.message : String(error))); }
        }
        restoreActiveSheet(workbook, originalSheetName);
        if (!total) {
            MsgBox("在所选范围的" + columnToLetters(targetColumn) + "列中，没有找到与“" + selectedText + "”完全相同的文字。", JS_INFORMATION, "没有匹配行");
            return;
        }
        if (MsgBox("目标文字：“" + selectedText + "”\n检测到" + total + "个匹配行，涉及" + plans.length + "张工作表。\n\n是否删除？", JS_YES_NO + JS_QUESTION, "按选中文字删除整行") !== JS_RESULT_YES) return;
        createUndoPoint(workbook, plans.map(function (item) { return item.sheetName; }), "按选中文字删除整行");
        runStableRewriteSelfTest(workbook);
        var deleted = 0;
        var orphanDeleted = 0;
        for (var d = 0; d < plans.length; d += 1) {
            try {
                var active = activateWorksheet(workbook, plans[d].sheetName);
                var result = rewriteRowsOnActiveSheet(plans[d].rows);
                deleted += result.deleted;
                var post = postProcessAfterRowChange(active);
                orphanDeleted += post.orphanRows;
                for (var f = 0; f < result.failedRows.length; f += 1) failures.push("“" + plans[d].sheetName + "”第" + result.failedRows[f].row + "行：" + result.failedRows[f].message);
            } catch (error2) { failures.push("“" + plans[d].sheetName + "”：" + (error2 && error2.message ? error2.message : String(error2))); }
        }
        restoreActiveSheet(workbook, originalSheetName);
        var message = "实际上移并核验：" + deleted + "行。";
        if (orphanDeleted) message += "\n另清理尾页残留表头：" + orphanDeleted + "行。";
        if (failures.length) {
            message += "\n\n未完成：\n" + failures.slice(0, 18).join("\n");
            MsgBox(message, JS_EXCLAMATION, deleted ? "部分完成" : "处理失败");
        } else MsgBox(message, JS_INFORMATION, "操作完成");
    }


    function deleteNamedSheetSilently(workbook, sheetName) {
        var sheet = findWorksheet(workbook, sheetName);
        if (!sheet) return false;
        var oldAlerts = Application.DisplayAlerts;
        try {
            Application.DisplayAlerts = false;
            sheet.Delete();
            return true;
        } finally {
            try { Application.DisplayAlerts = oldAlerts; } catch (ignored) {}
        }
    }

    function cleanupLegacyVisibleSheets() {
        var workbook = Application.ActiveWorkbook;
        if (!workbook) return 0;
        var names = [PREVIEW_SHEET_NAME, DUPLICATE_SHEET_NAME, "__重写自检残留"];
        var activeName = "";
        try { activeName = String(Application.ActiveSheet.Name); } catch (ignored0) {}
        var mustMove = false;
        for (var n = 0; n < names.length; n += 1) {
            if (activeName === names[n]) { mustMove = true; break; }
        }
        if (mustMove) {
            for (var i = 1; i <= workbook.Worksheets.Count; i += 1) {
                var candidate = workbook.Worksheets.Item(i);
                if (!isAssistantInternalSheetName(String(candidate.Name))) {
                    candidate.Activate();
                    break;
                }
            }
        }
        var deleted = 0;
        for (var j = 0; j < names.length; j += 1) {
            if (deleteNamedSheetSilently(workbook, names[j])) deleted += 1;
        }
        return deleted;
    }

    function duplicateContentKey(value) {
        return normalizeText(value)
            .replace(/\s+/g, " ")
            .toLowerCase();
    }

    function normalizeSearchTextKey(
        value
    ) {
        return normalizeText(value)
            .replace(/[０-９]/g, function (character) {
                return String.fromCharCode(
                    character.charCodeAt(0) - 0xfee0
                );
            })
            .replace(/[Ａ-Ｚａ-ｚ]/g, function (character) {
                return String.fromCharCode(
                    character.charCodeAt(0) - 0xfee0
                );
            })
            .replace(/\s+/g, "")
            .toLowerCase();
    }

    function textMatchesSearch(
        cellText,
        targetText,
        matchMode
    ) {
        var cellKey =
            normalizeSearchTextKey(
                cellText
            );
        var targetKey =
            normalizeSearchTextKey(
                targetText
            );

        if (!cellKey || !targetKey) {
            return false;
        }
        if (matchMode === "exact") {
            return cellKey === targetKey;
        }
        return cellKey.indexOf(targetKey) >= 0;
    }


    function getOccurrenceFromCell(
        sheet,
        cell,
        bounds
    ) {
        var rowStart = Number(cell.Row);
        var rowEnd = rowStart;
        var columnStart = Number(cell.Column);
        var columnEnd = columnStart;
        var address =
            columnToLetters(columnStart) +
            rowStart;

        try {
            if (cell.MergeCells) {
                var area = cell.MergeArea;
                if (!isTopLeftOfMerge(cell)) {
                    return null;
                }
                rowStart = Number(area.Row);
                rowEnd =
                    rowStart +
                    Number(area.Rows.Count) -
                    1;
                columnStart =
                    Number(area.Column);
                columnEnd =
                    columnStart +
                    Number(
                        area.Columns.Count
                    ) - 1;
                address =
                    columnToLetters(
                        columnStart
                    ) +
                    rowStart + ":" +
                    columnToLetters(
                        columnEnd
                    ) +
                    rowEnd;
            }
        } catch (ignored) {}

        return {
            sheetName:
                String(sheet.Name),
            rowStart: rowStart,
            rowEnd: rowEnd,
            columnStart: columnStart,
            columnEnd: columnEnd,
            address: address,
            summary: summarizeRow(
                sheet,
                rowStart,
                bounds,
                14,
                600
            )
        };
    }

    function scanExactTextOccurrences(
        workbook,
        context,
        mode,
        targetText,
        selectedLocation,
        matchMode
    ) {
        var originalSheetName =
            getOriginalActiveSheetName(
                workbook
            );
        var sheetNames =
            getScopeSheetNames(
                workbook,
                context.sheet,
                mode
            );
        var items = [];
        var failures = [];
        var scannedCells = 0;
        var effectiveMode =
            matchMode === "exact"
                ? "exact"
                : "contains";

        for (
            var i = 0;
            i < sheetNames.length;
            i += 1
        ) {
            updateProgress(
                "查找文字内容",
                i + 1,
                sheetNames.length
            );

            try {
                var sheet =
                    activateWorksheet(
                        workbook,
                        sheetNames[i]
                    );
                var actual =
                    getActualContentBounds(
                        sheet
                    );
                if (!actual) continue;

                var rowCount =
                    actual.lastRow -
                    actual.firstRow + 1;
                var columnCount =
                    actual.lastColumn -
                    actual.firstColumn + 1;
                scannedCells +=
                    rowCount * columnCount;

                if (
                    scannedCells >
                    MAX_TEXT_SEARCH_CELLS
                ) {
                    throw new Error(
                        "累计检查单元格超过" +
                        MAX_TEXT_SEARCH_CELLS +
                        "个。为避免WPS卡死，请改用当前工作表或先缩小范围。"
                    );
                }

                for (
                    var row = actual.firstRow;
                    row <= actual.lastRow;
                    row += 1
                ) {
                    for (
                        var column = actual.firstColumn;
                        column <= actual.lastColumn;
                        column += 1
                    ) {
                        var cell =
                            sheet.Cells.Item(
                                row,
                                column
                            );
                        try {
                            if (
                                cell.MergeCells &&
                                !isTopLeftOfMerge(cell)
                            ) {
                                continue;
                            }
                        } catch (ignored0) {}

                        var text =
                            normalizeText(
                                readMergedAwareValue(
                                    cell
                                )
                            );
                        if (
                            !text ||
                            !textMatchesSearch(
                                text,
                                targetText,
                                effectiveMode
                            )
                        ) {
                            continue;
                        }

                        var item =
                            getOccurrenceFromCell(
                                sheet,
                                cell,
                                actual
                            );
                        if (!item) continue;

                        item.text = text;
                        item.matchMode = effectiveMode;
                        item.isSelected =
                            selectedLocation &&
                            item.sheetName === selectedLocation.sheetName &&
                            item.rowStart <= selectedLocation.row &&
                            item.rowEnd >= selectedLocation.row &&
                            item.columnStart <= selectedLocation.column &&
                            item.columnEnd >= selectedLocation.column;
                        items.push(item);

                        if (
                            items.length >
                            MAX_DUPLICATE_RECORDS
                        ) {
                            throw new Error(
                                "匹配文字记录超过" +
                                MAX_DUPLICATE_RECORDS +
                                "条，已停止检查。"
                            );
                        }
                    }
                }
            } catch (error) {
                failures.push(
                    "“" + sheetNames[i] +
                    "”：" +
                    (
                        error && error.message
                            ? error.message
                            : String(error)
                    )
                );
            }
        }

        restoreActiveSheet(
            workbook,
            originalSheetName
        );

        items.sort(function (a, b) {
            if (a.isSelected !== b.isSelected) {
                return a.isSelected ? -1 : 1;
            }
            if (a.sheetName !== b.sheetName) {
                return a.sheetName < b.sheetName ? -1 : 1;
            }
            if (a.rowStart !== b.rowStart) {
                return a.rowStart - b.rowStart;
            }
            return a.columnStart - b.columnStart;
        });

        return {
            items: items,
            failures: failures,
            originalSheetName: originalSheetName,
            scannedCells: scannedCells,
            matchMode: effectiveMode
        };
    }


    function buildAllNumberSelection(maxValue) {
        var result = [];
        var upper = Number(maxValue);
        if (!isFinite(upper) || upper < 1) return result;
        upper = Math.floor(upper);
        for (var i = 1; i <= upper; i += 1) {
            result.push(i);
        }
        return result;
    }

    function promptExactOccurrenceDeletion(
        targetText,
        items
    ) {
        var lines = [];
        var displayCount =
            Math.min(items.length, 24);

        for (
            var i = 0;
            i < displayCount;
            i += 1
        ) {
            var rowLabel =
                items[i].rowStart ===
                items[i].rowEnd
                    ? "第" +
                      items[i].rowStart +
                      "行"
                    : "第" +
                      items[i].rowStart +
                      "—" +
                      items[i].rowEnd +
                      "行";
            lines.push(
                (i + 1) + ". " +
                items[i].sheetName +
                "!" + items[i].address +
                " ｜ " + rowLabel +
                (
                    items[i].isSelected
                        ? "（当前选中）"
                        : ""
                ) +
                "\n   " +
                clipPromptText(
                    items[i].summary,
                    82
                )
            );
        }

        if (
            items.length >
            displayCount
        ) {
            lines.push(
                "……另有" +
                (
                    items.length -
                    displayCount
                ) +
                "条未展开。"
            );
        }

        var answer =
            normalizeText(
                InputBox(
                    "查找文字：“" +
                    clipPromptText(
                        targetText,
                        90
                    ) +
                    "”\n共找到" +
                    items.length +
                    "处：\n\n" +
                    lines.join("\n") +
                    "\n\n请输入需要清空文字的记录：" +
                    "\n• 全部：清空所有匹配文字，包括当前选中项" +
                    "\n• 除当前：保留当前选中项，清空其他匹配文字" +
                    "\n• 也可输入1-8或1,3,6" +
                    "\n• 留空取消" +
                    "\n\n只清空单元格内容，表格行列、格式、边框和合并状态不变。",
                    "选择要清空的匹配文字",
                    "全部"
                )
            );

        if (!answer) return [];

        var command = answer
            .replace(/\s+/g, "")
            .toLowerCase();

        if (
            command === "全部" ||
            command === "全清" ||
            command === "all"
        ) {
            return buildAllNumberSelection(
                items.length
            );
        }

        if (
            command === "除当前" ||
            command === "保留当前" ||
            command === "exceptcurrent"
        ) {
            var exceptCurrent = [];
            for (
                var index = 0;
                index < items.length;
                index += 1
            ) {
                if (!items[index].isSelected) {
                    exceptCurrent.push(index + 1);
                }
            }
            return exceptCurrent;
        }

        if (
            command === "当前" ||
            command === "只清当前"
        ) {
            for (
                var currentIndex = 0;
                currentIndex < items.length;
                currentIndex += 1
            ) {
                if (items[currentIndex].isSelected) {
                    return [currentIndex + 1];
                }
            }
            return [1];
        }

        return parseNumberSelection(
            answer,
            items.length
        );
    }


    function clearExactTextOccurrence(
        sheet,
        item,
        targetText,
        matchMode
    ) {
        var sourceCell =
            sheet.Cells.Item(
                item.rowStart,
                item.columnStart
            );
        var writableCell =
            getWritableMergedTopLeft(
                sourceCell
            );
        var currentText =
            normalizeText(
                readMergedAwareValue(
                    sourceCell
                )
            );

        if (
            !textMatchesSearch(
                currentText,
                targetText,
                matchMode
            )
        ) {
            throw new Error(
                "原位置内容已经变化，当前为“" +
                clipPromptText(
                    currentText,
                    60
                ) +
                "”。"
            );
        }

        var cleared = false;
        try {
            writableCell.ClearContents();
            cleared =
                isBlankValue(
                    readMergedAwareValue(
                        sourceCell
                    )
                );
        } catch (ignored0) {}

        if (!cleared) {
            try {
                writableCell.Formula = "";
            } catch (ignored1) {}
            try {
                writableCell.Value2 = "";
            } catch (ignored2) {}
            cleared =
                isBlankValue(
                    readMergedAwareValue(
                        sourceCell
                    )
                );
        }

        if (!cleared) {
            throw new Error(
                "内容清空后核验失败。"
            );
        }
        return true;
    }


    function buildDuplicateSelectionByScope(mode) {
        var context =
            getTargetContext(false);
        var workbook =
            context.workbook;
        var activeCell =
            Application.ActiveCell;
        var defaultText = "";

        try {
            if (activeCell) {
                defaultText =
                    normalizeText(
                        readMergedAwareValue(
                            activeCell
                        )
                    );
            }
        } catch (ignored0) {}

        var targetText =
            normalizeText(
                InputBox(
                    "请输入要查找并清空的文字。" +
                    "\n默认采用“包含关键词”匹配，并忽略空格。" +
                    "\n例如输入“第1页”，可匹配“第 1 页 共 2 页”“第1页共5页”。" +
                    "\n\n只清空匹配单元格文字，不删除行列或改变表格。",
                    "输入查找文字",
                    defaultText
                )
            );
        if (!targetText) return;

        var modeAnswer =
            normalizeText(
                InputBox(
                    "请选择匹配方式：" +
                    "\n• 包含：单元格文字中包含关键词即可" +
                    "\n• 完全相同：忽略空格后必须完全一致" +
                    "\n\n输入“包含”或“完全相同”。",
                    "选择匹配方式",
                    "包含"
                )
            );
        if (!modeAnswer) return;

        var modeKey =
            modeAnswer
                .replace(/\s+/g, "")
                .toLowerCase();
        var matchMode = "contains";
        if (
            modeKey === "完全相同" ||
            modeKey === "完全" ||
            modeKey === "精确" ||
            modeKey === "exact"
        ) {
            matchMode = "exact";
        } else if (
            modeKey !== "包含" &&
            modeKey !== "模糊" &&
            modeKey !== "contains"
        ) {
            throw new Error(
                "无法识别匹配方式：“" +
                modeAnswer +
                "”。请输入“包含”或“完全相同”。"
            );
        }

        var selectedLocation = null;
        try {
            if (activeCell) {
                selectedLocation = {
                    sheetName: String(
                        Application.ActiveSheet.Name
                    ),
                    row: Number(activeCell.Row),
                    column: Number(activeCell.Column)
                };
            }
        } catch (ignored1) {}

        var scan =
            scanExactTextOccurrences(
                workbook,
                context,
                mode,
                targetText,
                selectedLocation,
                matchMode
            );

        if (!scan.items.length) {
            MsgBox(
                "所选范围没有找到匹配文字：“" +
                targetText +
                "”。" +
                "\n匹配方式：" +
                (
                    matchMode === "exact"
                        ? "完全相同"
                        : "包含关键词"
                ) +
                (
                    scan.failures.length
                        ? "\n另有" +
                          scan.failures.length +
                          "张表未完成检查。"
                        : ""
                ),
                scan.failures.length
                    ? JS_EXCLAMATION
                    : JS_INFORMATION,
                "没有匹配文字"
            );
            return;
        }

        var selectedIndexes =
            promptExactOccurrenceDeletion(
                targetText,
                scan.items
            );
        if (!selectedIndexes.length) return;

        var plansBySheet = {};
        var selectedOccurrences = 0;
        for (
            var s = 0;
            s < selectedIndexes.length;
            s += 1
        ) {
            var item =
                scan.items[
                    selectedIndexes[s] - 1
                ];
            if (!plansBySheet[item.sheetName]) {
                plansBySheet[item.sheetName] = [];
            }
            plansBySheet[item.sheetName].push(item);
            selectedOccurrences += 1;
        }

        var targetNames = [];
        for (var sheetName in plansBySheet) {
            if (plansBySheet.hasOwnProperty(sheetName)) {
                targetNames.push(sheetName);
            }
        }

        if (
            MsgBox(
                "查找文字：“" +
                clipPromptText(targetText, 90) +
                "”" +
                "\n匹配方式：" +
                (
                    matchMode === "exact"
                        ? "完全相同"
                        : "包含关键词"
                ) +
                "\n将清空" +
                selectedOccurrences +
                "处匹配文字，涉及" +
                targetNames.length +
                "张工作表。" +
                "\n\n只清空整个匹配单元格的文字内容：" +
                "\n• 不删除整行或整列" +
                "\n• 不移动任何表格" +
                "\n• 不改变边框、格式和合并状态" +
                "\n\n是否继续？",
                JS_YES_NO + JS_QUESTION,
                "查找并清空文字"
            ) !== JS_RESULT_YES
        ) {
            return;
        }

        createUndoPoint(
            workbook,
            targetNames,
            "查找并清空文字"
        );

        var clearedCount = 0;
        var failures = scan.failures.slice();

        for (
            var t = 0;
            t < targetNames.length;
            t += 1
        ) {
            var targetName = targetNames[t];
            updateProgress(
                "查找并清空文字",
                t + 1,
                targetNames.length
            );
            try {
                var sheet =
                    activateWorksheet(
                        workbook,
                        targetName
                    );
                var items = plansBySheet[targetName];
                for (
                    var i = 0;
                    i < items.length;
                    i += 1
                ) {
                    try {
                        if (
                            clearExactTextOccurrence(
                                sheet,
                                items[i],
                                targetText,
                                matchMode
                            )
                        ) {
                            clearedCount += 1;
                        }
                    } catch (itemError) {
                        failures.push(
                            "“" + targetName +
                            "”" + items[i].address +
                            "：" +
                            (
                                itemError && itemError.message
                                    ? itemError.message
                                    : String(itemError)
                            )
                        );
                    }
                }
            } catch (sheetError) {
                failures.push(
                    "“" + targetName +
                    "”：" +
                    (
                        sheetError && sheetError.message
                            ? sheetError.message
                            : String(sheetError)
                    )
                );
            }
        }

        restoreActiveSheet(
            workbook,
            scan.originalSheetName
        );

        var verifyContext =
            getTargetContext(false);
        var verifyScan =
            scanExactTextOccurrences(
                workbook,
                verifyContext,
                mode,
                targetText,
                null,
                matchMode
            );
        var expectedRemaining =
            scan.items.length -
            selectedOccurrences;
        if (
            verifyScan.items.length !==
            expectedRemaining
        ) {
            failures.push(
                "清空后仍找到" +
                verifyScan.items.length +
                "处匹配文字，预计" +
                expectedRemaining +
                "处。"
            );
        }

        var message =
            "查找文字：“" + targetText +
            "”" +
            "\n匹配方式：" +
            (
                matchMode === "exact"
                    ? "完全相同"
                    : "包含关键词"
            ) +
            "\n清空前找到：" +
            scan.items.length +
            "处。" +
            "\n选择清空：" +
            selectedOccurrences +
            "处。" +
            "\n实际清空并核验：" +
            clearedCount +
            "处。" +
            "\n清空后剩余：" +
            verifyScan.items.length +
            "处。" +
            "\n\n表格行列、格式、边框和合并状态均未改动。";

        if (failures.length) {
            message +=
                "\n未完成或核验异常：" +
                failures.length +
                "项。";
            MsgBox(
                message,
                JS_EXCLAMATION,
                clearedCount
                    ? "部分完成"
                    : "处理失败"
            );
        } else {
            MsgBox(
                message,
                JS_INFORMATION,
                "文字内容已清空"
            );
        }
    }


    function executeDuplicateSelectionDelete() {
        MsgBox(
            "新版已将查找、选择和清空文字合并到【清除相同文字内容】中。",
            JS_INFORMATION,
            "功能已合并"
        );
    }

    function locateDuplicateSelectionRow() {
        MsgBox(
            "新版不生成选择工作表，清空文字时不会移动表格，因此不需要定位按钮。",
            JS_INFORMATION,
            "功能已取消"
        );
    }

    function lettersToColumn(letters) {
        var text = normalizeText(letters).toUpperCase().replace(/\$/g, "");
        if (!/^[A-Z]+$/.test(text)) return 0;
        var result = 0;
        for (var i = 0; i < text.length; i += 1) result = result * 26 + (text.charCodeAt(i) - 64);
        return result;
    }

    function uniqueSortedColumnsDescending(columns) {
        var map = {};
        var result = [];
        for (var i = 0; i < columns.length; i += 1) {
            var column = Number(columns[i]);
            if (!isFinite(column) || column < 1 || Math.floor(column) !== column || map[column]) continue;
            map[column] = true;
            result.push(column);
        }
        result.sort(function (a, b) { return b - a; });
        return result;
    }

    function formatColumnList(columns) {
        var sorted = uniqueSortedColumnsDescending(columns).slice().sort(function (a, b) { return a - b; });
        var parts = [];
        var startValue = null;
        var previous = null;
        for (var i = 0; i < sorted.length; i += 1) {
            var value = sorted[i];
            if (startValue === null) { startValue = value; previous = value; continue; }
            if (value === previous + 1) { previous = value; continue; }
            parts.push(startValue === previous ? columnToLetters(startValue) : columnToLetters(startValue) + ":" + columnToLetters(previous));
            startValue = value;
            previous = value;
        }
        if (startValue !== null) parts.push(startValue === previous ? columnToLetters(startValue) : columnToLetters(startValue) + ":" + columnToLetters(previous));
        return parts.join("、");
    }

    function parseColumnSpec(specification) {
        var text = normalizeText(specification).toUpperCase().replace(/[，；;、\s]+/g, ",");
        var tokens = text.split(",");
        var columns = [];
        for (var i = 0; i < tokens.length; i += 1) {
            var token = normalizeText(tokens[i]);
            if (!token) continue;
            var rangeMatch = token.match(/^([A-Z]+)\s*:\s*([A-Z]+)$/);
            if (rangeMatch) {
                var startColumn = lettersToColumn(rangeMatch[1]);
                var endColumn = lettersToColumn(rangeMatch[2]);
                if (!startColumn || !endColumn) throw new Error("无法识别列范围：“" + token + "”。");
                if (startColumn > endColumn) { var temp = startColumn; startColumn = endColumn; endColumn = temp; }
                if (endColumn - startColumn + 1 > 100) throw new Error("一次输入的连续列不能超过100列。");
                for (var column = startColumn; column <= endColumn; column += 1) columns.push(column);
            } else {
                var number = lettersToColumn(token);
                if (!number) throw new Error("无法识别列：“" + token + "”。示例：C、F、H:J");
                columns.push(number);
            }
        }
        columns = uniqueSortedColumnsDescending(columns);
        if (columns.length === 0) throw new Error("没有输入有效列号。");
        if (columns.length > 100) throw new Error("一次最多删除100列。");
        return columns;
    }

    function getSelectedColumnNumbers() {
        var selection = Application.Selection;
        if (!selection) throw new Error("请先在工作表中选中需要删除的列或单元格。");
        var result = [];
        try {
            var areaCount = Number(selection.Areas.Count);
            if (isFinite(areaCount) && areaCount > 0) {
                for (var areaIndex = 1; areaIndex <= areaCount; areaIndex += 1) {
                    var area = selection.Areas.Item(areaIndex);
                    var firstColumn = Number(area.Column);
                    var columnCount = Number(area.Columns.Count);
                    for (var offset = 0; offset < columnCount; offset += 1) result.push(firstColumn + offset);
                }
            }
        } catch (e) {}
        if (result.length === 0) {
            var first = Number(selection.Column);
            var count = Number(selection.Columns.Count);
            for (var i = 0; i < count; i += 1) result.push(first + i);
        }
        result = uniqueSortedColumnsDescending(result);
        if (result.length === 0) throw new Error("没有识别到需要删除的列。");
        if (result.length > 100) throw new Error("当前选区覆盖" + result.length + "列，超过安全限制。");
        return result;
    }

    function collectAffectedColumnMerges(sheet, columns, bounds) {
        var map = {};
        var result = [];
        for (var c = 0; c < columns.length; c += 1) {
            var column = columns[c];
            if (column < bounds.firstColumn || column > bounds.lastColumn) continue;
            for (var row = bounds.firstRow; row <= bounds.lastRow; row += 1) {
                var cell = sheet.Cells.Item(row, column);
                try {
                    if (!cell.MergeCells) continue;
                    var area = cell.MergeArea;
                    var key = area.Row + ":" + area.Column + ":" + area.Rows.Count + ":" + area.Columns.Count;
                    if (map[key]) continue;
                    map[key] = true;
                    result.push(saveMergeArea(area));
                } catch (e) {}
            }
        }
        return result;
    }

    function restoreColumnMerges(sheet, merges, deletedColumns) {
        var ascending = uniqueSortedColumnsDescending(deletedColumns).slice().sort(function (a, b) { return a - b; });
        for (var i = 0; i < merges.length; i += 1) {
            var item = merges[i];
            var oldEndColumn = item.column + item.columnCount - 1;
            var removedInside = countDeletedInRange(ascending, item.column, oldEndColumn);
            var remainingColumns = item.columnCount - removedInside;
            if (remainingColumns <= 0) continue;
            var newColumn = item.column - countDeletedLessThan(ascending, item.column);
            var newEndColumn = newColumn + remainingColumns - 1;
            var address = columnToLetters(newColumn) + item.row + ":" + columnToLetters(newEndColumn) + (item.row + item.rowCount - 1);
            var area = sheet.Range(address);
            if (item.rowCount > 1 || remainingColumns > 1) area.Merge();
            var topLeft = sheet.Cells.Item(item.row, newColumn);
            try { if (item.hasFormula && item.formula) topLeft.Formula = item.formula; else topLeft.Value2 = item.value; } catch (e1) {}
            try { if (item.horizontalAlignment !== null) area.HorizontalAlignment = item.horizontalAlignment; } catch (e2) {}
            try { if (item.verticalAlignment !== null) area.VerticalAlignment = item.verticalAlignment; } catch (e3) {}
            try { if (item.wrapText !== null) area.WrapText = item.wrapText; } catch (e4) {}
            try { if (item.numberFormat !== null) topLeft.NumberFormat = item.numberFormat; } catch (e5) {}
        }
    }


    function groupColumnBlocksAscending(columns) {
        var ascending = uniqueSortedColumnsDescending(columns).slice().sort(function (a, b) { return a - b; });
        var blocks = [];
        if (!ascending.length) return blocks;
        var start = ascending[0], end = ascending[0];
        for (var i = 1; i < ascending.length; i += 1) {
            if (ascending[i] === end + 1) end = ascending[i];
            else { blocks.push({ start: start, end: end }); start = ascending[i]; end = ascending[i]; }
        }
        blocks.push({ start: start, end: end });
        return blocks;
    }

    function captureColumnWidthPlans(sheet, columns) {
        var ascending = uniqueSortedColumnsDescending(columns).slice().sort(function (a, b) { return a - b; });
        var deleteMap = {};
        for (var i = 0; i < ascending.length; i += 1) deleteMap[ascending[i]] = true;
        var blocks = groupColumnBlocksAscending(ascending);
        var plans = [];
        for (var b = 0; b < blocks.length; b += 1) {
            var block = blocks[b];
            var absorber = block.start - 1;
            if (absorber < 1 || deleteMap[absorber]) absorber = block.end + 1;
            var deletedWidth = 0;
            for (var c = block.start; c <= block.end; c += 1) {
                try { deletedWidth += Number(sheet.Columns.Item(c).ColumnWidth) || 0; } catch (ignored0) {}
            }
            var absorberWidth = 0;
            try { absorberWidth = Number(sheet.Columns.Item(absorber).ColumnWidth) || 0; } catch (ignored1) {}
            var newIndex = absorber - countDeletedLessThan(ascending, absorber);
            if (newIndex < 1) newIndex = 1;
            plans.push({ newIndex: newIndex, width: absorberWidth + deletedWidth });
        }
        return plans;
    }

    function applyColumnWidthPlans(sheet, plans) {
        var applied = 0;
        for (var i = 0; i < plans.length; i += 1) {
            try {
                var targetWidth = plans[i].width;
                if (targetWidth > 255) targetWidth = 255;
                if (targetWidth < 0.1) targetWidth = 0.1;
                sheet.Columns.Item(plans[i].newIndex).ColumnWidth = targetWidth;
                applied += 1;
            } catch (ignored) {}
        }
        return applied;
    }

    function deleteColumnsOnSheet(workbook, sheetName, columns) {
        var sorted = uniqueSortedColumnsDescending(columns);
        var sheet = activateWorksheet(workbook, sheetName);
        var widthPlans = captureColumnWidthPlans(sheet, sorted);
        var bounds = getUsedBounds(sheet);
        var merges = collectAffectedColumnMerges(sheet, sorted, bounds);
        if (merges.length > 0) unmergeSavedAreas(sheet, merges);
        for (var i = 0; i < sorted.length; i += 1) {
            var address = columnToLetters(sorted[i]) + ":" + columnToLetters(sorted[i]);
            try { Application.Range(address).Delete(); }
            catch (error1) {
                try { Application.ActiveSheet.Range(address).Delete(); }
                catch (error2) { Application.ActiveSheet.Columns.Item(sorted[i]).Delete(); }
            }
        }
        if (merges.length > 0) restoreColumnMerges(sheet, merges, sorted);
        var expanded = applyColumnWidthPlans(sheet, widthPlans);
        try { alignPrintBottomOnSheet(sheet); } catch (ignored0) {}
        return { mergeCount: merges.length, expandedColumns: expanded };
    }

    function deleteColumnsByScope(mode, columns, sourceLabel) {
        var context = getTargetContext(false);
        var workbook = context.workbook;
        var originalSheetName = getOriginalActiveSheetName(workbook);
        var sheetNames = getScopeSheetNames(workbook, context.sheet, mode);
        var columnLabel = formatColumnList(columns);
        if (MsgBox("处理范围：" + (mode === "workbook" ? "整个工作簿（" + sheetNames.length + "张工作表）" : "当前工作表“" + context.sheet.Name + "”") + "\n将删除整列：" + columnLabel + "\n选择来源：" + sourceLabel + "\n\n删除后，每个连续删除区左侧的保留列会吸收被删列宽度，表格总宽度保持不变。\n\n是否继续？", JS_YES_NO + JS_QUESTION, "删除多余列") !== JS_RESULT_YES) return;
        createUndoPoint(workbook, sheetNames, "删除多余列");
        var completed = 0;
        var mergeCount = 0;
        var expandedCount = 0;
        var failed = [];
        for (var i = 0; i < sheetNames.length; i += 1) {
            updateProgress("删除多余列", i + 1, sheetNames.length);
            try {
                var result = deleteColumnsOnSheet(workbook, sheetNames[i], columns);
                mergeCount += result.mergeCount;
                expandedCount += result.expandedColumns;
                completed += 1;
            } catch (error) { failed.push("“" + sheetNames[i] + "”：" + (error && error.message ? error.message : String(error))); }
        }
        restoreActiveSheet(workbook, originalSheetName);
        var message = "已在" + completed + "张工作表中删除列 " + columnLabel + "。\n已自动扩宽" + expandedCount + "个相邻保留列。";
        if (mergeCount > 0) message += "\n已重建" + mergeCount + "个合并区域。";
        if (failed.length > 0) {
            message += "\n\n失败：\n" + failed.slice(0, 12).join("\n");
            MsgBox(message, JS_EXCLAMATION, "部分完成");
        } else MsgBox(message, JS_INFORMATION, "删除完成");
    }

    function deleteSelectedColumnsByScope(mode) {
        return deleteColumnsByScope(mode, getSelectedColumnNumbers(), "当前选区覆盖的列");
    }

    function promptDeleteColumnsByLetters() {
        var specification = InputBox("请输入需要删除的列号或范围。\n示例：C、C,F、H:J", "输入要删除的列", "");
        specification = normalizeText(specification);
        if (!specification) return;
        var mode = promptScope();
        if (!mode) return;
        return deleteColumnsByScope(mode, parseColumnSpec(specification), "手动输入“" + specification + "”");
    }

    function getActualContentBounds(sheet) {
        var used = getUsedBounds(sheet);
        var minRow = null;
        var minColumn = null;
        var maxRow = null;
        var maxColumn = null;
        for (var row = used.firstRow; row <= used.lastRow; row += 1) {
            for (var column = used.firstColumn; column <= used.lastColumn; column += 1) {
                var cell = sheet.Cells.Item(row, column);
                if (!isTopLeftOfMerge(cell)) continue;
                var hasContent = cellHasFormula(cell) || !isBlankValue(readMergedAwareValue(cell));
                if (!hasContent) continue;
                var top = row, left = column, bottom = row, right = column;
                try {
                    if (cell.MergeCells) {
                        var area = cell.MergeArea;
                        top = Number(area.Row);
                        left = Number(area.Column);
                        bottom = top + Number(area.Rows.Count) - 1;
                        right = left + Number(area.Columns.Count) - 1;
                    }
                } catch (e) {}
                if (minRow === null || top < minRow) minRow = top;
                if (minColumn === null || left < minColumn) minColumn = left;
                if (maxRow === null || bottom > maxRow) maxRow = bottom;
                if (maxColumn === null || right > maxColumn) maxColumn = right;
            }
        }
        if (minRow === null) return null;
        return { firstRow: minRow, firstColumn: minColumn, lastRow: maxRow, lastColumn: maxColumn };
    }

    function absoluteRangeAddress(bounds) {
        return "$" + columnToLetters(bounds.firstColumn) + "$" + bounds.firstRow + ":$" + columnToLetters(bounds.lastColumn) + "$" + bounds.lastRow;
    }

    function adjustPrintAreaByScope(mode) {
        var context = getTargetContext(false);
        var workbook = context.workbook;
        var originalSheetName = getOriginalActiveSheetName(workbook);
        var sheetNames = getScopeSheetNames(workbook, context.sheet, mode);
        createUndoPoint(workbook, sheetNames, "调整右侧分页线");
        var adjusted = [];
        var failed = [];
        for (var i = 0; i < sheetNames.length; i += 1) {
            updateProgress("调整右侧分页线", i + 1, sheetNames.length);
            var sheetName = sheetNames[i];
            try {
                var sheet = activateWorksheet(workbook, sheetName);
                /*
                 * 只调整横向缩放：一页宽，高度自动。
                 * 不修改 PrintArea，不 ResetAllPageBreaks，因此保留表格底部的横向分页和多张纵向页面。
                 */
                sheet.PageSetup.Zoom = false;
                sheet.PageSetup.FitToPagesWide = 1;
                try { sheet.PageSetup.FitToPagesTall = false; }
                catch (error1) { sheet.PageSetup.FitToPagesTall = 0; }
                try { sheet.DisplayPageBreaks = true; } catch (displayError) {}
                adjusted.push(sheetName);
            } catch (error) {
                failed.push("“" + sheetName + "”：" + (error && error.message ? error.message : String(error)));
            }
        }
        restoreActiveSheet(workbook, originalSheetName);
        var message = "已调整" + adjusted.length + "张工作表：仅将宽度缩放为1页，保留原打印区域、横向分页线和纵向多页结构。";
        if (failed.length) {
            message += "\n\n失败：\n" + failed.slice(0, 12).join("\n");
            MsgBox(message, JS_EXCLAMATION, "部分完成");
        } else {
            MsgBox(message, JS_INFORMATION, "页面宽度已调整");
        }
    }


    function enableManualPageBreakEditingByScope(
        mode
    ) {
        var context =
            getTargetContext(false);
        var workbook =
            context.workbook;
        var originalSheetName =
            getOriginalActiveSheetName(
                workbook
            );
        var sheetNames =
            getScopeSheetNames(
                workbook,
                context.sheet,
                mode
            );
        var completed = 0;
        var protectedSheets = [];
        var failures = [];

        for (
            var i = 0;
            i < sheetNames.length;
            i += 1
        ) {
            updateProgress(
                "启用手动拖动分页线",
                i + 1,
                sheetNames.length
            );

            try {
                var sheet =
                    activateWorksheet(
                        workbook,
                        sheetNames[i]
                    );

                try {
                    if (
                        sheet.ProtectContents ||
                        sheet.ProtectDrawingObjects ||
                        sheet.ProtectScenarios
                    ) {
                        protectedSheets.push(
                            sheetNames[i]
                        );
                    }
                } catch (ignored0) {}

                try {
                    sheet.PageSetup.Zoom = 100;
                } catch (ignored1) {}
                try {
                    sheet.PageSetup.FitToPagesWide = false;
                } catch (ignored2) {
                    try {
                        sheet.PageSetup.FitToPagesWide = 0;
                    } catch (ignored3) {}
                }
                try {
                    sheet.PageSetup.FitToPagesTall = false;
                } catch (ignored4) {
                    try {
                        sheet.PageSetup.FitToPagesTall = 0;
                    } catch (ignored5) {}
                }
                try {
                    sheet.DisplayPageBreaks = true;
                } catch (ignored6) {}
                try {
                    Application.ActiveWindow.View = 2;
                } catch (ignored7) {}

                completed += 1;
            } catch (error) {
                failures.push(
                    "“" + sheetNames[i] +
                    "”：" +
                    (
                        error && error.message
                            ? error.message
                            : String(error)
                    )
                );
            }
        }

        restoreActiveSheet(
            workbook,
            originalSheetName
        );
        try {
            Application.ActiveWindow.View = 2;
        } catch (ignored8) {}

        var message =
            "已将" + completed +
            "张工作表切换为可手动调整分页线的模式：" +
            "\n• 取消按页数缩放" +
            "\n• 打印缩放改为100%" +
            "\n• 打开分页预览" +
            "\n\n现在可直接拖动蓝色分页线。";

        if (protectedSheets.length) {
            message +=
                "\n\n以下工作表处于保护状态，若仍不能拖动，请先取消保护：\n" +
                protectedSheets
                    .slice(0, 12)
                    .join("、");
        }
        if (failures.length) {
            message +=
                "\n\n失败：\n" +
                failures
                    .slice(0, 12)
                    .join("\n");
            MsgBox(
                message,
                JS_EXCLAMATION,
                completed ? "部分完成" : "处理失败"
            );
        } else {
            MsgBox(
                message,
                protectedSheets.length
                    ? JS_EXCLAMATION
                    : JS_INFORMATION,
                "手动分页模式已启用"
            );
        }
    }

    function setAllWorksheetViews(viewValue, label) {
        var context = getTargetContext(true);
        var workbook = context.workbook;
        var originalSheetName = getOriginalActiveSheetName(workbook);
        var sheetNames = getScopeSheetNames(workbook, context.sheet, "workbook");
        var changed = 0;
        var failures = [];
        for (var i = 0; i < sheetNames.length; i += 1) {
            updateProgress(label, i + 1, sheetNames.length);
            try {
                activateWorksheet(workbook, sheetNames[i]);
                Application.ActiveWindow.View = viewValue;
                changed += 1;
            } catch (error) {
                failures.push(
                    "“" + sheetNames[i] + "”：" +
                    (error && error.message ? error.message : String(error))
                );
            }
        }
        restoreActiveSheet(workbook, originalSheetName);
        var message = "已将" + changed + "张工作表切换为" + label + "。";
        if (failures.length) {
            message += "\n失败：" + failures.length + "张。";
            MsgBox(message, JS_EXCLAMATION, "部分完成");
        } else {
            MsgBox(message, JS_INFORMATION, "视图切换完成");
        }
    }

    function getVisibleBorder(
        cell,
        borderIndex
    ) {
        try {
            var border =
                cell.Borders.Item(
                    borderIndex
                );
            var lineStyle =
                border.LineStyle;
            if (
                lineStyle === null ||
                lineStyle === undefined ||
                lineStyle === "" ||
                Number(lineStyle) === 0 ||
                Number(lineStyle) === -4142
            ) {
                return null;
            }

            var weight =
                Number(border.Weight);
            if (!isFinite(weight)) {
                weight = 2;
            }

            return {
                border: border,
                weight: weight
            };
        } catch (ignored) {
            return null;
        }
    }

    function cellHasAnyVisibleBorder(cell) {
        var indexes = [7, 8, 9, 10];
        for (
            var i = 0;
            i < indexes.length;
            i += 1
        ) {
            if (
                getVisibleBorder(
                    cell,
                    indexes[i]
                )
            ) {
                return true;
            }
        }
        return false;
    }

    function cellBelongsToTable(
        sheet,
        row,
        column,
        bounds,
        cache
    ) {
        if (
            row < bounds.firstRow ||
            row > bounds.lastRow ||
            column < bounds.firstColumn ||
            column > bounds.lastColumn
        ) {
            return false;
        }

        var key =
            row + ":" + column;
        if (
            cache.hasOwnProperty(key)
        ) {
            return cache[key];
        }

        var cell =
            sheet.Cells.Item(
                row,
                column
            );
        var belongs = false;

        try {
            if (cellHasFormula(cell)) {
                belongs = true;
            }
        } catch (ignored0) {}

        if (!belongs) {
            try {
                belongs =
                    !isBlankValue(
                        readMergedAwareValue(
                            cell
                        )
                    );
            } catch (ignored1) {}
        }

        if (!belongs) {
            try {
                belongs =
                    !!cell.MergeCells;
            } catch (ignored2) {}
        }

        if (!belongs) {
            belongs =
                cellHasAnyVisibleBorder(
                    cell
                );
        }

        cache[key] = belongs;
        return belongs;
    }

    function boundaryTouchesWideTitle(
        topCell,
        bottomCell,
        boundaryRow,
        tableWidth
    ) {
        var minimumWidth =
            Math.max(
                3,
                Math.floor(
                    tableWidth * 0.5
                )
            );

        var cells = [
            topCell,
            bottomCell
        ];
        for (
            var i = 0;
            i < cells.length;
            i += 1
        ) {
            var cell = cells[i];
            if (!cell) continue;
            try {
                if (!cell.MergeCells) {
                    continue;
                }
                var area =
                    cell.MergeArea;
                var width =
                    Number(
                        area.Columns.Count
                    );
                var top =
                    Number(area.Row);
                var bottom =
                    top +
                    Number(
                        area.Rows.Count
                    );
                if (
                    width >= minimumWidth &&
                    (
                        boundaryRow === top ||
                        boundaryRow === bottom
                    )
                ) {
                    return true;
                }
            } catch (ignored) {}
        }
        return false;
    }

    function borderWeightRank(weight) {
        var value = Number(weight);
        if (value === 1) return 1;
        if (value === 2) return 2;
        if (value === -4138) return 3;
        if (value === 4) return 4;
        if (!isFinite(value)) return 2;
        return value;
    }

    function chooseFrequentBorderWeight(
        frequency,
        preferThicker
    ) {
        var best = null;
        var bestCount = -1;
        var bestRank =
            preferThicker
                ? -Infinity
                : Infinity;

        for (
            var key in frequency
        ) {
            if (
                !frequency
                    .hasOwnProperty(key)
            ) {
                continue;
            }

            var count =
                frequency[key];
            var weight =
                Number(key);
            var rank =
                borderWeightRank(
                    weight
                );

            if (
                count > bestCount ||
                (
                    count === bestCount &&
                    (
                        preferThicker
                            ? rank > bestRank
                            : rank < bestRank
                    )
                )
            ) {
                best = weight;
                bestCount = count;
                bestRank = rank;
            }
        }
        return best;
    }

    function addWeightFrequency(
        frequency,
        weight
    ) {
        var key = String(
            Number(weight)
        );
        frequency[key] =
            (frequency[key] || 0) + 1;
    }

    function getMergedRangeDescriptor(cell) {
        try {
            if (!cell.MergeCells) return null;
            var area = cell.MergeArea;
            return {
                area: area,
                firstRow: Number(area.Row),
                firstColumn: Number(area.Column),
                lastRow:
                    Number(area.Row) +
                    Number(area.Rows.Count) - 1,
                lastColumn:
                    Number(area.Column) +
                    Number(area.Columns.Count) - 1,
                rowCount: Number(area.Rows.Count),
                columnCount:
                    Number(area.Columns.Count)
            };
        } catch (ignored) {
            return null;
        }
    }

    function mergedEdgeIsOuter(
        sheet,
        descriptor,
        borderIndex,
        bounds,
        memberCache
    ) {
        var row;
        var column;

        if (borderIndex === 7) {
            column = descriptor.firstColumn - 1;
            if (column < bounds.firstColumn) return true;
            for (
                row = descriptor.firstRow;
                row <= descriptor.lastRow;
                row += 1
            ) {
                if (
                    !cellBelongsToTable(
                        sheet,
                        row,
                        column,
                        bounds,
                        memberCache
                    )
                ) return true;
            }
            return false;
        }

        if (borderIndex === 10) {
            column = descriptor.lastColumn + 1;
            if (column > bounds.lastColumn) return true;
            for (
                row = descriptor.firstRow;
                row <= descriptor.lastRow;
                row += 1
            ) {
                if (
                    !cellBelongsToTable(
                        sheet,
                        row,
                        column,
                        bounds,
                        memberCache
                    )
                ) return true;
            }
            return false;
        }

        if (borderIndex === 8) {
            row = descriptor.firstRow - 1;
            if (row < bounds.firstRow) return true;
            for (
                column = descriptor.firstColumn;
                column <= descriptor.lastColumn;
                column += 1
            ) {
                if (
                    !cellBelongsToTable(
                        sheet,
                        row,
                        column,
                        bounds,
                        memberCache
                    )
                ) return true;
            }
            return false;
        }

        row = descriptor.lastRow + 1;
        if (row > bounds.lastRow) return true;
        for (
            column = descriptor.firstColumn;
            column <= descriptor.lastColumn;
            column += 1
        ) {
            if (
                !cellBelongsToTable(
                    sheet,
                    row,
                    column,
                    bounds,
                    memberCache
                )
            ) return true;
        }
        return false;
    }

    function collectMergedEdgeBorderRefs(
        sheet,
        descriptor,
        borderIndex
    ) {
        var refs = [];
        var seen = {};

        function addBorder(cell, index) {
            if (!cell) return;
            var visible = getVisibleBorder(
                cell,
                index
            );
            if (!visible) return;
            var key = String(index) + ":" +
                Number(cell.Row) + ":" +
                Number(cell.Column);
            if (seen[key]) return;
            seen[key] = true;
            refs.push(visible);
        }

        try {
            var areaVisible = getVisibleBorder(
                descriptor.area,
                borderIndex
            );
            if (areaVisible) {
                refs.push(areaVisible);
            }
        } catch (ignored0) {}

        var row;
        var column;
        if (borderIndex === 7 || borderIndex === 10) {
            column = borderIndex === 7
                ? descriptor.firstColumn
                : descriptor.lastColumn;
            for (
                row = descriptor.firstRow;
                row <= descriptor.lastRow;
                row += 1
            ) {
                addBorder(
                    sheet.Cells.Item(row, column),
                    borderIndex
                );
            }
        } else {
            row = borderIndex === 8
                ? descriptor.firstRow
                : descriptor.lastRow;
            for (
                column = descriptor.firstColumn;
                column <= descriptor.lastColumn;
                column += 1
            ) {
                addBorder(
                    sheet.Cells.Item(row, column),
                    borderIndex
                );
            }
        }
        return refs;
    }

    function countHorizontalMergedGroupsInRow(
        sheet,
        rowNumber,
        bounds
    ) {
        var seen = {};
        var count = 0;

        for (
            var column = bounds.firstColumn;
            column <= bounds.lastColumn;
            column += 1
        ) {
            var descriptor =
                getMergedRangeDescriptor(
                    sheet.Cells.Item(
                        rowNumber,
                        column
                    )
                );
            if (
                !descriptor ||
                descriptor.firstRow !==
                    rowNumber ||
                descriptor.columnCount < 2
            ) {
                continue;
            }

            var key =
                descriptor.firstRow + ":" +
                descriptor.firstColumn + ":" +
                descriptor.lastRow + ":" +
                descriptor.lastColumn;
            if (seen[key]) continue;
            seen[key] = true;

            var text = normalizeText(
                readMergedAwareValue(
                    sheet.Cells.Item(
                        descriptor.firstRow,
                        descriptor.firstColumn
                    )
                )
            );
            if (text) count += 1;
        }
        return count;
    }

    function rowStartsStructuredHeader(
        sheet,
        rowNumber,
        bounds
    ) {
        if (
            rowNumber < bounds.firstRow ||
            rowNumber > bounds.lastRow
        ) {
            return false;
        }

        if (
            likelyRepeatedTitle(
                sheet,
                rowNumber,
                bounds
            )
        ) {
            return true;
        }

        return (
            countHorizontalMergedGroupsInRow(
                sheet,
                rowNumber,
                bounds
            ) >= 2
        );
    }

    function mergedDescriptorIsHeaderGroup(
        sheet,
        descriptor,
        bounds
    ) {
        var text = normalizeText(
            readMergedAwareValue(
                sheet.Cells.Item(
                    descriptor.firstRow,
                    descriptor.firstColumn
                )
            )
        );
        if (!text) return false;

        if (
            likelyRepeatedTitle(
                sheet,
                descriptor.firstRow,
                bounds
            )
        ) {
            return true;
        }

        if (
            descriptor.columnCount >= 2 &&
            countHorizontalMergedGroupsInRow(
                sheet,
                descriptor.firstRow,
                bounds
            ) >= 2
        ) {
            return true;
        }

        return false;
    }

    function normalizeMergedRangeBorders(
        sheet,
        bounds,
        memberCache,
        outerWeight,
        innerWeight
    ) {
        var seen = {};
        var changed = 0;
        var failed = 0;
        var checked = 0;
        var tableWidth =
            bounds.lastColumn -
            bounds.firstColumn + 1;

        for (
            var row = bounds.firstRow;
            row <= bounds.lastRow;
            row += 1
        ) {
            for (
                var column = bounds.firstColumn;
                column <= bounds.lastColumn;
                column += 1
            ) {
                var descriptor =
                    getMergedRangeDescriptor(
                        sheet.Cells.Item(
                            row,
                            column
                        )
                    );
                if (!descriptor) continue;

                var key =
                    descriptor.firstRow + ":" +
                    descriptor.firstColumn + ":" +
                    descriptor.lastRow + ":" +
                    descriptor.lastColumn;
                if (seen.hasOwnProperty(key)) {
                    continue;
                }
                seen[key] = true;

                var indexes = [7, 8, 9, 10];
                for (
                    var i = 0;
                    i < indexes.length;
                    i += 1
                ) {
                    var borderIndex = indexes[i];
                    var refs =
                        collectMergedEdgeBorderRefs(
                            sheet,
                            descriptor,
                            borderIndex
                        );
                    if (!refs.length) continue;

                    var isOuter = mergedEdgeIsOuter(
                        sheet,
                        descriptor,
                        borderIndex,
                        bounds,
                        memberCache
                    );

                    if (
                        mergedDescriptorIsHeaderGroup(
                            sheet,
                            descriptor,
                            bounds
                        ) ||
                        (
                            (borderIndex === 8 ||
                             borderIndex === 9) &&
                            descriptor.columnCount >=
                                Math.max(
                                    3,
                                    Math.floor(
                                        tableWidth * 0.5
                                    )
                                )
                        )
                    ) {
                        isOuter = true;
                    }

                    var targetWeight = isOuter
                        ? outerWeight
                        : innerWeight;

                    for (
                        var r = 0;
                        r < refs.length;
                        r += 1
                    ) {
                        checked += 1;
                        try {
                            var beforeStyle = Number(
                                refs[r].border.LineStyle
                            );
                            var beforeWeight = Number(
                                refs[r].border.Weight
                            );
                            refs[r].border.LineStyle = 1;
                            refs[r].border.Weight =
                                targetWeight;

                            if (
                                Number(
                                    refs[r].border.Weight
                                ) ===
                                Number(targetWeight)
                            ) {
                                if (
                                    beforeStyle !== 1 ||
                                    beforeWeight !==
                                        Number(targetWeight)
                                ) {
                                    changed += 1;
                                }
                            } else {
                                failed += 1;
                            }
                        } catch (error) {
                            failed += 1;
                        }
                    }
                }
            }
        }

        return {
            changed: changed,
            failed: failed,
            checked: checked
        };
    }


    function getRowVisibleBorderSpan(
        sheet,
        rowNumber,
        bounds
    ) {
        var firstColumn = 0;
        var lastColumn = 0;

        for (
            var column = bounds.firstColumn;
            column <= bounds.lastColumn;
            column += 1
        ) {
            if (
                !cellHasAnyVisibleBorder(
                    sheet.Cells.Item(
                        rowNumber,
                        column
                    )
                )
            ) {
                continue;
            }

            if (!firstColumn) {
                firstColumn = column;
            }
            lastColumn = column;
        }

        return {
            hasBorder: firstColumn > 0,
            firstColumn: firstColumn,
            lastColumn: lastColumn
        };
    }

    function rowIsWideTableTitle(
        sheet,
        rowNumber,
        bounds
    ) {
        var tableWidth =
            bounds.lastColumn -
            bounds.firstColumn + 1;
        var minimumWidth =
            Math.max(
                3,
                Math.floor(
                    tableWidth * 0.55
                )
            );

        for (
            var column = bounds.firstColumn;
            column <= bounds.lastColumn;
            column += 1
        ) {
            var cell =
                sheet.Cells.Item(
                    rowNumber,
                    column
                );
            var descriptor =
                getMergedRangeDescriptor(
                    cell
                );

            if (
                !descriptor ||
                descriptor.firstRow !==
                    rowNumber ||
                descriptor.firstColumn !==
                    column ||
                descriptor.columnCount <
                    minimumWidth
            ) {
                continue;
            }

            var text =
                normalizeText(
                    readMergedAwareValue(
                        cell
                    )
                );
            if (
                text &&
                text.length >= 4
            ) {
                return true;
            }
        }

        return false;
    }

    function detectBorderTableBlocks(
        sheet,
        bounds
    ) {
        var blocks = [];
        var current = null;

        function closeCurrent() {
            if (!current) return;
            blocks.push(current);
            current = null;
        }

        for (
            var row = bounds.firstRow;
            row <= bounds.lastRow;
            row += 1
        ) {
            var span =
                getRowVisibleBorderSpan(
                    sheet,
                    row,
                    bounds
                );

            if (!span.hasBorder) {
                closeCurrent();
                continue;
            }

            /*
             * 连续放置多张表且中间没有空行时，
             * 后续宽合并表名作为新表格块起点。
             */
            if (
                current &&
                row > current.firstRow &&
                rowIsWideTableTitle(
                    sheet,
                    row,
                    bounds
                )
            ) {
                closeCurrent();
            }

            if (!current) {
                current = {
                    firstRow: row,
                    lastRow: row,
                    firstColumn:
                        span.firstColumn,
                    lastColumn:
                        span.lastColumn
                };
            } else {
                current.lastRow = row;
                current.firstColumn =
                    Math.min(
                        current.firstColumn,
                        span.firstColumn
                    );
                current.lastColumn =
                    Math.max(
                        current.lastColumn,
                        span.lastColumn
                    );
            }
        }

        closeCurrent();
        return blocks;
    }

    function findBorderTableBlock(
        blocks,
        row,
        column
    ) {
        for (
            var i = 0;
            i < blocks.length;
            i += 1
        ) {
            var block = blocks[i];
            if (
                row >= block.firstRow &&
                row <= block.lastRow &&
                column >=
                    block.firstColumn &&
                column <=
                    block.lastColumn
            ) {
                return block;
            }
        }
        return null;
    }

    function targetBorderWeight(
        block,
        row,
        column,
        borderIndex,
        outerWeight,
        innerWeight
    ) {
        if (!block) return innerWeight;

        if (
            borderIndex === 7 &&
            column === block.firstColumn
        ) {
            return outerWeight;
        }
        if (
            borderIndex === 10 &&
            column === block.lastColumn
        ) {
            return outerWeight;
        }
        if (
            borderIndex === 8 &&
            row === block.firstRow
        ) {
            return outerWeight;
        }
        if (
            borderIndex === 9 &&
            row === block.lastRow
        ) {
            return outerWeight;
        }

        return innerWeight;
    }

    function writeVisibleBorder(
        borderRef,
        targetWeight
    ) {
        var beforeStyle = null;
        var beforeWeight = null;
        var succeeded = false;

        for (
            var attempt = 0;
            attempt < BORDER_CORRECTIVE_PASSES;
            attempt += 1
        ) {
            try {
                if (attempt === 0) {
                    beforeStyle =
                        Number(
                            borderRef.border.LineStyle
                        );
                    beforeWeight =
                        Number(
                            borderRef.border.Weight
                        );
                }
                borderRef.border.LineStyle = 1;
                borderRef.border.Weight =
                    targetWeight;
                if (
                    Number(
                        borderRef.border.LineStyle
                    ) === 1 &&
                    Number(
                        borderRef.border.Weight
                    ) === Number(targetWeight)
                ) {
                    succeeded = true;
                    break;
                }
            } catch (ignored) {}
        }

        if (!succeeded) {
            return {
                checked: 1,
                changed: 0,
                failed: 1
            };
        }
        return {
            checked: 1,
            changed:
                (
                    beforeStyle !== 1 ||
                    beforeWeight !== Number(targetWeight)
                )
                    ? 1
                    : 0,
            failed: 0
        };
    }

    function applyBoundaryRefs(
        refs,
        targetWeight
    ) {
        var result = {
            checked: 0,
            changed: 0,
            failed: 0
        };
        for (
            var i = 0;
            i < refs.length;
            i += 1
        ) {
            var current =
                writeVisibleBorder(
                    refs[i],
                    targetWeight
                );
            result.checked += current.checked;
            result.changed += current.changed;
            result.failed += current.failed;
        }
        return result;
    }

    function collectHorizontalBoundaryRefs(
        sheet,
        block,
        boundaryRow,
        column
    ) {
        var refs = [];
        if (boundaryRow > block.firstRow) {
            var top =
                getVisibleBorder(
                    sheet.Cells.Item(
                        boundaryRow - 1,
                        column
                    ),
                    9
                );
            if (top) refs.push(top);
        }
        if (boundaryRow <= block.lastRow) {
            var bottom =
                getVisibleBorder(
                    sheet.Cells.Item(
                        boundaryRow,
                        column
                    ),
                    8
                );
            if (bottom) refs.push(bottom);
        }
        return refs;
    }

    function collectVerticalBoundaryRefs(
        sheet,
        block,
        row,
        boundaryColumn
    ) {
        var refs = [];
        if (boundaryColumn > block.firstColumn) {
            var left =
                getVisibleBorder(
                    sheet.Cells.Item(
                        row,
                        boundaryColumn - 1
                    ),
                    10
                );
            if (left) refs.push(left);
        }
        if (boundaryColumn <= block.lastColumn) {
            var right =
                getVisibleBorder(
                    sheet.Cells.Item(
                        row,
                        boundaryColumn
                    ),
                    7
                );
            if (right) refs.push(right);
        }
        return refs;
    }

    function normalizeOrdinaryBoundaries(
        sheet,
        blocks,
        outerWeight,
        innerWeight
    ) {
        var totals = {
            checked: 0,
            changed: 0,
            failed: 0
        };

        for (
            var b = 0;
            b < blocks.length;
            b += 1
        ) {
            var block = blocks[b];

            for (
                var boundaryRow = block.firstRow;
                boundaryRow <= block.lastRow + 1;
                boundaryRow += 1
            ) {
                var horizontalTarget =
                    (
                        boundaryRow === block.firstRow ||
                        boundaryRow === block.lastRow + 1
                    )
                        ? outerWeight
                        : innerWeight;

                for (
                    var column = block.firstColumn;
                    column <= block.lastColumn;
                    column += 1
                ) {
                    var horizontal =
                        applyBoundaryRefs(
                            collectHorizontalBoundaryRefs(
                                sheet,
                                block,
                                boundaryRow,
                                column
                            ),
                            horizontalTarget
                        );
                    totals.checked += horizontal.checked;
                    totals.changed += horizontal.changed;
                    totals.failed += horizontal.failed;
                }
            }

            for (
                var boundaryColumn = block.firstColumn;
                boundaryColumn <= block.lastColumn + 1;
                boundaryColumn += 1
            ) {
                var verticalTarget =
                    (
                        boundaryColumn === block.firstColumn ||
                        boundaryColumn === block.lastColumn + 1
                    )
                        ? outerWeight
                        : innerWeight;

                for (
                    var row = block.firstRow;
                    row <= block.lastRow;
                    row += 1
                ) {
                    var vertical =
                        applyBoundaryRefs(
                            collectVerticalBoundaryRefs(
                                sheet,
                                block,
                                row,
                                boundaryColumn
                            ),
                            verticalTarget
                        );
                    totals.checked += vertical.checked;
                    totals.changed += vertical.changed;
                    totals.failed += vertical.failed;
                }
            }
        }
        return totals;
    }

    function normalizeMergedAreasByBlocks(
        sheet,
        bounds,
        blocks,
        outerWeight,
        innerWeight
    ) {
        var seen = {};
        var totals = {
            checked: 0,
            changed: 0,
            failed: 0
        };
        var indexes = [7, 8, 9, 10];

        for (
            var row = bounds.firstRow;
            row <= bounds.lastRow;
            row += 1
        ) {
            for (
                var column = bounds.firstColumn;
                column <= bounds.lastColumn;
                column += 1
            ) {
                var descriptor =
                    getMergedRangeDescriptor(
                        sheet.Cells.Item(
                            row,
                            column
                        )
                    );
                if (!descriptor) continue;

                var key =
                    descriptor.firstRow + ":" +
                    descriptor.firstColumn + ":" +
                    descriptor.lastRow + ":" +
                    descriptor.lastColumn;
                if (seen[key]) continue;
                seen[key] = true;

                var block =
                    findBorderTableBlock(
                        blocks,
                        descriptor.firstRow,
                        descriptor.firstColumn
                    );
                if (!block) continue;

                for (
                    var i = 0;
                    i < indexes.length;
                    i += 1
                ) {
                    var borderIndex = indexes[i];
                    var target = innerWeight;
                    if (
                        borderIndex === 7 &&
                        descriptor.firstColumn === block.firstColumn
                    ) {
                        target = outerWeight;
                    } else if (
                        borderIndex === 10 &&
                        descriptor.lastColumn === block.lastColumn
                    ) {
                        target = outerWeight;
                    } else if (
                        borderIndex === 8 &&
                        descriptor.firstRow === block.firstRow
                    ) {
                        target = outerWeight;
                    } else if (
                        borderIndex === 9 &&
                        descriptor.lastRow === block.lastRow
                    ) {
                        target = outerWeight;
                    }

                    var current =
                        applyBoundaryRefs(
                            collectMergedEdgeBorderRefs(
                                sheet,
                                descriptor,
                                borderIndex
                            ),
                            target
                        );
                    totals.checked += current.checked;
                    totals.changed += current.changed;
                    totals.failed += current.failed;
                }
            }
        }
        return totals;
    }

    function verifyBorderBlocks(
        sheet,
        blocks,
        outerWeight,
        innerWeight
    ) {
        var checked = 0;
        var failed = 0;

        for (
            var b = 0;
            b < blocks.length;
            b += 1
        ) {
            var block = blocks[b];

            for (
                var boundaryRow = block.firstRow;
                boundaryRow <= block.lastRow + 1;
                boundaryRow += 1
            ) {
                var horizontalTarget =
                    (
                        boundaryRow === block.firstRow ||
                        boundaryRow === block.lastRow + 1
                    )
                        ? outerWeight
                        : innerWeight;
                for (
                    var column = block.firstColumn;
                    column <= block.lastColumn;
                    column += 1
                ) {
                    var hRefs =
                        collectHorizontalBoundaryRefs(
                            sheet,
                            block,
                            boundaryRow,
                            column
                        );
                    for (
                        var h = 0;
                        h < hRefs.length;
                        h += 1
                    ) {
                        checked += 1;
                        try {
                            if (
                                Number(hRefs[h].border.LineStyle) !== 1 ||
                                Number(hRefs[h].border.Weight) !== Number(horizontalTarget)
                            ) {
                                failed += 1;
                            }
                        } catch (ignored0) {
                            failed += 1;
                        }
                    }
                }
            }

            for (
                var boundaryColumn = block.firstColumn;
                boundaryColumn <= block.lastColumn + 1;
                boundaryColumn += 1
            ) {
                var verticalTarget =
                    (
                        boundaryColumn === block.firstColumn ||
                        boundaryColumn === block.lastColumn + 1
                    )
                        ? outerWeight
                        : innerWeight;
                for (
                    var row = block.firstRow;
                    row <= block.lastRow;
                    row += 1
                ) {
                    var vRefs =
                        collectVerticalBoundaryRefs(
                            sheet,
                            block,
                            row,
                            boundaryColumn
                        );
                    for (
                        var v = 0;
                        v < vRefs.length;
                        v += 1
                    ) {
                        checked += 1;
                        try {
                            if (
                                Number(vRefs[v].border.LineStyle) !== 1 ||
                                Number(vRefs[v].border.Weight) !== Number(verticalTarget)
                            ) {
                                failed += 1;
                            }
                        } catch (ignored1) {
                            failed += 1;
                        }
                    }
                }
            }
        }
        return {
            checked: checked,
            failed: failed
        };
    }

    function normalizeTableBordersOnSheet(
        sheet
    ) {
        var actual =
            getActualContentBounds(
                sheet
            );
        if (!actual) {
            return {
                changed: 0,
                checkedMerged: 0,
                failedMerged: 0,
                verifyFailed: 0,
                tableBlocks: 0,
                outerSegments: 0,
                innerSegments: 0,
                outerWeight: -4138,
                innerWeight: 2
            };
        }

        var rowCount =
            actual.lastRow -
            actual.firstRow + 1;
        var columnCount =
            actual.lastColumn -
            actual.firstColumn + 1;
        var cellCount =
            rowCount * columnCount;
        if (
            cellCount >
            MAX_BORDER_SCAN_CELLS
        ) {
            throw new Error(
                "表格范围约" +
                cellCount +
                "个单元格，超过线宽检查上限" +
                MAX_BORDER_SCAN_CELLS +
                "。请缩小工作表范围后再试。"
            );
        }

        var blocks =
            detectBorderTableBlocks(
                sheet,
                actual
            );
        var outerWeight = -4138;
        var innerWeight = 2;
        var checked = 0;
        var changed = 0;
        var failed = 0;

        /*
         * 同一共享边界的两侧一起写入；再处理合并区域；
         * 最后重复边界修正。这样可消除左格右边和右格左边
         * 各自保存不同粗细造成的“局部概率失败”。
         */
        for (
            var pass = 0;
            pass < BORDER_CORRECTIVE_PASSES;
            pass += 1
        ) {
            var ordinary =
                normalizeOrdinaryBoundaries(
                    sheet,
                    blocks,
                    outerWeight,
                    innerWeight
                );
            checked += ordinary.checked;
            changed += ordinary.changed;
            failed += ordinary.failed;

            var merged =
                normalizeMergedAreasByBlocks(
                    sheet,
                    actual,
                    blocks,
                    outerWeight,
                    innerWeight
                );
            checked += merged.checked;
            changed += merged.changed;
            failed += merged.failed;

            var verifyNow =
                verifyBorderBlocks(
                    sheet,
                    blocks,
                    outerWeight,
                    innerWeight
                );
            if (!verifyNow.failed) break;
        }

        var verify =
            verifyBorderBlocks(
                sheet,
                blocks,
                outerWeight,
                innerWeight
            );

        return {
            changed: changed,
            checkedMerged: checked,
            failedMerged: failed,
            verifyFailed: verify.failed,
            tableBlocks: blocks.length,
            outerSegments: 0,
            innerSegments: 0,
            outerWeight: outerWeight,
            innerWeight: innerWeight
        };
    }


    function borderWeightLabel(weight) {
        var value = Number(weight);
        if (value === 1) return "极细线";
        if (value === 2) return "细线";
        if (value === -4138) {
            return "中等线";
        }
        if (value === 4) return "粗线";
        return String(value);
    }

    function normalizeTableBordersByScope(
        mode
    ) {
        var context =
            getTargetContext(false);
        var workbook =
            context.workbook;
        var originalSheetName =
            getOriginalActiveSheetName(
                workbook
            );
        var sheetNames =
            getScopeSheetNames(
                workbook,
                context.sheet,
                mode
            );

        if (
            MsgBox(
                "处理范围：" +
                (
                    mode === "workbook"
                        ? "整个工作簿"
                        : "当前工作表"
                ) +
                "\n将先识别连续表格块：" +
                "\n• 每个表格块最外一圈统一为中等实线" +
                "\n• 表格块内部全部已有线统一为细实线" +
                "\n• 原本无线的位置不新增边框" +
                "\n\n是否继续？",
                JS_YES_NO + JS_QUESTION,
                "统一外框/内部线宽"
            ) !== JS_RESULT_YES
        ) {
            return;
        }

        var changed = 0;
        var completed = 0;
        var checked = 0;
        var failed = 0;
        var verifyFailed = 0;
        var tableBlocks = 0;
        var failures = [];

        for (
            var i = 0;
            i < sheetNames.length;
            i += 1
        ) {
            updateProgress(
                "统一外框/内部线宽",
                i + 1,
                sheetNames.length
            );

            try {
                var sheet =
                    activateWorksheet(
                        workbook,
                        sheetNames[i]
                    );
                var result =
                    normalizeTableBordersOnSheet(
                        sheet
                    );

                changed +=
                    Number(
                        result.changed || 0
                    );
                checked +=
                    Number(
                        result.checkedMerged || 0
                    );
                failed +=
                    Number(
                        result.failedMerged || 0
                    );
                verifyFailed +=
                    Number(
                        result.verifyFailed || 0
                    );
                tableBlocks +=
                    Number(
                        result.tableBlocks || 0
                    );
                completed += 1;
            } catch (error) {
                failures.push(
                    "“" + sheetNames[i] +
                    "”：" +
                    (
                        error &&
                        error.message
                            ? error.message
                            : String(error)
                    )
                );
            }
        }

        restoreActiveSheet(
            workbook,
            originalSheetName
        );

        var message =
            "已处理" + completed +
            "张工作表，识别" +
            tableBlocks +
            "个连续表格块。" +
            "\n调整已有边框：" +
            changed + "条。" +
            "\n外圈：中等实线；内部：细实线。" +
            "\n边框写入检查：" +
            checked + "条；写入失败：" +
            failed + "条。" +
            "\n最终粗细核验不一致：" +
            verifyFailed + "条。" +
            "\n没有边框的位置未新增线。";

        if (
            failures.length ||
            failed ||
            verifyFailed
        ) {
            if (failures.length) {
                message +=
                    "\n工作表失败：" +
                    failures.length + "张。";
            }
            MsgBox(
                message,
                JS_EXCLAMATION,
                completed
                    ? "部分完成"
                    : "处理失败"
            );
        } else {
            MsgBox(
                message,
                JS_INFORMATION,
                "线宽统一完成"
            );
        }
    }


    function getCellDisplayArea(
        cell
    ) {
        try {
            if (cell.MergeCells) {
                return cell.MergeArea;
            }
        } catch (ignored) {}
        return cell;
    }

    function getCellFontSize(
        cell
    ) {
        var size = 10;
        try {
            var value =
                Number(cell.Font.Size);
            if (
                isFinite(value) &&
                value > 0
            ) {
                size = value;
            }
        } catch (ignored) {}
        return size;
    }

    function getTextVisualUnits(
        text
    ) {
        var value = String(text || "");
        var units = 0;

        for (
            var i = 0;
            i < value.length;
            i += 1
        ) {
            var character =
                value.charAt(i);
            var code =
                value.charCodeAt(i);

            if (
                character === "\n" ||
                character === "\r"
            ) {
                continue;
            }

            if (/\s/.test(character)) {
                units += 0.32;
            } else if (
                code >= 0x2e80 ||
                code >= 0xff00
            ) {
                units += 1.0;
            } else if (
                /[A-Z]/.test(character)
            ) {
                units += 0.68;
            } else if (
                /[a-z0-9]/.test(
                    character
                )
            ) {
                units += 0.56;
            } else {
                units += 0.52;
            }
        }

        return units;
    }

    function estimateWrappedLines(
        text,
        availableWidth,
        fontSize
    ) {
        var paragraphs =
            String(text || "")
                .replace(/\r/g, "")
                .split("\n");
        var total = 0;
        var widthPerUnit =
            Math.max(
                1,
                fontSize * 0.95
            );

        for (
            var i = 0;
            i < paragraphs.length;
            i += 1
        ) {
            var requiredWidth =
                getTextVisualUnits(
                    paragraphs[i]
                ) * widthPerUnit;
            total += Math.max(
                1,
                Math.ceil(
                    requiredWidth /
                    Math.max(
                        1,
                        availableWidth
                    )
                )
            );
        }

        return total;
    }

    function getPageBreakLocations(
        sheet,
        horizontal
    ) {
        var result = [];
        try {
            var collection =
                horizontal
                    ? sheet.HPageBreaks
                    : sheet.VPageBreaks;
            for (
                var i = 1;
                i <=
                    Number(
                        collection.Count
                    );
                i += 1
            ) {
                var item =
                    collection.Item(i);
                var location =
                    horizontal
                        ? Number(
                            item.Location.Row
                          )
                        : Number(
                            item.Location.Column
                          );
                if (isFinite(location)) {
                    result.push(location);
                }
            }
        } catch (ignored) {}

        result.sort(function (a, b) {
            return a - b;
        });
        return result;
    }

    function mergeCrossesBreak(
        descriptor,
        breakLocations,
        horizontal
    ) {
        for (
            var i = 0;
            i < breakLocations.length;
            i += 1
        ) {
            var location =
                breakLocations[i];

            if (horizontal) {
                if (
                    location >
                        descriptor.firstRow &&
                    location <=
                        descriptor.lastRow
                ) {
                    return true;
                }
            } else {
                if (
                    location >
                        descriptor.firstColumn &&
                    location <=
                        descriptor.lastColumn
                ) {
                    return true;
                }
            }
        }
        return false;
    }

    function addPdfIssue(
        issues,
        sheetName,
        address,
        reason,
        severity
    ) {
        if (
            issues.length >=
            MAX_PDF_ISSUES
        ) {
            return;
        }

        issues.push({
            sheetName:
                sheetName,
            address:
                address,
            reason:
                reason,
            severity:
                severity || "风险"
        });
    }

    function normalizeDisplayCompareText(
        value
    ) {
        return normalizeText(value)
            .replace(/\s+/g, "")
            .replace(/[，,]/g, "")
            .toLowerCase();
    }

    function isTextLikeValue(
        value
    ) {
        return (
            typeof value === "string" &&
            !/^[-+]?\d+(?:[.,]\d+)?$/.test(
                normalizeText(value)
            )
        );
    }

    function inspectPdfTextCell(
        sheet,
        cell,
        descriptor,
        issues,
        printArea
    ) {
        var raw =
            readMergedAwareValue(
                cell
            );
        var text = normalizeText(raw);
        if (!text) return;

        var displayText = "";
        try {
            displayText =
                String(cell.Text || "");
        } catch (ignored0) {}

        var address =
            descriptor
                ? columnToLetters(
                    descriptor.firstColumn
                  ) +
                  descriptor.firstRow +
                  ":" +
                  columnToLetters(
                    descriptor.lastColumn
                  ) +
                  descriptor.lastRow
                : columnToLetters(
                    Number(cell.Column)
                  ) +
                  Number(cell.Row);

        if (/#{3,}/.test(displayText)) {
            addPdfIssue(
                issues,
                String(sheet.Name),
                address,
                "当前显示为###，列宽不足或日期/数值无法完整显示",
                "明确"
            );
            return;
        }

        /*
         * WPS在部分情况下会让cell.Text只返回当前可见部分。
         * 文本原值与显示文本明显不同，直接标记为截断风险。
         */
        if (
            isTextLikeValue(raw) &&
            displayText &&
            normalizeDisplayCompareText(displayText) !==
                normalizeDisplayCompareText(raw) &&
            normalizeDisplayCompareText(displayText).length <
                normalizeDisplayCompareText(raw).length
        ) {
            addPdfIssue(
                issues,
                String(sheet.Name),
                address,
                "当前显示文字比单元格原值短，已出现或接近出现半字/截字",
                "明确"
            );
        }

        var area =
            getCellDisplayArea(
                cell
            );
        var width = 0;
        var height = 0;
        try {
            width = Number(area.Width);
        } catch (ignored1) {}
        try {
            height = Number(area.Height);
        } catch (ignored2) {}
        if (
            !isFinite(width) || width <= 0 ||
            !isFinite(height) || height <= 0
        ) {
            return;
        }

        var fontSize =
            getCellFontSize(
                cell
            );
        var wrap = false;
        var shrink = false;
        var orientation = 0;
        var alignment = 0;
        try {
            wrap = !!area.WrapText;
        } catch (ignored3) {}
        try {
            shrink = !!area.ShrinkToFit;
        } catch (ignored4) {}
        try {
            orientation =
                Number(area.Orientation) || 0;
        } catch (ignored5) {}
        try {
            alignment =
                Number(area.HorizontalAlignment) || 0;
        } catch (ignored6) {}
        if (orientation !== 0) return;

        var availableWidth =
            Math.max(4, width - 7);
        var atRightPrintEdge = false;
        var atBottomPrintEdge = false;
        if (printArea) {
            var lastColumn =
                descriptor
                    ? descriptor.lastColumn
                    : Number(cell.Column);
            var lastRow =
                descriptor
                    ? descriptor.lastRow
                    : Number(cell.Row);
            atRightPrintEdge =
                lastColumn ===
                Number(printArea.lastColumn);
            atBottomPrintEdge =
                lastRow ===
                Number(printArea.lastRow);
        }

        if (wrap) {
            var lineCount =
                estimateWrappedLines(
                    String(raw),
                    availableWidth * 0.92,
                    fontSize * 1.08
                );
            var requiredHeight =
                lineCount *
                    fontSize * 1.43 +
                4;
            var safeFactor =
                atBottomPrintEdge
                    ? 0.90
                    : 0.96;
            if (
                requiredHeight * safeFactor >
                height
            ) {
                addPdfIssue(
                    issues,
                    String(sheet.Name),
                    address,
                    "自动换行文字所需高度约" +
                        Math.ceil(requiredHeight) +
                        "磅，当前仅" +
                        Math.round(height) +
                        "磅，可能出现半行或半个字",
                    "高"
                );
            }
            return;
        }

        var requiredWidth =
            getTextVisualUnits(
                String(raw)
            ) *
            fontSize *
            1.10;

        if (
            alignment === -4108 ||
            alignment === -4130
        ) {
            requiredWidth *= 1.05;
        }

        if (shrink) {
            var shrinkRatio =
                availableWidth /
                Math.max(1, requiredWidth);
            if (shrinkRatio < 0.72) {
                addPdfIssue(
                    issues,
                    String(sheet.Name),
                    address,
                    "已启用缩小字体，但估算需缩至" +
                        Math.round(shrinkRatio * 100) +
                        "%才能放下，PDF中可能过小或显示异常",
                    "中"
                );
            }
            return;
        }

        var blocked = false;
        try {
            if (
                descriptor &&
                descriptor.columnCount > 1
            ) {
                blocked = true;
            }
        } catch (ignored7) {}
        try {
            if (
                getVisibleBorder(
                    cell,
                    10
                )
            ) {
                blocked = true;
            }
        } catch (ignored8) {}
        try {
            var nextCell =
                sheet.Cells.Item(
                    Number(cell.Row),
                    (
                        descriptor
                            ? descriptor.lastColumn
                            : Number(cell.Column)
                    ) + 1
                );
            if (
                !isBlankValue(
                    readMergedAwareValue(
                        nextCell
                    )
                )
            ) {
                blocked = true;
            }
        } catch (ignored9) {}

        var ratio =
            requiredWidth /
            Math.max(1, availableWidth);

        if (
            blocked &&
            ratio >= 1.0
        ) {
            addPdfIssue(
                issues,
                String(sheet.Name),
                address,
                "未换行且未缩小字体，估算文字宽度超过可用宽度，PDF可能横向截断",
                "高"
            );
        } else if (
            blocked &&
            ratio >= 0.86
        ) {
            addPdfIssue(
                issues,
                String(sheet.Name),
                address,
                "文字已接近单元格右边界，字体替换或PDF渲染差异可能造成半个字",
                "中"
            );
        } else if (
            atRightPrintEdge &&
            ratio >= 0.80
        ) {
            addPdfIssue(
                issues,
                String(sheet.Name),
                address,
                "文字位于打印区域最右列且接近边界，导出PDF可能切掉末尾半个字",
                "中"
            );
        }
    }


    function inspectPdfRisksOnSheet(
        sheet
    ) {
        var issues = [];
        var actual =
            getActualContentBounds(
                sheet
            );
        if (!actual) {
            return {
                issues: issues,
                scannedCells: 0
            };
        }

        var printArea = null;
        try {
            printArea =
                parseSinglePrintArea(
                    sheet.PageSetup
                        .PrintArea
                );
        } catch (ignored0) {}

        if (!printArea) {
            addPdfIssue(
                issues,
                String(sheet.Name),
                columnToLetters(
                    actual.firstColumn
                ) +
                actual.firstRow,
                "未设置打印区域，导出结果可能受UsedRange残留格式影响",
                "中"
            );
        } else if (
            actual.firstRow <
                printArea.firstRow ||
            actual.lastRow >
                printArea.lastRow ||
            actual.firstColumn <
                printArea.firstColumn ||
            actual.lastColumn >
                printArea.lastColumn
        ) {
            addPdfIssue(
                issues,
                String(sheet.Name),
                columnToLetters(
                    actual.firstColumn
                ) +
                actual.firstRow +
                ":" +
                columnToLetters(
                    actual.lastColumn
                ) +
                actual.lastRow,
                "实际内容超出PrintArea，导出PDF可能漏行、漏列或切掉右侧文字",
                "明确"
            );
        }

        var scanBounds =
            printArea || actual;
        scanBounds = {
            firstRow: Math.min(
                scanBounds.firstRow,
                actual.firstRow
            ),
            lastRow: Math.max(
                scanBounds.lastRow,
                actual.lastRow
            ),
            firstColumn: Math.min(
                scanBounds.firstColumn,
                actual.firstColumn
            ),
            lastColumn: Math.max(
                scanBounds.lastColumn,
                actual.lastColumn
            )
        };

        var rowCount =
            scanBounds.lastRow -
            scanBounds.firstRow + 1;
        var columnCount =
            scanBounds.lastColumn -
            scanBounds.firstColumn + 1;
        var scanCells =
            rowCount * columnCount;

        if (
            scanCells >
            MAX_PDF_SCAN_CELLS
        ) {
            addPdfIssue(
                issues,
                String(sheet.Name),
                columnToLetters(
                    scanBounds.firstColumn
                ) +
                scanBounds.firstRow +
                ":" +
                columnToLetters(
                    scanBounds.lastColumn
                ) +
                scanBounds.lastRow,
                "检查范围超过" +
                    MAX_PDF_SCAN_CELLS +
                    "个单元格，已跳过逐格文字检测",
                "提示"
            );
            return {
                issues: issues,
                scannedCells:
                    scanCells
            };
        }

        var horizontalBreaks =
            getPageBreakLocations(
                sheet,
                true
            );
        var verticalBreaks =
            getPageBreakLocations(
                sheet,
                false
            );
        var seenMerge = {};

        for (
            var row =
                scanBounds.firstRow;
            row <=
                scanBounds.lastRow;
            row += 1
        ) {
            for (
                var column =
                    scanBounds.firstColumn;
                column <=
                    scanBounds.lastColumn;
                column += 1
            ) {
                var cell =
                    sheet.Cells.Item(
                        row,
                        column
                    );

                if (
                    !isTopLeftOfMerge(
                        cell
                    )
                ) {
                    continue;
                }

                var descriptor =
                    getMergedRangeDescriptor(
                        cell
                    );

                if (descriptor) {
                    var key =
                        descriptor.firstRow +
                        ":" +
                        descriptor.firstColumn +
                        ":" +
                        descriptor.lastRow +
                        ":" +
                        descriptor.lastColumn;

                    if (!seenMerge[key]) {
                        seenMerge[key] = true;

                        if (
                            mergeCrossesBreak(
                                descriptor,
                                horizontalBreaks,
                                true
                            ) ||
                            mergeCrossesBreak(
                                descriptor,
                                verticalBreaks,
                                false
                            )
                        ) {
                            addPdfIssue(
                                issues,
                                String(sheet.Name),
                                columnToLetters(
                                    descriptor.firstColumn
                                ) +
                                descriptor.firstRow +
                                ":" +
                                columnToLetters(
                                    descriptor.lastColumn
                                ) +
                                descriptor.lastRow,
                                "分页线穿过合并单元格，PDF可能出现文字或边框被切开",
                                "高"
                            );
                        }
                    }
                }

                inspectPdfTextCell(
                    sheet,
                    cell,
                    descriptor,
                    issues,
                    printArea
                );
            }
        }

        return {
            issues: issues,
            scannedCells:
                scanCells
        };
    }

    function detectPdfRisksByScope(
        mode
    ) {
        var context =
            getTargetContext(false);
        var workbook =
            context.workbook;
        var originalSheetName =
            getOriginalActiveSheetName(
                workbook
            );
        var sheetNames =
            getScopeSheetNames(
                workbook,
                context.sheet,
                mode
            );
        var issues = [];
        var scannedCells = 0;
        var failures = [];

        for (
            var i = 0;
            i < sheetNames.length;
            i += 1
        ) {
            updateProgress(
                "PDF导出风险检测",
                i + 1,
                sheetNames.length
            );

            try {
                var sheet =
                    activateWorksheet(
                        workbook,
                        sheetNames[i]
                    );
                var result =
                    inspectPdfRisksOnSheet(
                        sheet
                    );

                scannedCells +=
                    Number(
                        result.scannedCells ||
                        0
                    );

                for (
                    var j = 0;
                    j <
                        result.issues.length &&
                    issues.length <
                        MAX_PDF_ISSUES;
                    j += 1
                ) {
                    issues.push(
                        result.issues[j]
                    );
                }
            } catch (error) {
                failures.push(
                    "“" + sheetNames[i] +
                    "”：" +
                    (
                        error &&
                        error.message
                            ? error.message
                            : String(error)
                    )
                );
            }
        }

        restoreActiveSheet(
            workbook,
            originalSheetName
        );

        if (!issues.length) {
            MsgBox(
                "已检查" +
                sheetNames.length +
                "张工作表、约" +
                scannedCells +
                "个单元格。" +
                "\n没有发现明显的PDF文字截断、打印范围遗漏或分页穿越合并单元格风险。" +
                (
                    failures.length
                        ? "\n另有" +
                          failures.length +
                          "张工作表检查失败。"
                        : ""
                ) +
                "\n\n该功能属于几何估算，导出前仍建议查看打印预览。",
                failures.length
                    ? JS_EXCLAMATION
                    : JS_INFORMATION,
                "PDF风险检测完成"
            );
            return;
        }

        var displayCount =
            Math.min(
                issues.length,
                24
            );
        var lines = [];

        for (
            var k = 0;
            k < displayCount;
            k += 1
        ) {
            lines.push(
                (k + 1) +
                ". [" +
                issues[k].severity +
                "] “" +
                issues[k].sheetName +
                "”!" +
                issues[k].address +
                "\n   " +
                issues[k].reason
            );
        }

        if (
            issues.length >
            displayCount
        ) {
            lines.push(
                "……另有" +
                (
                    issues.length -
                    displayCount
                ) +
                "项未展开。"
            );
        }

        MsgBox(
            "已检查" +
            sheetNames.length +
            "张工作表、约" +
            scannedCells +
            "个单元格。" +
            "\n发现" +
            issues.length +
            "项PDF导出风险：\n\n" +
            lines.join("\n") +
            (
                failures.length
                    ? "\n\n另有" +
                      failures.length +
                      "张工作表检查失败。"
                    : ""
            ) +
            "\n\n关闭窗口后将定位第一项风险。",
            JS_EXCLAMATION,
            "PDF导出风险检测"
        );

        try {
            var first =
                issues[0];
            var firstSheet =
                findWorksheet(
                    workbook,
                    first.sheetName
                );
            if (firstSheet) {
                firstSheet.Activate();
                firstSheet
                    .Range(
                        first.address
                    )
                    .Select();
            }
        } catch (ignored) {}
    }


    function showHelp() {
        MsgBox(
            "工程表清理助手 v1.8.0\n\n" +
            "1. 查找并清空文字：可直接输入关键词；默认包含匹配并忽略空格。例如输入第1页，可匹配第 1 页 共 2 页。\n" +
            "2. 自动重建分页：按后续宽合并表名或连续表格块重建分页，不强制一页宽。\n" +
            "3. 允许手动拖动：取消按页数缩放，改为100%并打开分页预览。\n" +
            "4. 线宽统一：共享边界两侧同步写入，并进行最多三轮纠正和最终核验。\n" +
            "5. PDF风险检测：增加显示文本短于原值、接近右边界、换行高度不足等半字风险。\n" +
            "6. 普通删行删列只收缩PrintArea，不重置用户手动分页线。\n\n" +
            "打印布局和内容清除前请保存文件副本。",
            JS_INFORMATION,
            "使用说明"
        );
    }


    if (action === "deleteBlankRows") return deleteBlankRowsByScope(scopeMode || "sheet");
    if (action === "deleteSelectedBlankRows") return deleteSelectedBlankCellRows();
    if (action === "deleteSelectedColumns") return deleteSelectedColumnsByScope(scopeMode || "sheet");
    if (action === "promptDeleteColumnsByLetters") return promptDeleteColumnsByLetters();
    if (action === "adjustPrintWide") return adjustPrintAreaByScope(scopeMode || "sheet");
    if (action === "buildBlankColumnPreview") return buildBlankColumnPreviewByScope(scopeMode || "sheet");
    if (action === "executePreviewDelete") return executePreviewDelete();
    if (action === "locatePreviewRow") return locatePreviewRow();
    if (action === "renumberMajorItems") return renumberMajorItemsByScope(scopeMode || "sheet");
    if (action === "renumberNumericItems") return renumberNumericItemsByScope(scopeMode || "sheet");
    if (action === "deleteRowsBySelectedText") return deleteRowsBySelectedTextByScope(scopeMode || "sheet");
    if (action === "buildDuplicateSelection") return buildDuplicateSelectionByScope(scopeMode || "sheet");
    if (action === "executeDuplicateDelete") return executeDuplicateSelectionDelete();
    if (action === "locateDuplicateRow") return locateDuplicateSelectionRow();
    if (action === "alignTableBottom") return alignTableBottomByScope(scopeMode || "sheet");
    if (action === "enableManualPageBreakEditing") return enableManualPageBreakEditingByScope(scopeMode || "sheet");
    if (action === "cleanupOrphanHeaders") return cleanupOrphanHeadersByScope(scopeMode || "sheet");
    if (action === "batchDeleteWorksheets") return promptBatchDeleteWorksheets();
    if (action === "undoLastOperation") return undoLastOperation();
    if (action === "clearUndoPoint") return clearUndoPointCommand();
    if (action === "cleanupLegacySheets") return cleanupLegacyVisibleSheets();
    if (action === "allPageBreakPreview") return setAllWorksheetViews(2, "分页预览");
    if (action === "allNormalView") return setAllWorksheetViews(1, "普通视图");
    if (action === "normalizeTableBorders") return normalizeTableBordersByScope(scopeMode || "sheet");
    if (action === "detectPdfRisks") return detectPdfRisksByScope(scopeMode || "sheet");
    if (action === "help") return showHelp();
    if (action === "promptDeleteBlankRows") {
        var scope1 = promptScope();
        if (scope1) return deleteBlankRowsByScope(scope1);
        return;
    }
    if (action === "promptBuildPreview") {
        var scope2 = promptScope();
        if (scope2) return buildBlankColumnPreviewByScope(scope2);
        return;
    }
    if (action === "promptRenumber") {
        var scope3 = promptScope();
        if (scope3) return renumberMajorItemsByScope(scope3);
        return;
    }
    if (action === "promptRenumberNumeric") {
        var scope4 = promptScope();
        if (scope4) return renumberNumericItemsByScope(scope4);
        return;
    }
    if (action === "promptDeleteRowsByText") {
        var scope5 = promptScope();
        if (scope5) return deleteRowsBySelectedTextByScope(scope5);
        return;
    }
    throw new Error("未知功能：" + action);
}

function 工程表清理_安全执行(action, scopeMode) {
    var fastState = 工程表清理_快速模式开始("工程表清理助手正在处理…");
    try {
        return 工程表清理_内部执行(action, scopeMode);
    } catch (error) {
        var message = error && error.message ? error.message : String(error);
        MsgBox("操作失败：\n" + message, 48, "工程表清理助手");
    } finally {
        工程表清理_快速模式结束(fastState);
    }
}

/* 宏编辑器中可直接运行的公开宏 */
function 工程表_删除完全空白行() { 工程表清理_安全执行("promptDeleteBlankRows"); }
function 工程表_删除所选空白格所在行() { 工程表清理_安全执行("deleteSelectedBlankRows"); }
function 工程表_删除当前表所选列() { 工程表清理_安全执行("deleteSelectedColumns", "sheet"); }
function 工程表_整个工作簿删除同列() { 工程表清理_安全执行("deleteSelectedColumns", "workbook"); }
function 工程表_输入列号删除() { 工程表清理_安全执行("promptDeleteColumnsByLetters"); }
function 工程表_调整当前表打印范围() { 工程表清理_安全执行("adjustPrintWide", "sheet"); }
function 工程表_调整整个工作簿打印范围() { 工程表清理_安全执行("adjustPrintWide", "workbook"); }
function 工程表_按指定空白列生成预览() { 工程表清理_安全执行("promptBuildPreview"); }
function 工程表_按指定列为空删除行() { 工程表清理_安全执行("promptBuildPreview"); }
function 工程表_执行预览删除() { 工程表清理_安全执行("executePreviewDelete"); }
function 工程表_定位预览原行() { 工程表清理_安全执行("locatePreviewRow"); }
function 工程表_修复中文大项编号() { 工程表清理_安全执行("promptRenumber"); }
function 工程表_修复数字序号() { 工程表清理_安全执行("promptRenumberNumeric"); }
function 工程表_按选中文字删除行() { 工程表清理_安全执行("promptDeleteRowsByText"); }
function 工程表_查找重复内容() { 工程表清理_安全执行("buildDuplicateSelection", "sheet"); }
function 工程表_执行重复内容删除() { 工程表清理_安全执行("executeDuplicateDelete"); }
function 工程表_定位重复内容原行() { 工程表清理_安全执行("locateDuplicateRow"); }
function 工程表_对齐当前表格底边() { 工程表清理_安全执行("alignTableBottom", "sheet"); }
function 工程表_对齐整个工作簿底边() { 工程表清理_安全执行("alignTableBottom", "workbook"); }
function 工程表_启用当前表手动分页() { 工程表清理_安全执行("enableManualPageBreakEditing", "sheet"); }
function 工程表_启用整个工作簿手动分页() { 工程表清理_安全执行("enableManualPageBreakEditing", "workbook"); }
function 工程表_清理当前表尾页表头() { 工程表清理_安全执行("cleanupOrphanHeaders", "sheet"); }
function 工程表_清理整个工作簿尾页表头() { 工程表清理_安全执行("cleanupOrphanHeaders", "workbook"); }
function 工程表_批量删除工作表() { 工程表清理_安全执行("batchDeleteWorksheets"); }
function 工程表_撤回上一步() { 工程表清理_安全执行("undoLastOperation"); }
function 工程表_清除撤回点() { 工程表清理_安全执行("clearUndoPoint"); }
function 工程表_全部分页预览() { 工程表清理_安全执行("allPageBreakPreview"); }
function 工程表_全部普通视图() { 工程表清理_安全执行("allNormalView"); }
function 工程表_统一当前表格线宽() { 工程表清理_安全执行("normalizeTableBorders", "sheet"); }
function 工程表_统一整个工作簿线宽() { 工程表清理_安全执行("normalizeTableBorders", "workbook"); }
function 工程表_PDF风险检测当前表() { 工程表清理_安全执行("detectPdfRisks", "sheet"); }
function 工程表_PDF风险检测整个工作簿() { 工程表清理_安全执行("detectPdfRisks", "workbook"); }
function 工程表_使用说明() { 工程表清理_安全执行("help"); }

/* 功能区回调 */
function RibbonDeleteBlankCurrent(control) { 工程表清理_安全执行("deleteBlankRows", "sheet"); }
function RibbonDeleteBlankWorkbook(control) { 工程表清理_安全执行("deleteBlankRows", "workbook"); }
function RibbonDeleteSelectedBlankRows(control) { 工程表清理_安全执行("deleteSelectedBlankRows"); }
function RibbonDeleteColumnsCurrent(control) { 工程表清理_安全执行("deleteSelectedColumns", "sheet"); }
function RibbonDeleteColumnsWorkbook(control) { 工程表清理_安全执行("deleteSelectedColumns", "workbook"); }
function RibbonDeleteColumnsInput(control) { 工程表清理_安全执行("promptDeleteColumnsByLetters"); }
function RibbonBuildPreviewCurrent(control) { 工程表清理_安全执行("buildBlankColumnPreview", "sheet"); }
function RibbonBuildPreviewWorkbook(control) { 工程表清理_安全执行("buildBlankColumnPreview", "workbook"); }
function RibbonExecutePreviewDelete(control) { 工程表清理_安全执行("executePreviewDelete"); }
function RibbonLocateSource(control) { 工程表清理_安全执行("locatePreviewRow"); }
function RibbonDeleteTextCurrent(control) { 工程表清理_安全执行("deleteRowsBySelectedText", "sheet"); }
function RibbonDeleteTextWorkbook(control) { 工程表清理_安全执行("deleteRowsBySelectedText", "workbook"); }
function RibbonFindDuplicatesCurrent(control) { 工程表清理_安全执行("buildDuplicateSelection", "sheet"); }
function RibbonFindDuplicatesWorkbook(control) { 工程表清理_安全执行("buildDuplicateSelection", "workbook"); }
function RibbonExecuteDuplicateDelete(control) { 工程表清理_安全执行("executeDuplicateDelete"); }
function RibbonLocateDuplicateRow(control) { 工程表清理_安全执行("locateDuplicateRow"); }
function RibbonRenumberChineseCurrent(control) { 工程表清理_安全执行("renumberMajorItems", "sheet"); }
function RibbonRenumberChineseWorkbook(control) { 工程表清理_安全执行("renumberMajorItems", "workbook"); }
function RibbonRenumberNumericCurrent(control) { 工程表清理_安全执行("renumberNumericItems", "sheet"); }
function RibbonRenumberNumericWorkbook(control) { 工程表清理_安全执行("renumberNumericItems", "workbook"); }
function RibbonPrintWideCurrent(control) { 工程表清理_安全执行("adjustPrintWide", "sheet"); }
function RibbonPrintWideWorkbook(control) { 工程表清理_安全执行("adjustPrintWide", "workbook"); }
function RibbonAlignBottomCurrent(control) { 工程表清理_安全执行("alignTableBottom", "sheet"); }
function RibbonAlignBottomWorkbook(control) { 工程表清理_安全执行("alignTableBottom", "workbook"); }
function RibbonManualPageBreakCurrent(control) { 工程表清理_安全执行("enableManualPageBreakEditing", "sheet"); }
function RibbonManualPageBreakWorkbook(control) { 工程表清理_安全执行("enableManualPageBreakEditing", "workbook"); }
function RibbonCleanupOrphanCurrent(control) { 工程表清理_安全执行("cleanupOrphanHeaders", "sheet"); }
function RibbonCleanupOrphanWorkbook(control) { 工程表清理_安全执行("cleanupOrphanHeaders", "workbook"); }
function RibbonBatchDeleteWorksheets(control) { 工程表清理_安全执行("batchDeleteWorksheets"); }
function RibbonUndoLastOperation(control) { 工程表清理_安全执行("undoLastOperation"); }
function RibbonClearUndoPoint(control) { 工程表清理_安全执行("clearUndoPoint"); }
function RibbonAllPageBreakPreview(control) { 工程表清理_安全执行("allPageBreakPreview"); }
function RibbonAllNormalView(control) { 工程表清理_安全执行("allNormalView"); }
function RibbonNormalizeBordersCurrent(control) { 工程表清理_安全执行("normalizeTableBorders", "sheet"); }
function RibbonNormalizeBordersWorkbook(control) { 工程表清理_安全执行("normalizeTableBorders", "workbook"); }
function RibbonDetectPdfCurrent(control) { 工程表清理_安全执行("detectPdfRisks", "sheet"); }
function RibbonDetectPdfWorkbook(control) { 工程表清理_安全执行("detectPdfRisks", "workbook"); }
function RibbonShowHelp(control) { 工程表清理_安全执行("help"); }
