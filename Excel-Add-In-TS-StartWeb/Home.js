var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments)).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t;
    return { next: verb(0), "throw": verb(1), "return": verb(2) };
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
(function () {
    "use strict";
    var cellToHighlight;
    var messageBanner;
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function () {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', 1.1)) {
                $("#template-description").text("This sample will display the value of the cells you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");
                $('#highlight-button').click(displaySelectedCells);
                return;
            }
            $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the largest number.");
            loadSampleData();
            // Add a click event handler for the highlight button.
            $('#highlight-button').click(hightlightHighestValue);
        });
    };
    function loadSampleData() {
        var _this = this;
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) { return __awaiter(_this, void 0, void 0, function () {
            var sheet;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        sheet = ctx.workbook.worksheets.getActiveWorksheet();
                        // Queue a command to write the sample data to the worksheet
                        sheet.getRange("B3:D5").values = values;
                        return [4 /*yield*/, ctx.sync()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        }); }).catch(errorHandler);
    }
    function hightlightHighestValue() {
        var _this = this;
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) { return __awaiter(_this, void 0, void 0, function () {
            var sourceRange, highestRow, highestCol, highestValue, i, j, usedRange;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");
                        // Run the queued-up command, and return a promise to indicate task completion
                        return [4 /*yield*/, ctx.sync()];
                    case 1:
                        // Run the queued-up command, and return a promise to indicate task completion
                        _a.sent();
                        highestRow = 0;
                        highestCol = 0;
                        highestValue = sourceRange.values[0][0];
                        // Find the cell to highlight
                        for (i = 0; i < sourceRange.rowCount; i++) {
                            for (j = 0; j < sourceRange.columnCount; j++) {
                                if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                    highestRow = i;
                                    highestCol = j;
                                    highestValue = sourceRange.values[i][j];
                                }
                            }
                        }
                        cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                        usedRange = sourceRange.worksheet.getUsedRange();
                        usedRange.format.fill.clear();
                        usedRange.format.font.bold = false;
                        // Highlight the cell
                        cellToHighlight.format.fill.color = "orange";
                        cellToHighlight.format.font.bold = true;
                        return [4 /*yield*/, ctx.sync()];
                    case 2:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        }); }).catch(errorHandler);
    }
    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                showNotification("The selected text is:", "" + result.value);
            }
            else {
                showNotification('Error', result.error.message);
            }
        });
    }
    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
    // Helper function for displaying notifications
    function showNotification(header, text) {
        var container = document.getElementById('notification-popup');
        var headerPlaceholder = container.querySelector('.notification-popup-title');
        var textPlaceholder = container.querySelector('.ms-MessageBanner-clipper');
        headerPlaceholder.textContent = header;
        textPlaceholder.textContent = text;
        var closeButton = container.querySelector('.ms-MessageBanner-close');
        closeButton.addEventListener("click", function () {
            if (container.className.indexOf("hide") === -1) {
                container.className = "ms-MessageBanner is-hidden";
            }
            closeButton.removeEventListener("click", null);
        });
        container.className = "ms-MessageBanner is-expanded";
    }
})();
//# sourceMappingURL=Home.js.map