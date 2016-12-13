declare let fabric: any;

(() => {
    "use strict";

    let cellToHighlight: Excel.Range;
    let messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = () => {
        $(document).ready(() => {
            // Initialize the FabricUI notification mechanism and hide it
            let element = document.querySelector('.ms-MessageBanner');
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
    }

    function loadSampleData(): void {
        let values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Run a batch operation against the Excel object model
        Excel.run(async (ctx) => {
            // Create a proxy object for the active sheet
            let sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // Queue a command to write the sample data to the worksheet
            sheet.getRange("B3:D5").values = values;

            await ctx.sync();
        }).catch(errorHandler);            
    }
    
    function hightlightHighestValue(): void {
        // Run a batch operation against the Excel object model
        Excel.run(async (ctx) => {

            // Create a proxy object for the selected range and load its address and values properties
            let sourceRange = ctx.workbook.getSelectedRange().load(
                "values, rowCount, columnCount");

            // Run the queued-up command, and return a promise to indicate task completion
            await ctx.sync();

            let highestRow = 0;
            let highestCol = 0;
            let highestValue = sourceRange.values[0][0];

            // Find the cell to highlight
            for (let i = 0; i < sourceRange.rowCount; i++) {
                for (let j = 0; j < sourceRange.columnCount; j++) {
                    if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                        highestRow = i;
                        highestCol = j;
                        highestValue = sourceRange.values[i][j];
                    }
                }
            }

            cellToHighlight = sourceRange.getCell(highestRow, highestCol);
            let usedRange = sourceRange.worksheet.getUsedRange();
            usedRange.format.fill.clear();
            usedRange.format.font.bold = false;

            // Highlight the cell
            cellToHighlight.format.fill.color = "orange";
            cellToHighlight.format.font.bold = true;

            await ctx.sync();
        }).catch(errorHandler);            
    }

    function displaySelectedCells(): void {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                showNotification(`The selected text is:`, `${result.value}`);             
            } else {
                showNotification('Error', result.error.message);
            }
        });
    }

    // Helper function for treating errors
    function errorHandler(error): void {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, text): void {
        let container = document.getElementById('notification-popup');
        let headerPlaceholder = container.querySelector('.notification-popup-title');
        let textPlaceholder = container.querySelector('.ms-MessageBanner-clipper');

        headerPlaceholder.textContent = header;
        textPlaceholder.textContent = text;

        let closeButton = container.querySelector('.ms-MessageBanner-close');
        closeButton.addEventListener("click", function () {
            if (container.className.indexOf("hide") === -1) {
                container.className = "ms-MessageBanner is-hidden";
            }
            closeButton.removeEventListener("click", null);
        });

        container.className = "ms-MessageBanner is-expanded";
    }

})();
