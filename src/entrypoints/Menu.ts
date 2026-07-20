/**
 * Adds the custom update command when the spreadsheet UI opens.
 *
 * Apps Script discovers this global function by name, so it intentionally
 * remains a small entry point instead of living inside a service class.
 */
function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu(MENU_CONFIG.NAME)
    .addItem(MENU_CONFIG.UPDATE_ITEM_LABEL, UPDATE_QUEUE_CONFIG.HANDLER_NAME)
    .addToUi();
}
