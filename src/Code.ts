function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Functions')
    .addItem('Update', 'update')
    .addToUi();
}
