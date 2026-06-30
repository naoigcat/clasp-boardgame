function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Functions')
    .addItem('Update Games', 'updateGames')
    .addItem('Update Rankings', 'updateRankings')
    .addItem('Update Titles', 'updateTitles')
    .addItem('Update Ratings', 'updateRatings')
    .addToUi();
}
