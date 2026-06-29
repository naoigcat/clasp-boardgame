function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Functions')
    .addItem('Update Games', 'updateGames')
    .addItem('Update Arena Rankings', 'updateArenaRankings')
    .addItem('Update Arena Titles', 'updateArenaTitles')
    .addItem('Update Ratings', 'updateRatings')
    .addToUi();
}
