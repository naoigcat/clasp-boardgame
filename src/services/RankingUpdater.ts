/**
 * Imports Board Game Arena's current game catalog into the Rankings sheet.
 *
 * The sheet is a replaceable snapshot: Titles reads its URLs, and Games formulas
 * can join against its columns. An empty parse leaves the previous snapshot so
 * a temporary site change cannot wipe the catalog.
 */
class RankingUpdater {
  /**
   * Fetches the catalog and replaces existing ranking rows when data is found.
   */
  static run(): void {
    const sheet = findSheet(SHEET_NAMES.RANKINGS);
    if (sheet === null) {
      // Sync import: log and skip so a renamed Rankings tab is visible in the
      // execution log instead of looking like a successful no-op.
      Logger.log(
        `Rankings sheet "${SHEET_NAMES.RANKINGS}" is missing; skipping ranking import until the tab is restored.`,
      );
      return;
    }

    const response = HttpClient.get(BOARD_GAME_ARENA_CONFIG.HOME_PAGE_URL);
    HttpClient.requireSuccess(response, 'Board Game Arena');

    const page = response.getContentText();
    let rows: RankingSheetRow[];
    try {
      rows = RankingParser.parseRows(page);
    } catch (error: unknown) {
      Logger.log(
        `Unable to parse Board Game Arena catalog: ${getErrorMessage(error)}`,
      );
      throw error;
    }
    if (rows.length === 0) {
      // An unexpectedly empty catalog is safer to treat as a source problem
      // than to replace the last known ranking snapshot with no rows.
      Logger.log(
        'Board Game Arena catalog was empty; keeping the previous Rankings snapshot.',
      );
      return;
    }

    writeSheetSnapshot(sheet, rows, SHEET_LAYOUT.RANKING_COLUMN_COUNT);
  }
}
