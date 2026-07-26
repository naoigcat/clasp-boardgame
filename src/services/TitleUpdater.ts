/**
 * A mutable four-column row stored in the Titles sheet.
 */
type TitleSheetRow = SpreadsheetCellRow;

/**
 * Synchronizes Board Game Arena titles and converts them to canonical names.
 */
class TitleUpdater {
  /**
   * Updates one batch of incomplete title rows and reports whether work remains.
   */
  static run(): boolean {
    const rankingsSheet = findSheet(SHEET_NAMES.RANKINGS);
    const titlesSheet = findSheet(SHEET_NAMES.TITLES);
    if (rankingsSheet === null || titlesSheet === null) {
      return false;
    }

    const rows = TitleUpdater.loadRows(titlesSheet, rankingsSheet);
    const pendingRows = rows.filter((row) =>
      TitleUpdater.needsNormalization(row),
    );
    if (pendingRows.length === 0) {
      return false;
    }

    // Rows that already failed stay eligible for retry, but unattempted rows
    // must advance first so a permanently broken head cannot stall the queue.
    const preferredRows = pendingRows.filter(
      (row) => !row[TITLE_COLUMN.ERROR_MESSAGE],
    );
    const retryingOnlyFailures = preferredRows.length === 0;
    const batchRows = (
      retryingOnlyFailures ? pendingRows : preferredRows
    ).slice(0, UPDATE_QUEUE_CONFIG.TITLE_BATCH_SIZE);

    batchRows.forEach((row) => {
      TitleUpdater.updateRow(row);
    });

    TitleUpdater.writeRows(titlesSheet, rows);
    if (retryingOnlyFailures) {
      // One retry pass is enough for this cycle; remaining failures wait for a
      // later update instead of holding the shared trigger open forever.
      return false;
    }

    if (
      rows.some(
        (row) =>
          TitleUpdater.needsNormalization(row) &&
          !row[TITLE_COLUMN.ERROR_MESSAGE],
      )
    ) {
      return true;
    }

    // Keep the queue alive so the next invocation can run the failure-retry pass.
    return rows.some((row) => TitleUpdater.needsNormalization(row));
  }

  /**
   * Combines existing title rows with Ranking URLs that have not been seen yet.
   */
  private static loadRows(
    titlesSheet: GoogleAppsScript.Spreadsheet.Sheet,
    rankingsSheet: GoogleAppsScript.Spreadsheet.Sheet,
  ): TitleSheetRow[] {
    const rows = titlesSheet
      .getRange('$A$2:$D')
      .getValues()
      .filter((row) => Boolean(row[TITLE_COLUMN.URL]));
    const knownUrls = new Set<string>(
      rows.map((row) => String(row[TITLE_COLUMN.URL])),
    );

    rankingsSheet
      .getRange('$A$2:$A')
      .getValues()
      .forEach((rankingRow) => {
        const gameUrl = rankingRow[0];
        if (!gameUrl) {
          return;
        }

        const url = String(gameUrl);
        if (!knownUrls.has(url)) {
          rows.push([url, '', '', '']);
          knownUrls.add(url);
        }
      });

    return rows;
  }

  /**
   * Writes all managed title rows in their original and newly appended order.
   */
  private static writeRows(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    rows: TitleSheetRow[],
  ): void {
    if (rows.length === 0) {
      return;
    }

    sheet
      .getRange(
        SHEET_LAYOUT.FIRST_DATA_ROW,
        SHEET_LAYOUT.TITLES_START_COLUMN,
        rows.length,
        SHEET_LAYOUT.TITLE_COLUMN_COUNT,
      )
      .setValues(rows);
  }

  /**
   * Counts rows that have a source URL but no canonical title yet.
   */
  private static countPendingRows(rows: readonly TitleSheetRow[]): number {
    return rows.filter((row) => TitleUpdater.needsNormalization(row)).length;
  }

  /**
   * Determines whether a title row still needs its canonical title.
   */
  private static needsNormalization(row: TitleSheetRow): boolean {
    return (
      Boolean(row[TITLE_COLUMN.URL]) && !row[TITLE_COLUMN.NORMALIZED_TITLE]
    );
  }

  /**
   * Fetches a missing source title and stores either its canonical title or the
   * error that should be visible to spreadsheet users. Leaving the canonical
   * title blank keeps the row eligible for a later retry after source recovery.
   */
  private static updateRow(row: TitleSheetRow): void {
    const gameUrl = String(row[TITLE_COLUMN.URL]);

    try {
      if (!row[TITLE_COLUMN.SOURCE_TITLE]) {
        row[TITLE_COLUMN.SOURCE_TITLE] = TitleUpdater.fetchSourceTitle(gameUrl);
      }

      if (!row[TITLE_COLUMN.SOURCE_TITLE]) {
        row[TITLE_COLUMN.ERROR_MESSAGE] = 'game name not found';
        return;
      }

      row[TITLE_COLUMN.NORMALIZED_TITLE] = TitleUpdater.normalizeTitle(
        String(row[TITLE_COLUMN.SOURCE_TITLE]),
      );
      row[TITLE_COLUMN.ERROR_MESSAGE] = '';
    } catch (error: unknown) {
      const errorMessage = getErrorMessage(error);
      Logger.log(`Error updating title for ${gameUrl}: ${errorMessage}`);
      row[TITLE_COLUMN.ERROR_MESSAGE] = errorMessage;
    } finally {
      // Keep the same pacing after a parse or HTTP error to avoid amplifying a
      // temporary Board Game Arena failure with immediate retries.
      Utilities.sleep(BOARD_GAME_ARENA_TITLE_CONFIG.REQUEST_DELAY_MILLISECONDS);
    }
  }

  /**
   * Fetches a Board Game Arena game page and extracts its Japanese title.
   */
  private static fetchSourceTitle(gameUrl: string): string {
    const response = HttpClient.get(gameUrl);
    if (response.getResponseCode() !== 200) {
      throw new Error(
        `Board Game Arena returned HTTP ${response.getResponseCode()}`,
      );
    }

    const match = response
      .getContentText()
      .match(BOARD_GAME_ARENA_TITLE_CONFIG.GAME_NAME_PATTERN);
    return match?.[1] ?? '';
  }

  /**
   * Normalizes a source title into the spelling used for spreadsheet matching.
   *
   * Generic rules run first so aliases do not need duplicate entries for
   * punctuation and edition-label variants of the same title.
   */
  static normalizeTitle(sourceTitle: string): string {
    const normalizedTitle =
      TitleUpdater.applyGenericNormalizations(sourceTitle);
    const prefixAlias = Object.keys(TITLE_NORMALIZATION_PREFIX_ALIASES).find(
      (prefix) => normalizedTitle.startsWith(prefix),
    );
    if (prefixAlias !== undefined) {
      return `${TITLE_NORMALIZATION_PREFIX_ALIASES[prefixAlias]}${normalizedTitle.slice(
        prefixAlias.length,
      )}`;
    }

    return TITLE_NORMALIZATION_ALIASES[normalizedTitle] ?? normalizedTitle;
  }

  /**
   * Applies source-independent punctuation, edition, and whitespace rules.
   */
  private static applyGenericNormalizations(sourceTitle: string): string {
    return sourceTitle
      .replace(/-.*-/g, '')
      .replace(/&amp;/g, '＆')
      .replace(/!/g, '！')
      .replace(/ - /g, ' － ')
      .replace(/《?新版》?/g, '')
      .replace(/第\d+版/g, '')
      .replace(/\(.*?パック\)/, '')
      .replace(/\s*･\s*/g, '・')
      .replace(/\s*:\s*/g, '：')
      .trim();
  }
}
