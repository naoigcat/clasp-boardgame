/**
 * A mutable four-column row stored in the Titles sheet.
 */
type TitleSheetRow = SpreadsheetCellRow;

/**
 * Synchronizes Board Game Arena titles and converts them to canonical names.
 *
 * Rankings contributes new URLs; existing Titles rows are preserved so manual
 * corrections and prior scrape results survive catalog refreshes. Normalization
 * produces the join key used by Games sheet formulas.
 */
class TitleUpdater {
  /**
   * Updates one batch of incomplete title rows and reports whether work remains.
   *
   * Unattempted rows are preferred over rows that already failed so one broken
   * head of the queue cannot stall every later title. After only failures remain,
   * a single retry pass runs and the queue ends so permanent errors wait for the
   * next manual update instead of holding the shared trigger open forever.
   */
  static run(): boolean {
    const rankingsSheet = findSheet(SHEET_NAMES.RANKINGS);
    const titlesSheet = findSheet(SHEET_NAMES.TITLES);
    if (rankingsSheet === null || titlesSheet === null) {
      return false;
    }

    const rows = TitleUpdater.loadRows(titlesSheet, rankingsSheet);
    // countPendingRows is the incomplete gate so the early exit and the final
    // queue-alive decision share one definition of "still needs a title".
    if (TitleUpdater.countPendingRows(rows) === 0) {
      // Still compact blank URL slots when every title is already normalized so
      // mid-sheet empty rows do not linger until the next pending URL appears.
      TitleUpdater.writeRows(titlesSheet, rows);
      return false;
    }

    // Rows that already failed stay eligible for retry, but unattempted rows
    // must advance first so a permanently broken head cannot stall the queue.
    const pendingRows = rows.filter((row) =>
      TitleUpdater.needsNormalization(row),
    );
    const preferredRows = pendingRows.filter(
      (row) => !row[TITLE_COLUMN.ERROR_MESSAGE],
    );
    const retryingOnlyFailures = preferredRows.length === 0;
    const candidateRows = retryingOnlyFailures ? pendingRows : preferredRows;
    const startedAtMilliseconds = Date.now();
    let stoppedForRuntime = false;

    for (const row of candidateRows) {
      if (
        hasExceededRuntime(
          startedAtMilliseconds,
          UPDATE_QUEUE_CONFIG.MAX_RUNTIME_MILLISECONDS,
        )
      ) {
        stoppedForRuntime = true;
        break;
      }

      TitleUpdater.updateRow(row);
    }

    TitleUpdater.writeRows(titlesSheet, rows);
    if (retryingOnlyFailures) {
      // One completed retry pass ends the cycle. If the soft runtime budget
      // stopped the pass early, keep the trigger so remaining failures still
      // receive their one retry before permanent errors wait for a later update.
      return stoppedForRuntime;
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
    return TitleUpdater.countPendingRows(rows) > 0;
  }

  /**
   * Combines existing title rows with Ranking URLs that have not been seen yet.
   *
   * New URLs are appended with blank titles so later batches scrape them. Known
   * URLs keep their prior source title, canonical title, and error cells.
   * Both sheet reads stop at getLastRow so open-ended A1 spans cannot pull the
   * sheet's maximum row count into memory on large spreadsheets.
   */
  private static loadRows(
    titlesSheet: GoogleAppsScript.Spreadsheet.Sheet,
    rankingsSheet: GoogleAppsScript.Spreadsheet.Sheet,
  ): TitleSheetRow[] {
    const titlesDataRowCount =
      titlesSheet.getLastRow() - SHEET_LAYOUT.FIRST_DATA_ROW + 1;
    const rows =
      titlesDataRowCount <= 0
        ? []
        : titlesSheet
            .getRange(
              SHEET_LAYOUT.FIRST_DATA_ROW,
              SHEET_LAYOUT.TITLES_START_COLUMN,
              titlesDataRowCount,
              SHEET_LAYOUT.TITLE_COLUMN_COUNT,
            )
            .getValues()
            .filter((row) => Boolean(row[TITLE_COLUMN.URL]));
    const knownUrls = new Set<string>(
      rows.map((row) => String(row[TITLE_COLUMN.URL])),
    );

    const rankingsDataRowCount =
      rankingsSheet.getLastRow() - SHEET_LAYOUT.FIRST_DATA_ROW + 1;
    if (rankingsDataRowCount > 0) {
      rankingsSheet
        .getRange(
          SHEET_LAYOUT.FIRST_DATA_ROW,
          SHEET_LAYOUT.DEFAULT_START_COLUMN,
          rankingsDataRowCount,
          1,
        )
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
    }

    return rows;
  }

  /**
   * Writes compacted title rows and clears any prior surplus physical rows.
   *
   * loadRows drops blank URL slots so the in-memory list is shorter than the
   * sheet's previous used range. Writing first, then trimming only surplus rows,
   * keeps existing Titles data intact if setValues fails and still prevents
   * abandoned cells from keeping duplicate URLs after a successful rewrite.
   */
  private static writeRows(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    rows: TitleSheetRow[],
  ): void {
    if (rows.length === 0) {
      clearSheetDataRows(sheet, SHEET_LAYOUT.TITLE_COLUMN_COUNT);
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
    clearSurplusSheetDataRows(
      sheet,
      rows.length,
      SHEET_LAYOUT.TITLE_COLUMN_COUNT,
    );
  }

  /**
   * Counts rows that have a source URL but no canonical title yet.
   */
  private static countPendingRows(rows: readonly TitleSheetRow[]): number {
    return rows.filter((row) => TitleUpdater.needsNormalization(row)).length;
  }

  /**
   * Determines whether a title row still needs its canonical title.
   *
   * A blank normalized title keeps the row eligible even when an error message
   * is present, which is how failed scrapes re-enter a later retry pass.
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
   * An empty normalization result is also an error: clearing the error while
   * leaving the canonical title blank would keep the row preferred forever.
   */
  private static updateRow(row: TitleSheetRow): void {
    const gameUrl = String(row[TITLE_COLUMN.URL]);

    try {
      // Reuse a previously scraped source title so retries only re-normalize
      // when the HTTP fetch already succeeded in an earlier batch.
      if (!row[TITLE_COLUMN.SOURCE_TITLE]) {
        row[TITLE_COLUMN.SOURCE_TITLE] = TitleUpdater.fetchSourceTitle(gameUrl);
      }

      if (!row[TITLE_COLUMN.SOURCE_TITLE]) {
        row[TITLE_COLUMN.ERROR_MESSAGE] = 'game name not found';
        return;
      }

      const normalizedTitle = TitleUpdater.normalizeTitle(
        String(row[TITLE_COLUMN.SOURCE_TITLE]),
      );
      if (!normalizedTitle) {
        // Generic cleanup can strip every token from edition-only titles.
        // Record the failure instead of treating blank as a successful write.
        row[TITLE_COLUMN.ERROR_MESSAGE] = 'normalized title is empty';
        return;
      }

      row[TITLE_COLUMN.NORMALIZED_TITLE] = normalizedTitle;
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
   * punctuation and edition-label variants of the same title. Prefix aliases
   * cover whole series; exact aliases handle one-off naming mismatches.
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
   *
   * These cleanups strip marketing noise that differs between Board Game Arena
   * panels and the spreadsheet's canonical titles before alias lookup runs.
   */
  private static applyGenericNormalizations(sourceTitle: string): string {
    return (
      sourceTitle
        // Drop hyphen-wrapped edition or marketing suffixes.
        .replace(/-.*-/g, '')
        // Prefer full-width ampersands used by Japanese spreadsheet titles.
        .replace(/&amp;/g, '＆')
        .replace(/!/g, '！')
        .replace(/ - /g, ' － ')
        // Remove edition labels that would otherwise create near-duplicate keys.
        .replace(/《?新版》?/g, '')
        .replace(/第\d+版/g, '')
        .replace(/\(.*?パック\)/, '')
        // Normalize mixed-width separators to the forms used in the sheet.
        .replace(/\s*･\s*/g, '・')
        .replace(/\s*:\s*/g, '：')
        .trim()
    );
  }
}
