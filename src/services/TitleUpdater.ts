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
   * A missing Titles tab returns true after logging so a rename or deletion is
   * not treated as Titles completion. A missing Rankings tab logs and still
   * normalizes existing Titles rows; only new Ranking URLs are skipped.
   */
  static run(): boolean {
    const rankingsSheet = findSheet(SHEET_NAMES.RANKINGS);
    const titlesSheet = findSheet(SHEET_NAMES.TITLES);
    if (titlesSheet === null) {
      // Returning true keeps UPDATE_STEP on titles; false would finishUpdate and
      // hide a renamed or deleted contract tab as a successful Titles finish.
      Logger.log(
        `Titles sheet "${SHEET_NAMES.TITLES}" is missing; keeping the Titles phase until the tab is restored.`,
      );
      return true;
    }
    if (rankingsSheet === null) {
      // Rankings only supplies new URLs. Existing Titles rows can still normalize
      // without it, so do not finish the queue just because intake is unavailable.
      Logger.log(
        `Rankings sheet "${SHEET_NAMES.RANKINGS}" is missing; continuing Titles normalization without new Ranking URLs.`,
      );
    }

    const existingRows = TitleUpdater.loadRows(titlesSheet);
    const rows = existingRows.map((row) => row.slice());
    TitleUpdater.appendRankingUrls(rows, rankingsSheet);
    const originalRows = rows.map((row) => row.slice());
    const pendingRows = rows.filter(TitleUpdater.needsNormalization);
    if (pendingRows.length === 0) {
      // Still compact blank URL slots when every title is already normalized so
      // mid-sheet empty rows do not linger until the next pending URL appears.
      const result = TitleUpdater.writeRows(
        titlesSheet,
        existingRows,
        originalRows,
        rows,
      );
      return result.rows.some(TitleUpdater.needsNormalization);
    }

    // Rows that already failed stay eligible for retry, but unattempted rows
    // must advance first so a permanently broken head cannot stall the queue.
    const preferredRows = pendingRows.filter(
      (row) => !row[TITLE_COLUMN.ERROR_MESSAGE],
    );
    const retryingOnlyFailures = preferredRows.length === 0;
    // Use the full row position: successful retries disappear from pendingRows,
    // so an offset in that shrinking list would skip later failures.
    const retryStartRow = Number(
      ScriptPropertyStore.getOptionalValue(
        UPDATE_QUEUE_CONFIG.TITLE_RETRY_ROW_PROPERTY_KEY,
      ) ?? '0',
    );
    const candidateRows = retryingOnlyFailures
      ? rows.slice(retryStartRow)
      : preferredRows;
    let nextRetryRow = retryStartRow;
    const startedAtMilliseconds = Date.now();
    let stoppedForRuntime = false;

    for (const [index, row] of candidateRows.entries()) {
      if (!TitleUpdater.needsNormalization(row)) {
        continue;
      }
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
      nextRetryRow = retryStartRow + index + 1;
    }

    const result = TitleUpdater.writeRows(
      titlesSheet,
      existingRows,
      originalRows,
      rows,
    );
    if (result.edited) {
      // A positional retry cursor cannot skip new or moved URLs after an edit.
      ScriptPropertyStore.set(
        UPDATE_QUEUE_CONFIG.TITLE_RETRY_ROW_PROPERTY_KEY,
        '0',
      );
      return result.rows.some(TitleUpdater.needsNormalization);
    }
    if (retryingOnlyFailures) {
      // Advance only after the snapshot succeeds so a write failure cannot
      // make the next execution skip unsaved retry results.
      ScriptPropertyStore.set(
        UPDATE_QUEUE_CONFIG.TITLE_RETRY_ROW_PROPERTY_KEY,
        String(nextRetryRow),
      );
      // One completed retry pass ends the cycle. If the soft runtime budget
      // stopped the pass early, keep the trigger so remaining failures still
      // receive their one retry before permanent errors wait for a later update.
      return stoppedForRuntime;
    }

    // A later invocation retries failures once all unattempted rows are done.
    return rows.some(TitleUpdater.needsNormalization);
  }

  /** Loads non-empty URL rows using the sheet's current bounded data range. */
  private static loadRows(
    titlesSheet: GoogleAppsScript.Spreadsheet.Sheet,
  ): TitleSheetRow[] {
    const titlesDataRowCount = getDataRowCount(titlesSheet);
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
    return rows;
  }

  /** Appends unseen catalog URLs without replacing existing title cells. */
  private static appendRankingUrls(
    rows: TitleSheetRow[],
    rankingsSheet: GoogleAppsScript.Spreadsheet.Sheet | null,
  ): void {
    const knownUrls = new Set<string>(
      rows.map((row) => String(row[TITLE_COLUMN.URL])),
    );

    if (rankingsSheet === null) {
      return;
    }

    const rankingsDataRowCount = getDataRowCount(rankingsSheet);
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
  }

  /** Merges fetched cells into the latest URL rows without undoing user edits. */
  private static writeRows(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    existingRows: readonly TitleSheetRow[],
    originalRows: readonly TitleSheetRow[],
    updatedRows: readonly TitleSheetRow[],
  ): { rows: readonly TitleSheetRow[]; edited: boolean } {
    const currentRows = TitleUpdater.loadRows(sheet);
    const edited = JSON.stringify(currentRows) !== JSON.stringify(existingRows);
    if (!edited) {
      writeSheetSnapshot(sheet, updatedRows, SHEET_LAYOUT.TITLE_COLUMN_COUNT);
      return { rows: updatedRows, edited };
    }
    const originals = new Map<SpreadsheetCellValue, number | null>();
    originalRows.forEach((row, index) => {
      const url = row[TITLE_COLUMN.URL];
      originals.set(url, originals.has(url) ? null : index);
    });
    const currentCounts = new Map<SpreadsheetCellValue, number>();
    currentRows.forEach((row) => {
      const url = row[TITLE_COLUMN.URL];
      currentCounts.set(url, (currentCounts.get(url) ?? 0) + 1);
    });
    currentRows.forEach((row) => {
      const url = row[TITLE_COLUMN.URL];
      const index = originals.get(url);
      if (index == null || currentCounts.get(url) !== 1) {
        return;
      }
      const original = originalRows[index];
      // A changed source invalidates normalization fetched for the old source.
      if (
        row[TITLE_COLUMN.SOURCE_TITLE] !== original[TITLE_COLUMN.SOURCE_TITLE]
      ) {
        return;
      }
      row.forEach((value, column) => {
        if (column !== TITLE_COLUMN.URL && value === original[column]) {
          row[column] = updatedRows[index][column];
        }
      });
    });
    const existingUrls = new Set(
      existingRows.map((row) => row[TITLE_COLUMN.URL]),
    );
    updatedRows.forEach((row) => {
      const url = row[TITLE_COLUMN.URL];
      if (!existingUrls.has(url) && !currentCounts.has(url)) {
        currentRows.push(row);
      }
    });
    writeSheetSnapshot(sheet, currentRows, SHEET_LAYOUT.TITLE_COLUMN_COUNT);
    return { rows: currentRows, edited };
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
    // Titles/Rankings cells are editable, so reject anything outside the
    // configured game-panel prefix before UrlFetch can hit an arbitrary host.
    TitleUpdater.assertAllowedGamePanelUrl(gameUrl);

    const response = HttpClient.get(gameUrl);
    HttpClient.requireSuccess(response, 'Board Game Arena');

    const match = response
      .getContentText()
      .match(BOARD_GAME_ARENA_TITLE_CONFIG.GAME_NAME_PATTERN);
    return match?.[1] ?? '';
  }

  /**
   * Ensures a title scrape URL stays on the configured Board Game Arena panel.
   *
   * Rankings normally writes GAME_PANEL_URL rows, but existing Titles cells and
   * Rankings column A can be edited to any string. Prefix matching against that
   * constant covers host and path before UrlFetch so callers cannot point the
   * scrape at an arbitrary origin. Redirects issued by that already-allowed
   * host are fine; this check is about the request URL, not redirect policy.
   */
  private static assertAllowedGamePanelUrl(gameUrl: string): void {
    if (!gameUrl.startsWith(BOARD_GAME_ARENA_CONFIG.GAME_PANEL_URL)) {
      throw new Error(`Unsupported Board Game Arena URL: ${gameUrl}`);
    }
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
        // Keep this broad rule as an intentional policy for the current BGA
        // catalog; a concrete source title must justify narrowing it or adding
        // a dedicated alias for an ambiguous hyphenated subtitle.
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
