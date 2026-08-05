/**
 * A Games row with column A kept as rich text and columns B through AA kept as
 * mutable cell values.
 *
 * Splitting the link from the value columns is required because Apps Script's
 * `getValues` / `setValues` APIs do not round-trip hyperlink formatting.
 */
interface GameSheetRow {
  /** Rich-text link in column A, or null when the sheet row is empty. */
  readonly gameLink: GoogleAppsScript.Spreadsheet.RichTextValue | null;
  /** Values from columns B through AA, in their original sheet order. */
  readonly values: SpreadsheetCellRow;
}

/**
 * Identifies a BoardGameGeek resource extracted from a game URL.
 */
interface BoardGameGeekGameReference {
  /** BoardGameGeek resource type, such as `boardgame`. */
  readonly type: string;
  /** BoardGameGeek's numeric game identifier. */
  readonly id: string;
}

/**
 * Maps a player count to BoardGameGeek's preferred recommendation label.
 */
type PlayerRecommendations = Record<string, string>;

/**
 * Tracks work consumed by the current Apps Script invocation.
 *
 * The soft runtime budget is checked before each eligible row rather than by
 * slicing the sorted list up front so skipped fresh rows do not consume time
 * budget slots, and partial progress can still flush before a hard timeout.
 */
interface GameBatchProgress {
  /** Epoch milliseconds when this invocation began processing rows. */
  readonly startedAtMilliseconds: number;
  /** Number of source rows attempted in this invocation. */
  processedCount: number;
}

/**
 * Updates BoardGameGeek metadata in bounded batches.
 *
 * Each invocation refreshes only stale rows, oldest first, then reports whether
 * more work remains so the coordinator can keep or advance the queue phase.
 */
class GameUpdater {
  /**
   * Updates one batch of stale game rows and reports whether work remains.
   *
   * Returns false when the Games sheet is missing or every linked row is still
   * within the refresh window, signaling the coordinator to move to Titles.
   */
  static run(): boolean {
    const sheet = findSheet(SHEET_NAMES.GAMES);
    if (sheet === null) {
      return false;
    }

    const rows = GameUpdater.loadRows(sheet);
    // A single reference time prevents a long batch from making equivalent rows
    // appear fresh or stale solely because they were evaluated later.
    const current = new Date();
    if (GameUpdater.countPendingRows(rows, current) === 0) {
      // Still flush surplus B–AA cleanup when every managed link is fresh so a
      // shortened column-A list cannot leave orphaned metadata until the next
      // stale refresh, including the empty-list case after all links are removed.
      GameUpdater.writeRows(sheet, rows);
      return false;
    }

    const progress: GameBatchProgress = {
      startedAtMilliseconds: Date.now(),
      processedCount: 0,
    };
    try {
      GameUpdater.sortByOldestUpdate(rows).forEach((row) => {
        GameUpdater.updateRow(row, current, progress);
      });

      GameUpdater.writeRows(sheet, rows);
      return GameUpdater.countPendingRows(rows, current) > 0;
    } catch (error: unknown) {
      Logger.log(
        `Failed after processing ${progress.processedCount} games: ${getErrorMessage(error)}`,
      );
      throw error;
    }
  }

  /**
   * Loads contiguous game rows, stopping at the first empty rich-text link.
   *
   * The sheet intentionally uses a blank link as its end marker, so trailing
   * empty rows must never be written back as data. Ranges are bounded by
   * getLastRow so open-ended A1 spans cannot pull the sheet's maximum row
   * count into memory on large spreadsheets.
   */
  private static loadRows(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
  ): GameSheetRow[] {
    const dataRowCount = sheet.getLastRow() - SHEET_LAYOUT.FIRST_DATA_ROW + 1;
    if (dataRowCount <= 0) {
      return [];
    }

    const linkValues = sheet
      .getRange(
        SHEET_LAYOUT.FIRST_DATA_ROW,
        SHEET_LAYOUT.GAMES_LINK_COLUMN,
        dataRowCount,
        1,
      )
      .getRichTextValues();
    const valueRows = sheet
      .getRange(
        SHEET_LAYOUT.FIRST_DATA_ROW,
        SHEET_LAYOUT.GAMES_WRITE_START_COLUMN,
        dataRowCount,
        SHEET_LAYOUT.GAMES_VALUE_COLUMN_COUNT,
      )
      .getValues();
    const rows = linkValues.map((linkRow, index) => ({
      gameLink: linkRow[0],
      values: valueRows[index] as SpreadsheetCellRow,
    }));
    const firstEmptyRowIndex = rows.findIndex(
      (row) => row.gameLink === null || row.gameLink.getText().length === 0,
    );

    return firstEmptyRowIndex === -1 ? rows : rows.slice(0, firstEmptyRowIndex);
  }

  /**
   * Writes columns B through AA for every managed game row in one operation.
   *
   * Column A is omitted on purpose so BoardGameGeek hyperlinks are not replaced
   * by plain-text URLs during the batch flush. After a successful write, surplus
   * B–AA cells below the last managed link are cleared so a shortened list cannot
   * leave stale metadata that formulas or manual review would misread. An empty
   * managed list skips setValues and clears every B–AA data row the same way
   * Titles and Ratings clear on an empty write, without touching column A.
   */
  private static writeRows(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    rows: readonly GameSheetRow[],
  ): void {
    if (rows.length === 0) {
      clearSurplusSheetDataRows(
        sheet,
        0,
        SHEET_LAYOUT.GAMES_VALUE_COLUMN_COUNT,
        SHEET_LAYOUT.GAMES_WRITE_START_COLUMN,
      );
      return;
    }

    sheet
      .getRange(
        SHEET_LAYOUT.FIRST_DATA_ROW,
        SHEET_LAYOUT.GAMES_WRITE_START_COLUMN,
        rows.length,
        rows[0].values.length,
      )
      .setValues(rows.map((row) => GameUpdater.valuesForWrite(row)));
    clearSurplusSheetDataRows(
      sheet,
      rows.length,
      SHEET_LAYOUT.GAMES_VALUE_COLUMN_COUNT,
      SHEET_LAYOUT.GAMES_WRITE_START_COLUMN,
    );
  }

  /**
   * Nulls sheet-owned array-formula columns on every flush.
   *
   * Failed refreshes keep prior BoardGameGeek metadata for troubleshooting, but
   * formula input cells must always be `null`: writing `''` or a stale derived
   * value blocks ARRAYFORMULA until the next successful clear, which can be as
   * long as the seven-day refresh window after a failed attempt.
   */
  private static valuesForWrite(row: GameSheetRow): SpreadsheetCellRow {
    const values = row.values.slice();
    GAME_ARRAY_FORMULA_INPUT_COLUMNS.forEach((column) => {
      values[column] = null;
    });
    return values;
  }

  /**
   * Counts rows that have a source URL and are missing or past their refresh
   * timestamp.
   */
  private static countPendingRows(
    rows: readonly GameSheetRow[],
    current: Date,
  ): number {
    return rows.filter((row) => GameUpdater.needsRefresh(row, current)).length;
  }

  /**
   * Determines whether a row can be fetched again under the refresh policy.
   *
   * Rows without a hyperlink are ownership markers or blanks and are never
   * queued for BoardGameGeek requests.
   */
  private static needsRefresh(row: GameSheetRow, current: Date): boolean {
    const gameUrl = row.gameLink?.getLinkUrl() ?? null;
    if (gameUrl === null) {
      return false;
    }

    const lastUpdatedAt = GameUpdater.getLastUpdatedAt(row);
    return (
      lastUpdatedAt === null ||
      !isWithinRefreshWindow(
        lastUpdatedAt,
        current,
        UPDATE_QUEUE_CONFIG.GAME_REFRESH_INTERVAL_DAYS,
      )
    );
  }

  /**
   * Returns a valid last-update timestamp, treating malformed values as stale.
   *
   * Non-Date cell contents are ignored so a corrupted timestamp cannot block
   * refresh forever; the row becomes eligible on the next batch.
   */
  private static getLastUpdatedAt(row: GameSheetRow): Date | null {
    const value = row.values[GAME_VALUE_COLUMN.LAST_UPDATED_AT];
    return value instanceof Date ? value : null;
  }

  /**
   * Sorts rows so the least recently refreshed games receive the next batch.
   *
   * Never-updated rows sort first (timestamp treated as epoch) so new links
   * added to the sheet are prioritized over recently failed retries.
   */
  private static sortByOldestUpdate(
    rows: readonly GameSheetRow[],
  ): GameSheetRow[] {
    return rows.slice().sort((first, second) => {
      const firstUpdatedAt =
        GameUpdater.getLastUpdatedAt(first)?.getTime() ?? 0;
      const secondUpdatedAt =
        GameUpdater.getLastUpdatedAt(second)?.getTime() ?? 0;
      return firstUpdatedAt - secondUpdatedAt;
    });
  }

  /**
   * Refreshes one eligible row and clears its dependent formula inputs after a
   * successful metadata update.
   */
  private static updateRow(
    row: GameSheetRow,
    current: Date,
    progress: GameBatchProgress,
  ): void {
    const gameUrl = row.gameLink?.getLinkUrl() ?? null;
    if (
      gameUrl === null ||
      !GameUpdater.needsRefresh(row, current) ||
      hasExceededRuntime(
        progress.startedAtMilliseconds,
        UPDATE_QUEUE_CONFIG.MAX_RUNTIME_MILLISECONDS,
      )
    ) {
      return;
    }

    progress.processedCount += 1;
    try {
      const gameReference = GameUpdater.parseGameReference(gameUrl);
      const gameItem = GameUpdater.fetchGameItem(gameReference);
      GameUpdater.applyGameItem(row, gameItem, gameReference.id, current);
      GameUpdater.clearArrayFormulaInputs(row);
    } catch (error: unknown) {
      // Row-level failures must not abort the batch; the error is recorded and
      // the next eligible game continues in the same invocation.
      GameUpdater.recordFailure(row, gameUrl, error, current);
    }
  }

  /**
   * Clears values owned by array formulas after a successful metadata refresh.
   *
   * Uses null rather than an empty string so Apps Script `setValues` truly
   * blanks the cell; empty strings block ARRAYFORMULA from re-expanding into
   * the cleared cells. `valuesForWrite` applies the same nulling for failed
   * and skipped rows so a refresh error cannot leave blocking blanks.
   */
  private static clearArrayFormulaInputs(row: GameSheetRow): void {
    GAME_ARRAY_FORMULA_INPUT_COLUMNS.forEach((column) => {
      row.values[column] = null;
    });
  }

  /**
   * Extracts the BoardGameGeek resource type and ID from an absolute game URL.
   *
   * Paths look like `/boardgame/12345/title`; type and ID are the first two
   * segments and are enough to call the thing API.
   */
  private static parseGameReference(
    gameUrl: string,
  ): BoardGameGeekGameReference {
    const match = /^https?:\/\/[^/]+\/([^/?#]+)\/([^/?#]+)/.exec(gameUrl);
    if (match === null) {
      throw new Error(`Unsupported BoardGameGeek URL: ${gameUrl}`);
    }

    return { type: match[1], id: match[2] };
  }

  /**
   * Fetches and parses a successful BoardGameGeek XML response.
   */
  private static fetchGameItem(
    gameReference: BoardGameGeekGameReference,
  ): XmlElement {
    const endpoint = GameUpdater.buildThingEndpoint(gameReference);
    Logger.log(endpoint);

    const response = HttpClient.getWithOptionalBearerToken(
      endpoint,
      ScriptPropertyStore.getOptionalValue(
        SCRIPT_PROPERTY_KEYS.BOARD_GAME_GEEK_TOKEN,
      ),
    );
    // Throttle failed responses too, so retries cannot turn an upstream outage
    // into a burst of requests.
    Utilities.sleep(BOARD_GAME_GEEK_CONFIG.REQUEST_DELAY_MILLISECONDS);

    if (response.getResponseCode() !== 200) {
      throw new Error(`HTTP ${response.getResponseCode()}`);
    }

    const document = XmlService.parse(response.getContentText());
    return getRequiredXmlChild(document.getRootElement(), 'item');
  }

  /**
   * Builds the BoardGameGeek thing endpoint for one resource.
   *
   * `stats=1` is required so ranks, averages, and weight are present in the XML.
   */
  private static buildThingEndpoint(
    gameReference: BoardGameGeekGameReference,
  ): string {
    return `${BOARD_GAME_GEEK_CONFIG.THING_ENDPOINT}?type=${encodeURIComponent(
      gameReference.type,
    )}&stats=1&id=${encodeURIComponent(gameReference.id)}`;
  }

  /**
   * Copies parsed BoardGameGeek values into their spreadsheet columns.
   *
   * All XML fields are resolved before any cell is mutated so a mid-parse
   * failure leaves the prior row intact for `recordFailure` to stamp.
   * Success clears the previous error cell and stamps the shared batch time so
   * the refresh window starts from this attempt.
   */
  private static applyGameItem(
    row: GameSheetRow,
    gameItem: XmlElement,
    gameId: string,
    current: Date,
  ): void {
    const recommendations = GameUpdater.getPlayerRecommendations(
      gameItem,
      gameId,
    );
    const ratings = getRequiredXmlChild(
      getRequiredXmlChild(gameItem, 'statistics'),
      'ratings',
    );
    const ranks = getRequiredXmlChild(ratings, 'ranks');
    // Prefer the overall board-game rank over family-specific rank entries.
    const boardGameRank = findXmlElementByAttribute(
      ranks.getChildren('rank'),
      'name',
      'boardgame',
    );
    const boardGameRankValue = parseDisplayNumber(
      getRequiredXmlAttributeValue(boardGameRank, 'value'),
    );
    const bayesAverage = parseDisplayNumber(
      getRequiredXmlAttributeValue(
        getRequiredXmlChild(ratings, 'bayesaverage'),
        'value',
      ),
    );
    const averageWeight = parseDisplayNumber(
      getRequiredXmlAttributeValue(
        getRequiredXmlChild(ratings, 'averageweight'),
        'value',
      ),
    );
    const playTime = GameUpdater.formatPlayTime(
      getRequiredXmlAttributeValue(
        getRequiredXmlChild(gameItem, 'minplaytime'),
        'value',
      ),
      getRequiredXmlAttributeValue(
        getRequiredXmlChild(gameItem, 'maxplaytime'),
        'value',
      ),
    );
    const publicationYear = parseDisplayNumber(
      getRequiredXmlAttributeValue(
        getRequiredXmlChild(gameItem, 'yearpublished'),
        'value',
      ),
    );

    GameUpdater.writePlayerRecommendations(row, recommendations);
    row.values[GAME_VALUE_COLUMN.BOARD_GAME_RANK] = boardGameRankValue;
    row.values[GAME_VALUE_COLUMN.BAYES_AVERAGE] = bayesAverage;
    row.values[GAME_VALUE_COLUMN.AVERAGE_WEIGHT] = averageWeight;
    row.values[GAME_VALUE_COLUMN.PLAY_TIME] = playTime;
    row.values[GAME_VALUE_COLUMN.PUBLICATION_YEAR] = publicationYear;
    row.values[GAME_VALUE_COLUMN.LAST_UPDATED_AT] = current;
    row.values[GAME_VALUE_COLUMN.ERROR_MESSAGE] = '';
  }

  /**
   * Reads preferred player counts from the BoardGameGeek poll and applies known
   * per-game corrections.
   *
   * For each player count, the result with the most votes wins. Overrides are
   * applied last so curated corrections beat the raw community poll.
   */
  private static getPlayerRecommendations(
    gameItem: XmlElement,
    gameId: string,
  ): PlayerRecommendations {
    const poll = findXmlElementByAttribute(
      gameItem.getChildren('poll'),
      'name',
      'suggested_numplayers',
    );
    const recommendations: PlayerRecommendations = {};

    poll.getChildren('results').forEach((results) => {
      const mostVotedResult = sortXmlElementsByNumericAttribute(
        results.getChildren('result'),
        'numvotes',
      )[0];
      if (mostVotedResult === undefined) {
        return;
      }

      recommendations[getRequiredXmlAttributeValue(results, 'numplayers')] =
        getRequiredXmlAttributeValue(mostVotedResult, 'value');
    });

    return {
      ...recommendations,
      ...(GAME_PLAYER_COUNT_OVERRIDES[gameId] ?? {}),
    };
  }

  /**
   * Writes recommendations for the player counts represented in the sheet.
   *
   * Missing poll entries become blanks rather than guessed labels so the sheet
   * does not invent recommendations BoardGameGeek never provided.
   */
  private static writePlayerRecommendations(
    row: GameSheetRow,
    recommendations: PlayerRecommendations,
  ): void {
    for (
      let offset = 0;
      offset < BOARD_GAME_GEEK_CONFIG.SUPPORTED_PLAYER_COUNT_TOTAL;
      offset += 1
    ) {
      const playerCount =
        BOARD_GAME_GEEK_CONFIG.MIN_SUPPORTED_PLAYER_COUNT + offset;
      row.values[GAME_VALUE_COLUMN.PLAYER_RECOMMENDATION_START + offset] =
        recommendations[playerCount.toString()] ?? '';
    }
  }

  /**
   * Formats equal play times as one value and ranges as a readable interval.
   */
  private static formatPlayTime(
    minimumValue: string,
    maximumValue: string,
  ): NumericDisplayValue | string {
    const minimum = parseDisplayNumber(minimumValue);
    const maximum = parseDisplayNumber(maximumValue);
    return minimum === maximum ? minimum : `${minimum}-${maximum}`;
  }

  /**
   * Records a row-level error while allowing the rest of the batch to proceed.
   *
   * Advancing the attempt timestamp keeps permanent failures from monopolizing
   * every oldest-first batch; the refresh window later makes the row eligible
   * again so transient outages are still retried.
   */
  private static recordFailure(
    row: GameSheetRow,
    gameUrl: string,
    error: unknown,
    current: Date,
  ): void {
    const rowIdentifier = row.gameLink?.getText() || 'unnamed game';
    const errorMessage = getErrorMessage(error);
    Logger.log(
      `Error processing ${rowIdentifier} (URL: ${gameUrl}): ${errorMessage}`,
    );
    row.values[GAME_VALUE_COLUMN.LAST_UPDATED_AT] = current;
    row.values[GAME_VALUE_COLUMN.ERROR_MESSAGE] = errorMessage;
  }
}
