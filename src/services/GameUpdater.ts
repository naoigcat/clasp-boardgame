/**
 * A Games row with column A kept as rich text and columns B through AA kept as
 * mutable cell values.
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
 */
interface GameBatchProgress {
  /** Number of source rows attempted in this invocation. */
  processedCount: number;
}

/**
 * Updates BoardGameGeek metadata in bounded batches.
 */
class GameUpdater {
  /**
   * Updates one batch of stale game rows and reports whether work remains.
   */
  static run(): boolean {
    const sheet = findSheet(SHEET_NAMES.GAMES);
    if (sheet === null) {
      return false;
    }

    const rows = GameUpdater.loadRows(sheet);
    const current = new Date();
    if (GameUpdater.countPendingRows(rows, current) === 0) {
      return false;
    }

    const progress: GameBatchProgress = { processedCount: 0 };
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
   * empty rows must never be written back as data.
   */
  private static loadRows(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
  ): GameSheetRow[] {
    const linkValues = sheet.getRange('$A$2:$A').getRichTextValues();
    const primaryValues = sheet.getRange('$B$2:$Z').getValues();
    const updateValues = sheet.getRange('$AA$2:$AA').getValues();
    const rows = linkValues.map((linkRow, index) => ({
      gameLink: linkRow[0],
      values: [
        ...primaryValues[index],
        updateValues[index][0],
      ] as SpreadsheetCellRow,
    }));
    const firstEmptyRowIndex = rows.findIndex(
      (row) => row.gameLink === null || row.gameLink.getText().length === 0,
    );

    return firstEmptyRowIndex === -1
      ? rows
      : rows.slice(0, firstEmptyRowIndex);
  }

  /**
   * Writes columns B through AA for every managed game row in one operation.
   */
  private static writeRows(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    rows: readonly GameSheetRow[],
  ): void {
    if (rows.length === 0) {
      return;
    }

    sheet
      .getRange(
        SHEET_LAYOUT.FIRST_DATA_ROW,
        SHEET_LAYOUT.GAMES_WRITE_START_COLUMN,
        rows.length,
        rows[0].values.length,
      )
      .setValues(rows.map((row) => row.values));
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
   */
  private static getLastUpdatedAt(row: GameSheetRow): Date | null {
    const value = row.values[GAME_VALUE_COLUMN.LAST_UPDATED_AT];
    return value instanceof Date ? value : null;
  }

  /**
   * Sorts rows so the least recently refreshed games receive the next batch.
   */
  private static sortByOldestUpdate(
    rows: readonly GameSheetRow[],
  ): GameSheetRow[] {
    return rows.slice().sort((first, second) => {
      const firstUpdatedAt = GameUpdater.getLastUpdatedAt(first)?.getTime() ?? 0;
      const secondUpdatedAt =
        GameUpdater.getLastUpdatedAt(second)?.getTime() ?? 0;
      return firstUpdatedAt - secondUpdatedAt;
    });
  }

  /**
   * Clears dependent formula inputs and refreshes one eligible row if capacity
   * remains in the current batch.
   */
  private static updateRow(
    row: GameSheetRow,
    current: Date,
    progress: GameBatchProgress,
  ): void {
    GameUpdater.clearArrayFormulaInputs(row);

    const gameUrl = row.gameLink?.getLinkUrl() ?? null;
    if (
      gameUrl === null ||
      !GameUpdater.needsRefresh(row, current) ||
      progress.processedCount >= UPDATE_QUEUE_CONFIG.GAME_BATCH_SIZE
    ) {
      return;
    }

    progress.processedCount += 1;
    try {
      const gameReference = GameUpdater.parseGameReference(gameUrl);
      const gameItem = GameUpdater.fetchGameItem(gameReference);
      GameUpdater.applyGameItem(row, gameItem, gameReference.id, current);
    } catch (error: unknown) {
      GameUpdater.recordFailure(row, gameUrl, error);
    }
  }

  /**
   * Clears values owned by array formulas before their source metadata changes.
   *
   * Resetting these cells prevents stale formula outputs from blocking formula
   * expansion after a refreshed row changes its matching data.
   */
  private static clearArrayFormulaInputs(row: GameSheetRow): void {
    GAME_ARRAY_FORMULA_INPUT_COLUMNS.forEach((column) => {
      row.values[column] = '';
    });
  }

  /**
   * Extracts the BoardGameGeek resource type and ID from an absolute game URL.
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
    Utilities.sleep(BOARD_GAME_GEEK_CONFIG.REQUEST_DELAY_MILLISECONDS);

    if (response.getResponseCode() !== 200) {
      throw new Error(`HTTP ${response.getResponseCode()}`);
    }

    const document = XmlService.parse(response.getContentText());
    return getRequiredXmlChild(document.getRootElement(), 'item');
  }

  /**
   * Builds the BoardGameGeek thing endpoint for one resource.
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
   */
  private static applyGameItem(
    row: GameSheetRow,
    gameItem: XmlElement,
    gameId: string,
    current: Date,
  ): void {
    GameUpdater.writePlayerRecommendations(
      row,
      GameUpdater.getPlayerRecommendations(gameItem, gameId),
    );

    const ratings = getRequiredXmlChild(
      getRequiredXmlChild(gameItem, 'statistics'),
      'ratings',
    );
    const ranks = getRequiredXmlChild(ratings, 'ranks');
    const boardGameRank = findXmlElementByAttribute(
      ranks.getChildren('rank'),
      'name',
      'boardgame',
    );

    row.values[GAME_VALUE_COLUMN.BOARD_GAME_RANK] = parseDisplayNumber(
      getRequiredXmlAttributeValue(boardGameRank, 'value'),
    );
    row.values[GAME_VALUE_COLUMN.BAYES_AVERAGE] = parseDisplayNumber(
      getRequiredXmlAttributeValue(
        getRequiredXmlChild(ratings, 'bayesaverage'),
        'value',
      ),
    );
    row.values[GAME_VALUE_COLUMN.AVERAGE_WEIGHT] = parseDisplayNumber(
      getRequiredXmlAttributeValue(
        getRequiredXmlChild(ratings, 'averageweight'),
        'value',
      ),
    );
    row.values[GAME_VALUE_COLUMN.PLAY_TIME] = GameUpdater.formatPlayTime(
      getRequiredXmlAttributeValue(
        getRequiredXmlChild(gameItem, 'minplaytime'),
        'value',
      ),
      getRequiredXmlAttributeValue(
        getRequiredXmlChild(gameItem, 'maxplaytime'),
        'value',
      ),
    );
    row.values[GAME_VALUE_COLUMN.PUBLICATION_YEAR] = parseDisplayNumber(
      getRequiredXmlAttributeValue(
        getRequiredXmlChild(gameItem, 'yearpublished'),
        'value',
      ),
    );
    row.values[GAME_VALUE_COLUMN.LAST_UPDATED_AT] = current;
    row.values[GAME_VALUE_COLUMN.ERROR_MESSAGE] = '';
  }

  /**
   * Reads preferred player counts from the BoardGameGeek poll and applies known
   * per-game corrections.
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

      recommendations[
        getRequiredXmlAttributeValue(results, 'numplayers')
      ] = getRequiredXmlAttributeValue(mostVotedResult, 'value');
    });

    return {
      ...recommendations,
      ...(GAME_PLAYER_COUNT_OVERRIDES[gameId] ?? {}),
    };
  }

  /**
   * Writes recommendations for the player counts represented in the sheet.
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
   * The update timestamp intentionally remains unchanged so a later trigger can
   * retry a transient upstream failure instead of treating stale data as fresh.
   */
  private static recordFailure(
    row: GameSheetRow,
    gameUrl: string,
    error: unknown,
  ): void {
    const rowIdentifier = row.gameLink?.getText() || 'unnamed game';
    const errorMessage = getErrorMessage(error);
    Logger.log(
      `Error processing ${rowIdentifier} (URL: ${gameUrl}): ${errorMessage}`,
    );
    row.values[GAME_VALUE_COLUMN.ERROR_MESSAGE] = errorMessage;
  }
}
