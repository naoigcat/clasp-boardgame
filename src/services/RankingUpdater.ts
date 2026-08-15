/**
 * Board Game Arena tag metadata used to label a game.
 */
interface BoardGameArenaTag {
  /** Numeric tag identifier referenced by games. */
  readonly id: number;
  /** Human-readable Japanese tag name. */
  readonly name: string;
}

/**
 * Board Game Arena game metadata required by the Rankings sheet.
 *
 * Only fields that map to spreadsheet columns are retained; the catalog embed
 * contains additional properties that this project intentionally ignores.
 */
interface BoardGameArenaGame {
  /** URL-safe Board Game Arena game name used in the panel URL. */
  readonly name: string;
  /** Numeric tag IDs attached to the game. */
  readonly tagIds: readonly number[];
  /** Number of games played on Board Game Arena. */
  readonly gamesPlayed: number;
  /** Typical game duration in minutes. */
  readonly averageDuration: number;
  /** Default player count shown by Board Game Arena. */
  readonly defaultPlayerCount: number;
  /** Supported player counts shown by Board Game Arena. */
  readonly playerCounts: readonly number[];
}

/**
 * A row written into the Rankings sheet.
 */
type RankingSheetRow = Array<string | number | boolean | null>;

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

    const rows = RankingUpdater.parseRows(response.getContentText());
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

  /**
   * Parses the embedded catalog and maps it to spreadsheet rows.
   *
   * Tag names and games arrive as separate JSON arrays; tags are resolved after
   * both are validated so a partial parse cannot write half-built rows.
   */
  private static parseRows(page: string): RankingSheetRow[] {
    try {
      const tagNames = RankingUpdater.createTagNameMap(
        extractEmbeddedJsonArray(page, 'game_tags'),
      );
      const games = extractEmbeddedJsonArray(page, 'game_list').map((value) =>
        RankingUpdater.toGame(value),
      );
      return games.map((game) => RankingUpdater.toSheetRow(game, tagNames));
    } catch (error: unknown) {
      Logger.log(
        `Unable to parse Board Game Arena catalog: ${getErrorMessage(error)}`,
      );
      throw error;
    }
  }

  /**
   * Creates a lookup table from tag ID to display name.
   */
  private static createTagNameMap(
    rawTags: readonly unknown[],
  ): Record<number, string> {
    return rawTags.reduce<Record<number, string>>((tagNames, rawTag) => {
      const tag = RankingUpdater.toTag(rawTag);
      tagNames[tag.id] = tag.name;
      return tagNames;
    }, {});
  }

  /**
   * Converts a parsed JSON value into validated tag metadata.
   *
   * Runtime validation is required because the embed is untyped page source,
   * not a versioned API schema.
   */
  private static toTag(value: unknown): BoardGameArenaTag {
    const record = RankingUpdater.requireRecord(value, 'game tag');
    return {
      id: RankingUpdater.requireNumber(record, 'id', 'game tag'),
      name: RankingUpdater.requireString(record, 'name', 'game tag'),
    };
  }

  /**
   * Converts a parsed JSON value into the game fields required by this service.
   */
  private static toGame(value: unknown): BoardGameArenaGame {
    const record = RankingUpdater.requireRecord(value, 'game');
    return {
      name: RankingUpdater.requireString(record, 'name', 'game'),
      tagIds: RankingUpdater.requireTagIds(record, 'tags', 'game'),
      gamesPlayed: RankingUpdater.requireNumber(record, 'games_played', 'game'),
      averageDuration: RankingUpdater.requireNumber(
        record,
        'average_duration',
        'game',
      ),
      defaultPlayerCount: RankingUpdater.requireNumber(
        record,
        'default_num_players',
        'game',
      ),
      playerCounts: RankingUpdater.requireNumberArray(
        record,
        'player_numbers',
        'game',
      ),
    };
  }

  /**
   * Builds a Rankings row from one validated Board Game Arena game.
   *
   * Column B is left null for spreadsheet-side formulas. Boolean player-count
   * columns follow the configured 2–10 layout rather than the source's raw list.
   */
  private static toSheetRow(
    game: BoardGameArenaGame,
    tagNames: Readonly<Record<number, string>>,
  ): RankingSheetRow {
    return [
      `${BOARD_GAME_ARENA_CONFIG.GAME_PANEL_URL}${encodeURIComponent(game.name)}`,
      null,
      RankingUpdater.resolveTagNames(game.tagIds, tagNames).join(' '),
      game.gamesPlayed,
      game.averageDuration,
      game.defaultPlayerCount,
      ...BOARD_GAME_ARENA_CONFIG.SUPPORTED_PLAYER_COUNTS.map((playerCount) =>
        game.playerCounts.includes(playerCount),
      ),
    ];
  }

  /**
   * Resolves displayable tags while excluding Board Game Arena implementation
   * tags that do not describe a game's characteristics.
   */
  private static resolveTagNames(
    tagIds: readonly number[],
    tagNames: Readonly<Record<number, string>>,
  ): string[] {
    const ignoredTagIds = new Set<number>(
      BOARD_GAME_ARENA_CONFIG.IGNORED_TAG_IDS,
    );
    return tagIds
      .filter((tagId) => !ignoredTagIds.has(tagId))
      .map((tagId) => tagNames[tagId])
      .filter((tagName): tagName is string => tagName !== undefined);
  }

  /**
   * Requires a JSON object before reading named properties from it.
   */
  private static requireRecord(
    value: unknown,
    label: string,
  ): Record<string, unknown> {
    if (typeof value !== 'object' || value === null || Array.isArray(value)) {
      throw new Error(`Expected ${label} to be an object`);
    }
    return value as Record<string, unknown>;
  }

  /**
   * Requires a string property from parsed Board Game Arena data.
   */
  private static requireString(
    record: Readonly<Record<string, unknown>>,
    propertyName: string,
    label: string,
  ): string {
    const value = record[propertyName];
    if (typeof value !== 'string') {
      throw new Error(`Expected ${label}.${propertyName} to be a string`);
    }
    return value;
  }

  /**
   * Requires a numeric property from parsed Board Game Arena data.
   */
  private static requireNumber(
    record: Readonly<Record<string, unknown>>,
    propertyName: string,
    label: string,
  ): number {
    const value = record[propertyName];
    if (typeof value !== 'number') {
      throw new Error(`Expected ${label}.${propertyName} to be a number`);
    }
    return value;
  }

  /**
   * Requires an array containing only numeric values.
   */
  private static requireNumberArray(
    record: Readonly<Record<string, unknown>>,
    propertyName: string,
    label: string,
  ): number[] {
    const value = record[propertyName];
    if (
      !Array.isArray(value) ||
      value.some((item) => typeof item !== 'number')
    ) {
      throw new Error(`Expected ${label}.${propertyName} to be a number array`);
    }
    return value;
  }

  /**
   * Requires Board Game Arena's nested tag tuples and returns their IDs.
   *
   * The embed stores tags as `[id, ...]` tuples; only the ID is needed once
   * display names have been loaded from `game_tags`.
   */
  private static requireTagIds(
    record: Readonly<Record<string, unknown>>,
    propertyName: string,
    label: string,
  ): number[] {
    const value = record[propertyName];
    if (
      !Array.isArray(value) ||
      value.some((tag) => !Array.isArray(tag) || typeof tag[0] !== 'number')
    ) {
      throw new Error(
        `Expected ${label}.${propertyName} to be a tag tuple array`,
      );
    }
    return value.map((tag) => tag[0] as number);
  }
}
