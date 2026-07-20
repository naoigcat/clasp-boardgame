/**
 * One-based row and column positions used by Apps Script range APIs.
 */
const SHEET_LAYOUT = {
  /** The header occupies the first row in every managed sheet. */
  FIRST_DATA_ROW: 2,
  /** First column of the Games range written after a metadata refresh. */
  GAMES_WRITE_START_COLUMN: 2,
  /** First column of the Titles range. */
  TITLES_START_COLUMN: 1,
  /** First column of the Rankings and Ratings ranges. */
  DEFAULT_START_COLUMN: 1,
  /** Number of columns written into the Titles sheet. */
  TITLE_COLUMN_COUNT: 4,
  /** Number of columns written into the Ratings sheet. */
  RATING_COLUMN_COUNT: 2,
  /** Number of columns written into the Rankings sheet. */
  RANKING_COLUMN_COUNT: 15,
} as const;

/**
 * Zero-based positions in a Games row after column A has been separated as a
 * rich-text link and columns B through AA have been loaded as cell values.
 */
const GAME_VALUE_COLUMN = {
  /** Formula input cleared before refreshing a game. */
  DERIVED_TITLE: 1,
  /** Formula input cleared before refreshing a game. */
  DERIVED_PLAYER_COUNT: 4,
  /** First column for recommended player counts, representing two players. */
  PLAYER_RECOMMENDATION_START: 7,
  /** BoardGameGeek's overall board-game rank. */
  BOARD_GAME_RANK: 16,
  /** BoardGameGeek's Bayesian average rating. */
  BAYES_AVERAGE: 17,
  /** BoardGameGeek's average complexity weight. */
  AVERAGE_WEIGHT: 18,
  /** Formatted minimum and maximum play time. */
  PLAY_TIME: 19,
  /** BoardGameGeek's publication year. */
  PUBLICATION_YEAR: 20,
  /** Formula input cleared before refreshing a game. */
  DERIVED_RANKING: 21,
  /** Formula input cleared before refreshing a game. */
  DERIVED_RATING: 22,
  /** Formula input cleared before refreshing a game. */
  DERIVED_TITLE_MATCH: 23,
  /** Timestamp of the last successful BoardGameGeek update. */
  LAST_UPDATED_AT: 24,
  /** Latest fetch or parsing error for the row. */
  ERROR_MESSAGE: 25,
} as const;

/**
 * Games columns reset to leave room for dependent array formulas.
 *
 * The formulas own these values, so a refresh must not preserve stale manual
 * values when their source data changes.
 */
const GAME_ARRAY_FORMULA_INPUT_COLUMNS = [
  GAME_VALUE_COLUMN.DERIVED_TITLE,
  GAME_VALUE_COLUMN.DERIVED_PLAYER_COUNT,
  GAME_VALUE_COLUMN.DERIVED_RANKING,
  GAME_VALUE_COLUMN.DERIVED_RATING,
  GAME_VALUE_COLUMN.DERIVED_TITLE_MATCH,
] as const;

/**
 * Zero-based positions in a Titles sheet row.
 */
const TITLE_COLUMN = {
  /** Board Game Arena game URL. */
  URL: 0,
  /** Original Japanese title scraped from the source page. */
  SOURCE_TITLE: 1,
  /** Canonical title used for spreadsheet matching. */
  NORMALIZED_TITLE: 2,
  /** Latest title lookup or normalization error. */
  ERROR_MESSAGE: 3,
} as const;
