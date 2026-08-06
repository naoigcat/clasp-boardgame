/**
 * One-based row and column positions used by Apps Script range APIs.
 *
 * Centralizing layout constants keeps services free of magic numbers and makes
 * a spreadsheet redesign a configuration change rather than a service rewrite.
 */
const SHEET_LAYOUT = {
  /** The header occupies the first row in every managed sheet. */
  FIRST_DATA_ROW: 2,
  /** Column A holds the BoardGameGeek rich-text link for each Games row. */
  GAMES_LINK_COLUMN: 1,
  /**
   * First column written after a Games metadata refresh. Column A stays
   * untouched so the rich-text BoardGameGeek link is preserved.
   */
  GAMES_WRITE_START_COLUMN: 2,
  /**
   * Number of Games value columns from B through AA. Read and write use this
   * width so column A is never part of the value range.
   */
  GAMES_VALUE_COLUMN_COUNT: 26,
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
 *
 * Derived columns are owned by spreadsheet array formulas. The updater clears
 * them after a successful refresh so those formulas recalculate from the new
 * BoardGameGeek metadata instead of retaining stale manual values.
 *
 * Recommendation columns cover players 2–10 only. Rank and later metadata
 * columns follow immediately so they stay aligned with the real sheet layout.
 */
const GAME_VALUE_COLUMN = {
  /** Column C: array-formula input that resolves the display title for the row. */
  DERIVED_TITLE: 1,
  /**
   * Column F: array-formula input that resolves the preferred player-count
   * summary.
   */
  DERIVED_PLAYER_COUNT: 4,
  /**
   * Column I: first recommendation column; offset 0 means two players and
   * offset 8 (column Q) means ten players.
   */
  PLAYER_RECOMMENDATION_START: 7,
  /** Column R: BoardGameGeek's overall board-game rank. */
  BOARD_GAME_RANK: 16,
  /** Column S: BoardGameGeek's Bayesian average rating. */
  BAYES_AVERAGE: 17,
  /** Column T: BoardGameGeek's average complexity weight. */
  AVERAGE_WEIGHT: 18,
  /** Column U: formatted minimum and maximum play time. */
  PLAY_TIME: 19,
  /** Column V: BoardGameGeek's publication year. */
  PUBLICATION_YEAR: 20,
  /** Column W: array-formula input that joins the row to Rankings data. */
  DERIVED_RANKING: 21,
  /** Column X: array-formula input that joins the row to Ratings data. */
  DERIVED_RATING: 22,
  /** Column Y: array-formula input that joins the row to Titles matching. */
  DERIVED_TITLE_MATCH: 23,
  /**
   * Column Z: timestamp of the last BoardGameGeek update attempt. Advanced on
   * failure as well as success so permanent errors rotate out of the
   * oldest-first queue; failed rows become eligible again after the short
   * failure backoff rather than the full success refresh interval.
   */
  LAST_UPDATED_AT: 24,
  /**
   * Column AA: latest fetch or parsing error shown beside the game for
   * troubleshooting.
   */
  ERROR_MESSAGE: 25,
} as const;

/**
 * Games columns reset to leave room for dependent array formulas.
 *
 * Cleared with null (not '') on every Games flush so ARRAYFORMULA can
 * re-expand; writing '' or stale derived values after a failed fetch would
 * block formulas for the full failure backoff.
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
 *
 * Existing rows are preserved across updates; Rankings only contributes new
 * URLs. The canonical title column is the join key used by Games formulas.
 */
const TITLE_COLUMN = {
  /** Board Game Arena game URL shared with the Rankings sheet. */
  URL: 0,
  /** Original Japanese title scraped from the game panel page. */
  SOURCE_TITLE: 1,
  /** Canonical title used for spreadsheet matching after normalization. */
  NORMALIZED_TITLE: 2,
  /** Latest title lookup or normalization error left blank after a success. */
  ERROR_MESSAGE: 3,
} as const;
