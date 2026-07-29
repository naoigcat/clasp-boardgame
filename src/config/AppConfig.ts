/**
 * Spreadsheet tab names used by the update services.
 *
 * Keeping these names together makes a renamed tab an explicit configuration
 * change instead of a hidden dependency inside a service. The updater treats
 * these names as a contract with the spreadsheet UI.
 */
const SHEET_NAMES = {
  /** BoardGameGeek metadata keyed by rich-text links in column A. */
  GAMES: 'Games',
  /** Current Board Game Arena catalog snapshot used as the Titles source. */
  RANKINGS: 'Rankings',
  /** Canonical titles that join Rankings URLs to spreadsheet matching formulas. */
  TITLES: 'Titles',
  /** Bodoge played-game ratings consumed by Games sheet formulas. */
  RATINGS: 'Ratings',
} as const;

/**
 * Settings for the single trigger that advances a multi-run update.
 *
 * Apps Script executions are time-capped, so Rankings/Ratings run immediately
 * while Games and Titles advance in smaller batches. One shared trigger keeps
 * quota use predictable and prevents legacy per-service triggers from racing.
 */
const UPDATE_QUEUE_CONFIG = {
  /** Global function name Apps Script binds to the time-driven trigger. */
  HANDLER_NAME: 'update',
  /** Script-property key that records which asynchronous phase should resume. */
  STEP_PROPERTY_KEY: 'UPDATE_STEP',
  /** Stored value while BoardGameGeek game metadata is still pending. */
  GAMES_STEP: 'games',
  /** Stored value after Games finishes and title normalization remains. */
  TITLES_STEP: 'titles',
  /**
   * Resume interval for the next batch. Five minutes balances progress against
   * UrlFetch and trigger quotas without overlapping typical batch runtimes.
   */
  TRIGGER_INTERVAL_MINUTES: 5,
  /**
   * Games processed per invocation. Sized so delayed BoardGameGeek requests
   * usually finish inside one Apps Script execution.
   */
  GAME_BATCH_SIZE: 50,
  /**
   * Titles normalized per invocation. Higher than the game batch because each
   * title fetch is cheaper and paced with a shorter delay.
   */
  TITLE_BATCH_SIZE: 100,
  /**
   * Days a successful or failed game fetch stays ineligible. Failures advance
   * the timestamp too so permanent errors do not monopolize oldest-first batches.
   */
  GAME_REFRESH_INTERVAL_DAYS: 7,
} as const;

/**
 * Legacy handlers that must be removed before the unified update queue starts.
 *
 * Older deployments created one trigger per sheet. Leaving them active would
 * let those handlers write the same spreadsheet while the unified queue runs.
 */
const LEGACY_UPDATE_HANDLER_NAMES = [
  'updateGames',
  'updateRankings',
  'updateTitles',
  'updateRatings',
] as const;

/**
 * Property keys managed in the Apps Script project settings.
 *
 * Values live outside source control so each deployment can use its own
 * BoardGameGeek token and Bodoge user without code changes.
 */
const SCRIPT_PROPERTY_KEYS = {
  /** Optional BoardGameGeek API token that relaxes anonymous request limits. */
  BOARD_GAME_GEEK_TOKEN: 'TOKEN',
  /** Bodoge user ID whose played-game ratings are imported into Ratings. */
  BODOGE_USER_ID: 'BODOGE_USER_ID',
} as const;

/**
 * BoardGameGeek endpoint settings.
 *
 * The thing API returns XML with statistics and player-count polls. Requests
 * are deliberately paced because anonymous and token-authenticated clients
 * both face upstream rate limits.
 */
const BOARD_GAME_GEEK_CONFIG = {
  /** XML API endpoint used to fetch one game's statistics and polls. */
  THING_ENDPOINT: 'https://boardgamegeek.com/xmlapi2/thing',
  /** Lowest player count represented by the spreadsheet's recommendation columns. */
  MIN_SUPPORTED_PLAYER_COUNT: 2,
  /**
   * Consecutive player-count columns written beside that minimum. Matches the
   * spreadsheet's 2–10 layout and Rankings' supported counts; there is no
   * eleventh-player recommendation column.
   */
  SUPPORTED_PLAYER_COUNT_TOTAL: 9,
  /** Delay between requests so retries cannot burst during an upstream outage. */
  REQUEST_DELAY_MILLISECONDS: 2000,
} as const;

/**
 * Board Game Arena endpoints and tag rules.
 *
 * The Japanese home page embeds the full catalog as JSON. Some tags describe
 * platform implementation details rather than game traits, so they are omitted
 * from the Rankings sheet's human-readable tag column.
 */
const BOARD_GAME_ARENA_CONFIG = {
  /** Page whose HTML embeds `game_list` and `game_tags` JSON arrays. */
  HOME_PAGE_URL: 'https://ja.boardgamearena.com',
  /** Base URL used to open a game's panel; Rankings stores the full URL. */
  GAME_PANEL_URL: 'https://ja.boardgamearena.com/gamepanel?game=',
  /** Tags that describe implementation details rather than game characteristics. */
  IGNORED_TAG_IDS: [2, 3, 4, 10, 11, 12, 20, 21, 28, 29, 31, 300, 301],
  /** Player counts exposed as boolean columns in the Rankings sheet. */
  SUPPORTED_PLAYER_COUNTS: [2, 3, 4, 5, 6, 7, 8, 9, 10],
} as const;

/**
 * Board Game Arena page parsing settings for title collection.
 *
 * Titles are scraped from individual game panels because the catalog embed
 * does not expose the Japanese display name used for spreadsheet matching.
 */
const BOARD_GAME_ARENA_TITLE_CONFIG = {
  /** Pattern used to extract the Japanese game name from a game panel page. */
  GAME_NAME_PATTERN:
    /id="game_name" class="block gamename"\s*>(.*?)(\(.*?\))?<\/a/m,
  /** Delay between title page requests to respect the source site. */
  REQUEST_DELAY_MILLISECONDS: 1000,
} as const;

/**
 * Bodoge endpoint and polling settings.
 *
 * Ratings are paginated HTML cards. The user ID is read from script properties
 * so the same project can target different Bodoge accounts per deployment.
 */
const BODOGE_CONFIG = {
  /** URL prefix for the paginated list of played games. */
  PLAYED_GAMES_URL_PREFIX: 'https://bodoge.hoobby.net/friends/',
  /** URL segment between the user ID and page query. */
  PLAYED_GAMES_URL_SUFFIX: '/boardgames/played?page=',
  /** Delay between pages to avoid burst requests to the source site. */
  REQUEST_DELAY_MILLISECONDS: 1000,
  /**
   * Hard stop for ratings pagination.
   *
   * Bodoge should eventually return its empty-list marker, but a stuck or
   * repeating card page must not consume the entire Apps Script runtime and
   * block Games/Titles trigger creation in the same update cycle.
   */
  MAX_PAGE_COUNT: 100,
} as const;

/**
 * Custom spreadsheet menu settings.
 *
 * Labels are configuration so the spreadsheet UI can change without hunting
 * through the Apps Script entry point.
 */
const MENU_CONFIG = {
  /** Menu label shown in the spreadsheet UI. */
  NAME: 'Functions',
  /** Label for the unified update command. */
  UPDATE_ITEM_LABEL: 'Update',
} as const;
