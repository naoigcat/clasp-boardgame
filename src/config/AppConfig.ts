/**
 * Spreadsheet tab names used by the update services.
 *
 * Keeping these names together makes a renamed tab an explicit configuration
 * change instead of a hidden dependency inside a service.
 */
const SHEET_NAMES = {
  /** Stores BoardGameGeek game metadata. */
  GAMES: 'Games',
  /** Stores games fetched from Board Game Arena. */
  RANKINGS: 'Rankings',
  /** Stores the original and normalized game titles. */
  TITLES: 'Titles',
  /** Stores ratings imported from Bodoge. */
  RATINGS: 'Ratings',
} as const;

/**
 * Settings for the single trigger that advances a multi-run update.
 *
 * A shared trigger keeps Apps Script trigger quotas predictable while batches
 * are still pending.
 */
const UPDATE_QUEUE_CONFIG = {
  /** Global function invoked by the time-driven trigger. */
  HANDLER_NAME: 'update',
  /** Script-property key that records the next asynchronous phase. */
  STEP_PROPERTY_KEY: 'UPDATE_STEP',
  /** Stored value for the BoardGameGeek game-metadata phase. */
  GAMES_STEP: 'games',
  /** Stored value for the title-normalization phase. */
  TITLES_STEP: 'titles',
  /** Interval used to resume a pending phase. */
  TRIGGER_INTERVAL_MINUTES: 5,
  /** Maximum number of BoardGameGeek games updated in one invocation. */
  GAME_BATCH_SIZE: 50,
  /** Maximum number of titles normalized in one invocation. */
  TITLE_BATCH_SIZE: 100,
  /** Number of days a BoardGameGeek response remains fresh. */
  GAME_REFRESH_INTERVAL_DAYS: 7,
} as const;

/**
 * Legacy handlers that must be removed before the unified update queue starts.
 *
 * Removing them prevents an older deployment from racing with the new queue.
 */
const LEGACY_UPDATE_HANDLER_NAMES = [
  'updateGames',
  'updateRankings',
  'updateTitles',
  'updateRatings',
] as const;

/**
 * Property keys managed in the Apps Script project settings.
 */
const SCRIPT_PROPERTY_KEYS = {
  /** Optional BoardGameGeek API token. */
  BOARD_GAME_GEEK_TOKEN: 'TOKEN',
  /** Bodoge user ID whose ratings are imported. */
  BODOGE_USER_ID: 'BODOGE_USER_ID',
} as const;

/**
 * BoardGameGeek endpoint settings.
 */
const BOARD_GAME_GEEK_CONFIG = {
  /** XML API endpoint used to fetch a game's statistics. */
  THING_ENDPOINT: 'https://boardgamegeek.com/xmlapi2/thing',
  /** Lowest player count represented by the spreadsheet's recommendation columns. */
  MIN_SUPPORTED_PLAYER_COUNT: 2,
  /** Number of consecutive player-count columns written to the spreadsheet. */
  SUPPORTED_PLAYER_COUNT_TOTAL: 10,
  /** Delay between requests to respect the upstream API. */
  REQUEST_DELAY_MILLISECONDS: 2000,
} as const;

/**
 * Board Game Arena endpoints and tag rules.
 */
const BOARD_GAME_ARENA_CONFIG = {
  /** Page that embeds the game and tag metadata. */
  HOME_PAGE_URL: 'https://ja.boardgamearena.com',
  /** Base URL used to open a game's panel. */
  GAME_PANEL_URL: 'https://ja.boardgamearena.com/gamepanel?game=',
  /** Tags that describe implementation details rather than game characteristics. */
  IGNORED_TAG_IDS: [2, 3, 4, 10, 11, 12, 20, 21, 28, 29, 31, 300, 301],
  /** Player counts exposed as boolean columns in the Rankings sheet. */
  SUPPORTED_PLAYER_COUNTS: [2, 3, 4, 5, 6, 7, 8, 9, 10],
} as const;

/**
 * Board Game Arena page parsing settings for title collection.
 */
const BOARD_GAME_ARENA_TITLE_CONFIG = {
  /** Pattern used to extract the Japanese game name from a game panel page. */
  GAME_NAME_PATTERN:
    /id="game_name" class="block gamename"\n\s*>(.*?)(\(.*?\))?<\/a/m,
  /** Delay between title page requests to respect the source site. */
  REQUEST_DELAY_MILLISECONDS: 1000,
} as const;

/**
 * Bodoge endpoint and polling settings.
 */
const BODOGE_CONFIG = {
  /** URL prefix for the paginated list of played games. */
  PLAYED_GAMES_URL_PREFIX: 'https://bodoge.hoobby.net/friends/',
  /** URL segment between the user ID and page query. */
  PLAYED_GAMES_URL_SUFFIX: '/boardgames/played?page=',
  /** Delay between pages to avoid burst requests to the source site. */
  REQUEST_DELAY_MILLISECONDS: 1000,
} as const;

/**
 * Custom spreadsheet menu settings.
 */
const MENU_CONFIG = {
  /** Menu label shown in the spreadsheet UI. */
  NAME: 'Functions',
  /** Label for the unified update command. */
  UPDATE_ITEM_LABEL: 'Update',
} as const;
