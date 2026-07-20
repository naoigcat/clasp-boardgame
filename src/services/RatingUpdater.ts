/**
 * A title and rating pair written into the Ratings sheet.
 */
type RatingSheetRow = [string, string];

/**
 * Parsed content from one Bodoge ratings page.
 */
interface RatingPage {
  /** Whether the page contains cards, which signals that pagination continues. */
  readonly hasCards: boolean;
  /** Ratings successfully extracted from the page. */
  readonly rows: readonly RatingSheetRow[];
}

/**
 * Imports a configured Bodoge user's played-game ratings.
 */
class RatingUpdater {
  /**
   * Fetches every ratings page and replaces the Ratings sheet contents.
   */
  static run(): void {
    const sheet = findSheet(SHEET_NAMES.RATINGS);
    if (sheet === null) {
      return;
    }

    const userId = ScriptPropertyStore.getOptionalValue(
      SCRIPT_PROPERTY_KEYS.BODOGE_USER_ID,
    );
    if (userId === null || userId.length === 0) {
      return;
    }

    const rows = RatingUpdater.fetchAllRows(userId);
    // Fetch before clearing so a request failure on a later page preserves the
    // last complete ratings snapshot instead of publishing a partial import.
    rows.sort(([firstTitle], [secondTitle]) =>
      firstTitle > secondTitle ? 1 : firstTitle < secondTitle ? -1 : 0,
    );

    clearSheetDataRows(sheet, SHEET_LAYOUT.RATING_COLUMN_COUNT);
    if (rows.length === 0) {
      return;
    }

    sheet
      .getRange(
        SHEET_LAYOUT.FIRST_DATA_ROW,
        SHEET_LAYOUT.DEFAULT_START_COLUMN,
        rows.length,
        SHEET_LAYOUT.RATING_COLUMN_COUNT,
      )
      .setValues(rows);
  }

  /**
   * Fetches and combines Bodoge pages until the first page without game cards.
   */
  private static fetchAllRows(userId: string): RatingSheetRow[] {
    const rows: RatingSheetRow[] = [];

    for (let pageNumber = 1; ; pageNumber += 1) {
      const response = HttpClient.get(
        RatingUpdater.buildPageUrl(userId, pageNumber),
      );
      if (response.getResponseCode() !== 200) {
        throw new Error(`Bodoge returned HTTP ${response.getResponseCode()}`);
      }

      const page = RatingUpdater.parsePage(response.getContentText());
      if (!page.hasCards) {
        return rows;
      }

      rows.push(...page.rows);
      Utilities.sleep(BODOGE_CONFIG.REQUEST_DELAY_MILLISECONDS);
    }
  }

  /**
   * Builds the URL for one page of a user's played-game ratings.
   */
  private static buildPageUrl(userId: string, pageNumber: number): string {
    return `${BODOGE_CONFIG.PLAYED_GAMES_URL_PREFIX}${encodeURIComponent(
      userId,
    )}${BODOGE_CONFIG.PLAYED_GAMES_URL_SUFFIX}${pageNumber}`;
  }

  /**
   * Parses rating cards from one Bodoge page.
   */
  private static parsePage(pageHtml: string): RatingPage {
    const cards =
      pageHtml.match(/<a class="list--interests-item-title".*?<\/a>/g) ?? [];
    const rows: RatingSheetRow[] = [];

    cards.forEach((card) => {
      const sourceTitle = RatingUpdater.extractSourceTitle(card);
      if (sourceTitle === null) {
        return;
      }

      const rating = RatingUpdater.extractRating(card);
      RatingUpdater.expandTitleAliases(sourceTitle).forEach((title) => {
        rows.push([title, rating]);
      });
    });

    return { hasCards: cards.length > 0, rows };
  }

  /**
   * Extracts and normalizes a Japanese title from one Bodoge rating card.
   */
  private static extractSourceTitle(cardHtml: string): string | null {
    const titleMatch = cardHtml.match(
      '<div class="list--interests-item-title-japanese">(.*?)</div>',
    );
    if (!titleMatch?.[1]) {
      return null;
    }

    return titleMatch[1]
      .split('/')[0]
      .replace(/（.*）/, '')
      .replace('：新版', '')
      .replace('（拡張）', '')
      .replace('&amp;', '＆')
      .trim();
  }

  /**
   * Extracts the star-rating value from one Bodoge rating card.
   */
  private static extractRating(cardHtml: string): string {
    const ratingMatch = cardHtml.match(
      '<div class="rating--result-stars" data-rating-mode="result" data-rating-result="(.*?)">',
    );
    return ratingMatch?.[1] ?? '';
  }

  /**
   * Expands a source title to its one or more canonical Ratings-sheet titles.
   */
  private static expandTitleAliases(sourceTitle: string): readonly string[] {
    if (EXCLUDED_RATING_TITLES.includes(sourceTitle)) {
      return [];
    }

    return RATING_TITLE_ALIASES[sourceTitle] ?? [sourceTitle];
  }
}
