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
 * Markup Bodoge renders for an empty played-games page or a page past the last
 * result. Absence of both this marker and rating cards means the HTML is not a
 * recognized ratings list page, so the import aborts instead of clearing Ratings.
 */
const BODOGE_EMPTY_PLAYED_GAMES_MARKER =
  '<p class="empty">検索結果が存在しないか、マイボードゲームが未登録のユーザーです</p>';

/**
 * Imports a configured Bodoge user's played-game ratings.
 *
 * The sheet is replaced only after every page has been fetched successfully so
 * a mid-import HTTP or HTML failure preserves the previous complete snapshot.
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
    // Without a configured user there is nothing to import; skip quietly so
    // deployments that only use Rankings/Games still run the rest of update.
    if (userId === null || userId.length === 0) {
      return;
    }

    const rows = RatingUpdater.fetchAllRows(userId);
    // Fetch before clearing so a request failure on a later page preserves the
    // last complete ratings snapshot instead of publishing a partial import.
    // Stable alphabetical order makes sheet diffs and formula lookups predictable.
    rows.sort(([firstTitle], [secondTitle]) =>
      firstTitle > secondTitle ? 1 : firstTitle < secondTitle ? -1 : 0,
    );

    // An explicit empty Bodoge list is the only case that may wipe Ratings.
    // Non-empty imports write first, then trim surplus, so a failed setValues
    // leaves the previous complete snapshot instead of a blank sheet.
    if (rows.length === 0) {
      clearSheetDataRows(sheet, SHEET_LAYOUT.RATING_COLUMN_COUNT);
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
    clearSurplusSheetDataRows(
      sheet,
      rows.length,
      SHEET_LAYOUT.RATING_COLUMN_COUNT,
    );
  }

  /**
   * Fetches and combines Bodoge pages until the first explicit empty page.
   *
   * Pagination normally ends on Bodoge's empty-list marker. A page cap aborts
   * runaway card responses so this import cannot exhaust the Apps Script
   * runtime before the rest of the update cycle schedules Games/Titles work.
   */
  private static fetchAllRows(userId: string): RatingSheetRow[] {
    const rows: RatingSheetRow[] = [];

    for (
      let pageNumber = 1;
      pageNumber <= BODOGE_CONFIG.MAX_PAGE_COUNT;
      pageNumber += 1
    ) {
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

    // Abort without writing so the previous complete Ratings snapshot remains.
    throw new Error(
      `Bodoge ratings pagination exceeded ${BODOGE_CONFIG.MAX_PAGE_COUNT} pages`,
    );
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
   *
   * Pages without cards must include Bodoge's empty-list marker; otherwise the
   * HTML is treated as unrecognized so callers can keep the prior snapshot.
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
      // Bundled products may expand into multiple spreadsheet titles.
      RatingUpdater.expandTitleAliases(sourceTitle).forEach((title) => {
        rows.push([title, rating]);
      });
    });

    if (cards.length > 0) {
      return { hasCards: true, rows };
    }

    if (pageHtml.includes(BODOGE_EMPTY_PLAYED_GAMES_MARKER)) {
      return { hasCards: false, rows: [] };
    }

    throw new Error('Unrecognized Bodoge ratings page HTML');
  }

  /**
   * Extracts and normalizes a Japanese title from one Bodoge rating card.
   *
   * Bodoge cards often include English/Japanese pairs, edition labels, and
   * expansion markers that would break joins against the spreadsheet's titles.
   */
  private static extractSourceTitle(cardHtml: string): string | null {
    const titleMatch = cardHtml.match(
      '<div class="list--interests-item-title-japanese">(.*?)</div>',
    );
    if (!titleMatch?.[1]) {
      return null;
    }

    return (
      titleMatch[1]
        // Keep the Japanese segment when English and Japanese are slash-separated.
        .split('/')[0]
        // Drop parenthetical notes that are not part of the canonical title.
        .replace(/（.*）/, '')
        .replace('：新版', '')
        .replace('（拡張）', '')
        .replace('&amp;', '＆')
        .trim()
    );
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
   *
   * Exclusions and aliases encode spreadsheet-specific naming decisions that
   * Bodoge's catalog cannot express on its own.
   */
  private static expandTitleAliases(sourceTitle: string): readonly string[] {
    if (EXCLUDED_RATING_TITLES.includes(sourceTitle)) {
      return [];
    }

    return RATING_TITLE_ALIASES[sourceTitle] ?? [sourceTitle];
  }
}
