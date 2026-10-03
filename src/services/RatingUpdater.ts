/**
 * A title, optional rating, and registration flags written into Ratings.
 */
type RatingSheetRow = [string, string, boolean, boolean, boolean, boolean];

/**
 * Parsed content from one Bodoge ratings page.
 */
interface RatingPage {
  /** Whether the page contains cards, which signals that pagination continues. */
  readonly hasCards: boolean;
  /** Whether at least one card reported no rating (missing/empty/"0"). */
  readonly hasUnratedCards: boolean;
  /**
   * True when at least one card matched but had no Japanese title to extract.
   * Distinct from an all-excluded page, which has cards and zero rows without
   * setting this flag, so pagination can continue without treating exclusions
   * as a markup failure.
   */
  readonly hasUnextractableCards: boolean;
  /** Ratings successfully extracted from the page after exclusions/aliases. */
  readonly rows: readonly RatingSheetRow[];
}

/**
 * Markup Bodoge renders for an empty registered-games page or a page past the last
 * result. Absence of both this marker and rating cards means the HTML is not a
 * recognized ratings list page, so the import aborts instead of clearing Ratings.
 */
const BODOGE_EMPTY_GAMES_MARKER =
  '<p class="empty">検索結果が存在しないか、マイボードゲームが未登録のユーザーです</p>';

/**
 * Imports a configured Bodoge user's ratings and registered games.
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
      // Sync import: log and skip so a renamed Ratings tab is visible in the
      // execution log instead of looking like a successful no-op.
      Logger.log(
        `Ratings sheet "${SHEET_NAMES.RATINGS}" is missing; skipping rating import until the tab is restored.`,
      );
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

    // fetchAllRows returns [] only for Bodoge's explicit empty-list marker;
    // writeSheetSnapshot clears it without issuing an invalid zero-row write.
    writeSheetSnapshot(sheet, rows, SHEET_LAYOUT.RATING_COLUMN_COUNT);
  }

  /**
   * Fetches and combines Bodoge pages until the first explicit empty page.
   *
   * Pagination normally ends on Bodoge's empty-list marker. Card pages are
   * capped so a stuck or repeating response cannot exhaust the Apps Script
   * runtime before the rest of the update cycle schedules Games/Titles work.
   * One extra fetch is allowed after that cap so a collection that fills
   * exactly MAX_PAGE_COUNT card pages can still observe the empty marker and
   * publish; cards beyond the cap still abort without writing.
   */
  private static fetchAllRows(userId: string): RatingSheetRow[] {
    const rows: RatingSheetRow[] = [];
    // Distinguishes Bodoge's explicit empty list from card markup that matched
    // but produced no importable rows (missing titles, exclusions, or no ratings/flags).
    let sawCards = false;
    let sawUnratedCards = false;

    // +1 lets a full MAX_PAGE_COUNT card collection still fetch the empty
    // marker that ends pagination; card pages past the cap throw below.
    for (
      let pageNumber = 1;
      pageNumber <= BODOGE_CONFIG.MAX_PAGE_COUNT + 1;
      pageNumber += 1
    ) {
      const response = HttpClient.get(
        RatingUpdater.buildPageUrl(userId, pageNumber),
      );
      HttpClient.requireSuccess(response, 'Bodoge');

      const page = RatingUpdater.parsePage(response.getContentText());
      if (!page.hasCards) {
        // Outer card wrappers without extractable titles must not clear Ratings;
        // an unrated card also allows a valid zero-row result.
        if (sawCards && rows.length === 0 && !sawUnratedCards) {
          throw new Error(
            'Bodoge ratings pages contained cards but yielded no importable titles',
          );
        }
        return rows;
      }

      // Still receiving cards after the budget means runaway pagination; keep
      // the previous complete Ratings snapshot instead of writing a truncated one.
      if (pageNumber > BODOGE_CONFIG.MAX_PAGE_COUNT) {
        throw new Error(
          `Bodoge ratings pagination exceeded ${BODOGE_CONFIG.MAX_PAGE_COUNT} pages`,
        );
      }

      sawCards = true;
      sawUnratedCards = sawUnratedCards || page.hasUnratedCards;
      // A later page with card chrome but no Japanese titles must not let earlier
      // pages publish a partial Ratings snapshot; fail the whole import instead.
      // Excluded-title-only pages still report hasUnextractableCards=false so
      // pagination can skip them without treating intentional omissions as loss.
      if (page.hasUnextractableCards) {
        throw new Error(
          'Bodoge ratings page contained cards without extractable Japanese titles',
        );
      }
      rows.push(...page.rows);
      Utilities.sleep(BODOGE_CONFIG.REQUEST_DELAY_MILLISECONDS);
    }

    // Unreachable while the loop bound is MAX_PAGE_COUNT + 1 and card pages
    // past the cap throw; kept so a future bound change cannot silently return.
    throw new Error(
      `Bodoge ratings pagination exceeded ${BODOGE_CONFIG.MAX_PAGE_COUNT} pages`,
    );
  }

  /**
   * Builds the URL for one page of a user's registered games.
   */
  private static buildPageUrl(userId: string, pageNumber: number): string {
    return `${BODOGE_CONFIG.GAMES_URL_PREFIX}${encodeURIComponent(
      userId,
    )}${BODOGE_CONFIG.GAMES_URL_SUFFIX}${pageNumber}`;
  }

  /**
   * Parses rating cards from one Bodoge page.
   *
   * Pages without cards must include Bodoge's empty-list marker; otherwise the
   * HTML is treated as unrecognized so callers can keep the prior snapshot.
   */
  private static parsePage(pageHtml: string): RatingPage {
    // [\s\S] so pretty-printed Bodoge HTML with newlines between tags still matches.
    const cards =
      pageHtml.match(
        /<div class="list--interests-item">[\s\S]*?(?=<div class="list--interests-item">|$)/g,
      ) ?? [];
    const rows: RatingSheetRow[] = [];
    let hasUnratedCards = false;
    let hasUnextractableCards = false;

    cards.forEach((card) => {
      const sourceTitle = RatingUpdater.extractSourceTitle(card);
      if (sourceTitle === null) {
        // Card matched but title extraction failed; callers must not treat this
        // the same as an excluded title (which still extracts, then drops).
        hasUnextractableCards = true;
        return;
      }

      const rating = RatingUpdater.extractRating(card);
      const interests = ['have', 'favorite', 'played', 'watching'].map(
        (interest) => {
          const match = card.match(
            new RegExp(`class="friend-interest ${interest} (on|off)"`),
          );
          if (!match) {
            throw new Error(
              'Bodoge ratings page contained cards without extractable interests',
            );
          }
          return match[1] === 'on';
        },
      );
      const [have, favorite, played, watching] = interests;
      // "0" is Bodoge's unrated/privacy-hidden marker, not a score.
      if (rating === null || rating === '0') {
        hasUnratedCards = true;
        if (!interests.some(Boolean)) {
          return;
        }
      }

      // Bundled products may expand into multiple spreadsheet titles.
      RatingUpdater.expandTitleAliases(sourceTitle).forEach((title) => {
        rows.push([
          title,
          rating === null || rating === '0' ? '' : rating,
          have,
          favorite,
          played,
          watching,
        ]);
      });
    });

    if (cards.length > 0) {
      return {
        hasCards: true,
        hasUnratedCards,
        hasUnextractableCards,
        rows,
      };
    }

    if (pageHtml.includes(BODOGE_EMPTY_GAMES_MARKER)) {
      return {
        hasCards: false,
        hasUnratedCards: false,
        hasUnextractableCards: false,
        rows: [],
      };
    }

    throw new Error('Unrecognized Bodoge ratings page HTML');
  }

  /**
   * Extracts and normalizes a Japanese title from one Bodoge rating card.
   *
   * Bodoge cards often include English/Japanese pairs, edition labels, and
   * expansion markers that would break joins against the spreadsheet's titles.
   * Returns null when cleanup leaves nothing importable so callers can abort
   * instead of writing blank title rows that wipe the prior Ratings snapshot.
   */
  private static extractSourceTitle(cardHtml: string): string | null {
    const titleMatch = cardHtml.match(
      '<div class="list--interests-item-title-japanese">(.*?)</div>',
    );
    if (!titleMatch?.[1]) {
      return null;
    }

    const cleanedTitle = titleMatch[1]
      // Keep the Japanese segment when English and Japanese are slash-separated.
      .split('/')[0]
      // Drop parenthetical notes that are not part of the canonical title.
      .replace(/（.*）/, '')
      .replace('：新版', '')
      .replace('（拡張）', '')
      .replace('&amp;', '＆')
      .trim();

    // Edition/expansion-only (or otherwise empty) cleanup results are not
    // importable titles; treat them like a missing Japanese title.
    if (cleanedTitle.length === 0) {
      return null;
    }

    return cleanedTitle;
  }

  /**
   * Extracts the star-rating value from one Bodoge rating card.
   *
   * Returns null for missing/empty ratings and "0" for unrated/privacy-hidden
   * cards. Both leave column B blank when registration flags qualify the row.
   */
  private static extractRating(cardHtml: string): string | null {
    const ratingMatch = cardHtml.match(
      '<div class="rating--result-stars" data-rating-mode="result" data-rating-result="(.*?)">',
    );
    const rating = ratingMatch?.[1];
    if (rating === undefined || rating.length === 0) {
      return null;
    }
    return rating;
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
