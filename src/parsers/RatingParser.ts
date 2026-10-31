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

/** Parses Bodoge markup without fetching pages or changing sheets. */
class RatingParser {
  /**
   * Parses rating cards from one Bodoge page.
   *
   * Pages without cards must include Bodoge's empty-list marker; otherwise the
   * HTML is treated as unrecognized so callers can keep the prior snapshot.
   */
  static parsePage(pageHtml: string): RatingPage {
    // [\s\S] so pretty-printed Bodoge HTML with newlines between tags still matches.
    const cards =
      pageHtml.match(
        /<div class="list--interests-item">[\s\S]*?(?=<div class="list--interests-item">|$)/g,
      ) ?? [];
    const rows: RatingSheetRow[] = [];
    let hasUnratedCards = false;
    let hasUnextractableCards = false;

    cards.forEach((card) => {
      const sourceTitle = RatingParser.extractSourceTitle(card);
      if (sourceTitle === null) {
        // Card matched but title extraction failed; callers must not treat this
        // the same as an excluded title (which still extracts, then drops).
        hasUnextractableCards = true;
        return;
      }

      const rating = RatingParser.extractRating(card) ?? '';

      const interests = RatingParser.extractInterests(card);
      const [have, favorite, played, watching] = interests;
      // "0" is Bodoge's unrated/privacy-hidden marker, not a score.
      const isUnrated = rating === '' || rating === '0';
      if (isUnrated) {
        hasUnratedCards = true;
        if (!interests.some(Boolean)) {
          return;
        }
      }

      // Bundled products may expand into multiple spreadsheet titles.
      RatingParser.expandTitleAliases(sourceTitle).forEach((title) => {
        rows.push([
          title,
          isUnrated ? '' : rating,
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
      // Remove trailing edition labels, including their separator.
      .replace(
        /[\s：:]*(?:第[\d０-９一二三四五六七八九十百]+版|新版|日本語版)\s*$/,
        '',
      )
      // Version subtitles use a space; other subtitle colons remain intact.
      .replace(/[：:]\s*([^：:]+\bver\.?)\s*$/i, ' $1')
      .replace(/\bver\.?\s*$/i, 'バージョン')
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

  /** Reads all registration flags, rejecting partial or unrecognized cards. */
  private static extractInterests(cardHtml: string): boolean[] {
    return ['have', 'favorite', 'played', 'watching'].map((interest) => {
      const match = cardHtml.match(
        new RegExp(`class="friend-interest ${interest} (on|off)"`),
      );
      if (!match) {
        throw new Error(
          'Bodoge ratings page contained cards without extractable interests',
        );
      }
      return match[1] === 'on';
    });
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
