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

      const page = RatingParser.parsePage(response.getContentText());
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
}
