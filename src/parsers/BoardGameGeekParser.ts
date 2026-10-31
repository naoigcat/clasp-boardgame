/**
 * Maps a player count to BoardGameGeek's preferred recommendation label.
 */
type PlayerRecommendations = Record<string, string>;

/** Metadata is fully parsed before any Games cells are mutated. */
interface BoardGameGeekMetadata {
  readonly recommendations: PlayerRecommendations;
  readonly boardGameRank: NumericDisplayValue;
  readonly bayesAverage: NumericDisplayValue;
  readonly averageWeight: NumericDisplayValue;
  readonly playTime: NumericDisplayValue | string;
  readonly publicationYear: NumericDisplayValue;
}

/** Parses BoardGameGeek XML independently of HTTP and spreadsheet writes. */
class BoardGameGeekParser {
  static parse(gameItem: XmlElement, gameId: string): BoardGameGeekMetadata {
    const recommendations = BoardGameGeekParser.getPlayerRecommendations(
      gameItem,
      gameId,
    );
    const ratings = getRequiredXmlChild(
      getRequiredXmlChild(gameItem, 'statistics'),
      'ratings',
    );
    const ranks = getRequiredXmlChild(ratings, 'ranks');
    // Prefer the overall board-game rank over family-specific rank entries.
    const boardGameRank = findXmlElementByAttribute(
      ranks.getChildren('rank'),
      'name',
      'boardgame',
    );
    const boardGameRankValue = parseDisplayNumber(
      getRequiredXmlAttributeValue(boardGameRank, 'value'),
    );
    const bayesAverage = parseDisplayNumber(
      getRequiredXmlAttributeValue(
        getRequiredXmlChild(ratings, 'bayesaverage'),
        'value',
      ),
    );
    const averageWeight = parseDisplayNumber(
      getRequiredXmlAttributeValue(
        getRequiredXmlChild(ratings, 'averageweight'),
        'value',
      ),
    );
    const playTime = BoardGameGeekParser.formatPlayTime(
      getRequiredXmlAttributeValue(
        getRequiredXmlChild(gameItem, 'minplaytime'),
        'value',
      ),
      getRequiredXmlAttributeValue(
        getRequiredXmlChild(gameItem, 'maxplaytime'),
        'value',
      ),
    );
    const publicationYear = parseDisplayNumber(
      getRequiredXmlAttributeValue(
        getRequiredXmlChild(gameItem, 'yearpublished'),
        'value',
      ),
    );

    return {
      recommendations,
      boardGameRank: boardGameRankValue,
      bayesAverage,
      averageWeight,
      playTime,
      publicationYear,
    };
  }

  /**
   * Reads preferred player counts from the BoardGameGeek poll and applies known
   * per-game corrections.
   *
   * For each player count, the result with the most votes wins. Overrides are
   * applied last so curated corrections beat the raw community poll.
   */
  private static getPlayerRecommendations(
    gameItem: XmlElement,
    gameId: string,
  ): PlayerRecommendations {
    const poll = findXmlElementByAttribute(
      gameItem.getChildren('poll'),
      'name',
      'suggested_numplayers',
    );
    const recommendations: PlayerRecommendations = {};

    poll.getChildren('results').forEach((results) => {
      const mostVotedResult = sortXmlElementsByNumericAttribute(
        results.getChildren('result'),
        'numvotes',
      )[0];
      if (mostVotedResult === undefined) {
        return;
      }

      recommendations[getRequiredXmlAttributeValue(results, 'numplayers')] =
        getRequiredXmlAttributeValue(mostVotedResult, 'value');
    });

    return {
      ...recommendations,
      ...(GAME_PLAYER_COUNT_OVERRIDES[gameId] ?? {}),
    };
  }

  /**
   * Formats equal play times as one value and ranges as a readable interval.
   */
  private static formatPlayTime(
    minimumValue: string,
    maximumValue: string,
  ): NumericDisplayValue | string {
    const minimum = parseDisplayNumber(minimumValue);
    const maximum = parseDisplayNumber(maximumValue);
    return minimum === maximum ? minimum : `${minimum}-${maximum}`;
  }
}
