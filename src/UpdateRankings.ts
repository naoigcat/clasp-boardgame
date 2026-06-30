class UpdateRankings {
  static run(): void {
    let rankings =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rankings');
    if (rankings === null) {
      return;
    }
    let html = fetch('https://ja.boardgamearena.com').getContentText();
    let tagMatches =
      (html.match(/"game_tags":([\s\S]*),\n?\s*"top_tags"/m) || [])[1].match(
        /\{"id":[\s\S]*?\}/gm,
      ) || [];
    let tagMaster: { [key: string]: string } = {};
    for (let index = 0; index < tagMatches.length; index++) {
      let tag: { [key: string]: any };
      try {
        tag = JSON.parse(tagMatches[index]);
      } catch (e: unknown) {
        if (e instanceof Error) {
          Logger.log(`Error: ${e.message}\n${tagMatches[index]}`);
        } else {
          Logger.log(`Unknown error: ${String(e)}\n${tagMatches[index]}`);
        }
        throw e;
      }
      tagMaster[tag['id']] = tag['name'];
    }
    let gameMatches =
      (html.match(/"game_list":([\s\S]*),\n?\s*"game_tags"/m) || [])[1].match(
        /\{"id":[\s\S]*?"watched":[\s\S]*?\}/gm,
      ) || [];
    let games = [];
    for (let index = 0; index < gameMatches.length; index++) {
      let game: { [key: string]: any };
      try {
        game = JSON.parse(gameMatches[index]);
      } catch (e: unknown) {
        if (e instanceof Error) {
          Logger.log(`Error: ${e.message}\n${gameMatches[index]}`);
        } else {
          Logger.log(`Unknown error: ${String(e)}\n${gameMatches[index]}`);
        }
        throw e;
      }
      let tags = [];
      for (let tagIndex = 0; tagIndex < game['tags'].length; tagIndex++) {
        let tagNumber = game['tags'][tagIndex][0];
        switch (tagNumber) {
          case 2: // 難易度:易しい
          case 3: // 難易度:普通
          case 4: // 難易度:難しい
          case 10: // 短時間ゲーム
          case 11: // 並の長さのゲーム
          case 12: // 長時間ゲーム
          case 20: // 賞を受けたゲーム
          case 21: // 新しい
          case 28: // リアルタイム推奨
          case 29: // ターンベース推奨
          case 31: // モバイルでも良好
          case 300: // Tags checked
          case 301: // PHP8
            continue;
        }
        tags.push(tagMaster[tagNumber]);
      }
      games.push([
        `https://ja.boardgamearena.com/gamepanel?game=${game['name']}`,
        null,
        tags.join(' '),
        game['games_played'],
        game['average_duration'],
        game['default_num_players'],
        game['player_numbers'].includes(2),
        game['player_numbers'].includes(3),
        game['player_numbers'].includes(4),
        game['player_numbers'].includes(5),
        game['player_numbers'].includes(6),
        game['player_numbers'].includes(7),
        game['player_numbers'].includes(8),
        game['player_numbers'].includes(9),
        game['player_numbers'].includes(10),
      ]);
    }
    if (games.length === 0) {
      return;
    }
    rankings
      .getRange(2, 1, rankings.getLastRow() - 1, games[0].length)
      .clearContent();
    rankings.getRange(2, 1, games.length, games[0].length).setValues(games);
  }
}

function updateRankings(): void {
  UpdateRankings.run();
}
