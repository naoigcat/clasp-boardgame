function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Functions')
    .addItem('Update Games', 'updateGames')
    .addItem('Update Arena Rankings', 'updateArenaRankings')
    .addItem('Update Arena Titles', 'updateArenaTitles')
    .addItem('Update Ratings', 'updateRatings')
    .addToUi();
}

const $ = {
  __: 0,
  _A: 1,
  _B: 2,
  _C: 3,
  _D: 4,
  _E: 5,
  _F: 6,
  _G: 7,
  _H: 8,
  _I: 9,
  _J: 10,
  _K: 11,
  _L: 12,
  _M: 13,
  _N: 14,
  _O: 15,
  _P: 16,
  _Q: 17,
  _R: 18,
  _S: 19,
  _T: 20,
  _U: 21,
  _V: 22,
  _W: 23,
  _X: 24,
  _Y: 25,
  _Z: 26,
} as const;

function getToken(): string | null {
  return PropertiesService.getScriptProperties().getProperty('TOKEN');
}

function fetch(
  url: string,
  params: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {}
): GoogleAppsScript.URL_Fetch.HTTPResponse {
  return UrlFetchApp.fetch(url, params);
}

function fetchWithAuth(
  url: string,
  params: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {}
): GoogleAppsScript.URL_Fetch.HTTPResponse {
  const token = getToken();
  const headers: { [key: string]: string } = {
    ...(params.headers as { [key: string]: string } | undefined),
  };
  if (token) {
    headers['Authorization'] = `Bearer ${token}`;
  }
  return UrlFetchApp.fetch(url, { ...params, headers });
}

// Game-specific overrides for player count recommendations
const GAME_OVERRIDES: { [gameId: string]: { [playerCount: string]: string } } =
  {
    '8172': {
      '7': 'Recommended',
      '8': 'Recommended',
      '9': 'Recommended',
      '10': 'Recommended',
    },
  };

function updateGames() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Games');
  if (sheet === null) {
    return;
  }
  let rows: any[][] = sheet
    .getRange('$A$2:$A')
    .getRichTextValues()
    .map((row: any[], index: number) => {
      row.unshift(index);
      return row;
    });
  sheet
    .getRange('$B$2:$Z')
    .getValues()
    .forEach((row: any[], index: number) => {
      rows[index] = rows[index].concat(row);
    });
  let current = new Date();
  let count = 0;
  const startTime = Date.now();
  const timeLimitMs = 5 * 60 * 1000;
  let errors: string[] = [];
  try {
    rows = rows
      .slice(
        0,
        rows.findIndex((row: any[]) => row[$._A].getText().length === 0)
      )
      .sort((a: any[], b: any[]) => {
        return a[$._Z] < b[$._Z] ? -1 : a[$._Z] > b[$._Z] ? 1 : 0;
      })
      .map((row: any[]) => {
        // Clear columns containing values by ARRAYFORMULA
        [$._C, $._F, $._W, $._X, $._Y].forEach((index) => {
          row[index] = '';
        });
        // Stop processing if the execution window has exceeded five minutes
        const elapsedMs = Date.now() - startTime;
        if (elapsedMs > timeLimitMs) {
          const identifier =
            row[$._A] && typeof row[$._A].getText === 'function'
              ? row[$._A].getText()
              : `index ${row[$.__]}`;
          const message = `updateGames stopped after ${elapsedMs}ms (limit ${timeLimitMs}ms) at ${identifier}`;
          Logger.log(message);
          return row;
        }
        Logger.log(row[$._A].getText());
        const url = row[$._A].getLinkUrl();
        if (url === null) {
          return row;
        }
        let updated = row[$._Z] as Date;
        // Skip if you have been running the API within the past week
        if (updated && updated.addDays(7) > current) {
          return row;
        }
        try {
          const type = url.split('/')[3];
          const id = url.split('/')[4];
          const endpoint = `https://boardgamegeek.com/xmlapi2/thing?type=${type}&stats=1&id=${id}`;
          Logger.log(endpoint);
          const response = fetchWithAuth(endpoint);
          Utilities.sleep(2000);
          count++;
          if (response.getResponseCode() !== 200) {
            return row;
          }
          const body = response.getContentText();
          const item = XmlService.parse(body).getRootElement().getChild('item');
          if (item === null) {
            Logger.log('item is null');
            Logger.log(body);
            return row;
          }
          let numbers = item
            .getChildren('poll')
            .findAttribute('name', 'suggested_numplayers')
            .getChildren('results')
            .reduce((acc: any, results: any) => {
              let numvotes = results
                .getChildren('result')
                .sortAttribute('numvotes')[0];
              if (numvotes === undefined) {
                return acc;
              }
              acc[results.getAttribute('numplayers').getValue()] = numvotes
                .getAttribute('value')
                .getValue();
              return acc;
            }, {});
          Logger.log(numbers);
          // Apply game-specific overrides if they exist
          const overrides = GAME_OVERRIDES[id];
          if (overrides) {
            numbers = { ...numbers, ...overrides };
          }
          const indexes = [...Array(10)].map((v, i) => i + $._I);
          indexes.forEach((index) => {
            row[index] = numbers[(index - $._G).toString()];
          });
          row[$._R] = item
            .getChild('statistics')
            .getChild('ratings')
            .getChild('ranks')
            .getChildren('rank')
            .findAttribute('name', 'boardgame')
            .getAttribute('value')
            .getValue()
            .toNumber();
          row[$._S] = item
            .getChild('statistics')
            .getChild('ratings')
            .getChild('bayesaverage')
            .getAttribute('value')
            .getValue()
            .toNumber();
          row[$._T] = item
            .getChild('statistics')
            .getChild('ratings')
            .getChild('averageweight')
            .getAttribute('value')
            .getValue()
            .toNumber();
          let minplaytime = item
            .getChild('minplaytime')
            .getAttribute('value')
            .getValue()
            .toNumber();
          let maxplaytime = item
            .getChild('maxplaytime')
            .getAttribute('value')
            .getValue()
            .toNumber();
          row[$._U] =
            minplaytime === maxplaytime
              ? minplaytime
              : `${minplaytime}-${maxplaytime}`;
          row[$._V] = item
            .getChild('yearpublished')
            .getAttribute('value')
            .getValue()
            .toNumber();
          row[$._Z] = current;
          return row;
        } catch (e: unknown) {
          const rowIdentifier = row[$._A].getText() || `row ${row[$.__]}`;
          const errorMessage = e instanceof Error ? e.message : String(e);
          Logger.log(
            `Error processing ${rowIdentifier} (URL: ${url}): ${errorMessage}`
          );
          errors.push(
            `Error processing ${rowIdentifier} (URL: ${url}): ${errorMessage}`
          );
          return row;
        }
      })
      .sort((a: any[], b: any[]) => {
        return a[$.__] < b[$.__] ? -1 : a[$.__] > b[$.__] ? 1 : 0;
      })
      .map((row: any[]) => row.slice($._B));
    sheet.getRange(2, $._B, rows.length, rows[0].length).setValues(rows);
  } catch (e: unknown) {
    const errorMessage = e instanceof Error ? e.message : String(e);
    Logger.log(`Failed after processing ${count} rows: ${errorMessage}`);
    // Combine outer error with collected row-specific errors if any exist
    if (errors.length > 0) {
      errors.push(`Failed after processing ${count} rows: ${errorMessage}`);
      throw new Error(errors.join('\n'));
    } else {
      throw e;
    }
  }
  if (errors.length > 0) {
    throw new Error(errors.join('\n'));
  }
}

function updateArenaRankings() {
  let rankings =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Arena Rankings');
  if (rankings === null) {
    return;
  }
  let html = fetch('https://ja.boardgamearena.com').getContentText();
  let tagMatches =
    (html.match(/"game_tags":([\s\S]*),\n?\s*"top_tags"/m) || [])[1].match(
      /\{"id":[\s\S]*?\}/gm
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
      /\{"id":[\s\S]*?"watched":[\s\S]*?\}/gm
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
  rankings
    .getRange(2, 1, rankings.getLastRow() - 1, games[0].length)
    .clearContent();
  rankings.getRange(2, 1, games.length, games[0].length).setValues(games);
}

function updateArenaTitles() {
  let rankings =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Arena Rankings');
  if (rankings === null) {
    return;
  }
  let titles =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Arena Titles');
  if (titles === null) {
    return;
  }
  let count = 0;
  const startTime = Date.now();
  const timeLimitMs = 5 * 60 * 1000;
  let rows: any[][] = titles
    .getRange('$A$2:$C')
    .getValues()
    .filter((row: any[]) => row[$._A - 1]);
  rows = rows.concat(
    rankings
      .getRange('$A$2:$A')
      .getValues()
      .filter((ranking: any[]) => {
        return !rows
          .map((row: any[]) => row[$._A - 1])
          .includes(ranking[$._A - 1]);
      })
      .map((ranking: any[]) => {
        return [ranking[0], '', ''];
      })
  );
  rows = rows
    .map((row: any[], index: number) => {
      row.unshift(index);
      return row;
    })
    .map((row: any[]) => {
      let url = row[$._A];
      let title = row[$._B];
      if (title) {
        return row;
      }
      // Stop processing if the execution window has exceeded five minutes
      const elapsedMs = Date.now() - startTime;
      if (elapsedMs > timeLimitMs) {
        const identifier = row[$._A] || `index ${row[$.__]}`;
        const message = `updateArenaTitles stopped after ${elapsedMs}ms (limit ${timeLimitMs}ms) at ${identifier}`;
        Logger.log(message);
        return row;
      }
      try {
        title = (fetch(url)
          .getContentText()
          .match(
            /id="game_name" class="block gamename"\n\s*>(.*?)(\(.*?\))?<\/a/m
          ) || [])[1];
      } catch (e: unknown) {
        if (e instanceof Error) {
          Logger.log(`Error: ${e.message}\n${url}`);
        } else {
          Logger.log(`Unknown error: ${String(e)}\n${url}`);
        }
        return row;
      } finally {
        Utilities.sleep(1000);
        count++;
      }
      row[$._B] = title;
      return row;
    })
    .map((row: any[]) => {
      let title = row[$._B].toString();
      if (typeof title.replace === 'function') {
        title = title.replace(/-.*-/g, '');
        title = title.replace(/&amp;/g, '＆');
        title = title.replace(/!/g, '！');
        title = title.replace(/ - /g, ' － ');
        title = title.replace(/《?新版》?/g, '');
        title = title.replace(/\s*･\s*/g, '・');
        title = title.replace(/\s*:\s*/g, '：');
        title = title.replace(/^\s+|\s+$/g, '');
        title = title.replace(
          /^テラフォーミング・マーズ$/,
          'テラフォーミングマーズ'
        );
        title = title.replace(/^チケット・トゥ・ライド/, 'チケットトゥライド');
        title = title.replace(/^ブルゴーニュの城$/, 'ブルゴーニュ');
        title = title.replace(/^サイズ$/, 'サイズ -大鎌戦役-');
        title = title.replace(
          /^ザ・クルー 深海に眠る遺跡$/,
          'ザ・クルー：深海に眠る遺跡'
        );
        title = title.replace(/^パンデミック$/, 'パンデミック：新たなる試練');
        title = title.replace(
          /^ドラフト＆ライトレコーズ$/,
          'ドラフト・アンド・ライト・レコード'
        );
        title = title.replace(/^ラッキーナンバー$/, 'ラッキー・ナンバー');
        title = title.replace(
          /^ガイアプロジェクト$/,
          'テラミスティカ：ガイアプロジェクト'
        );
        title = title.replace(
          /^タペストリー ～文明の錦の御旗～$/,
          'タペストリー'
        );
        title = title.replace(/^メモワール44$/, "メモワール'44");
        title = title.replace(
          /^レイルロード・インク$/,
          'レイルロード・インク：ディープブルー・エディション'
        );
        title = title.replace(/^キャプテン・フリップ$/, 'キャプテンフリップ');
        title = title.replace(/^リビング・フォレスト$/, 'リビングフォレスト');
        title = title.replace(/^アルハンブラ$/, 'アルハンブラの宮殿');
        title = title.replace(/^バニーキングダム$/, 'バニー・キングダム');
        title = title.replace(
          /^アイル・オブ・キャッツ ～ネコたちの楽園～$/,
          'アイル・オブ・キャッツ'
        );
        title = title.replace(
          /^センチュリー：スパイスロード$/,
          'センチュリー；ゴーレム'
        );
      }
      row[$._C] = title;
      return row;
    })
    .sort((a: any[], b: any[]) => {
      return a[$.__] < b[$.__] ? -1 : a[$.__] > b[$.__] ? 1 : 0;
    })
    .map((row: any[]) => row.slice($._A));
  titles.getRange(2, $._A, rows.length, rows[0].length).setValues(rows);
}

function updateRatings() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ratings');
  if (sheet === null) {
    return;
  }
  let base = 'https://bodoge.hoobby.net/friends/16159/boardgames/played?page=';
  let page = 1;
  let ratings = [];
  while (true) {
    let html = fetch(base + page.toString()).getContentText();
    let matches =
      html.match(
        new RegExp('<a class="list--interests-item-title".*?</a>', 'g')
      ) || [];
    if (matches.length === 0) {
      break;
    }
    for (let index = 0; index < matches.length; index++) {
      let title = ((matches[index] || '').match(
        '<div class="list--interests-item-title-japanese">(.*?)</div>'
      ) || [])[1]
        .split('/')[0]
        .replace(new RegExp('（.*）'), '')
        .replace('：新版', '')
        .replace('（拡張）', '')
        .replace('&amp;', '＆')
        .trim();
      let rating = ((matches[index] || '').match(
        '<div class="rating--result-stars" data-rating-mode="result" data-rating-result="(.*?)">'
      ) || [])[1];
      switch (title) {
        case '#hashtag':
          ratings.push(['ハッシュタグ', rating]);
          break;
        case 'ドミニオン：基本カードセット':
          break;
        case 'ドミニオン：錬金術＆収穫祭':
          ratings.push(['ドミニオン：錬金術', rating]);
          ratings.push(['ドミニオン：収穫祭', rating]);
          break;
        case 'ハートオブクラウン：セカンドエディション':
          ratings.push(['ハートオブクラウン', rating]);
          break;
        case 'ヒューゴ オバケと鬼ごっこ':
          ratings.push(['ヒューゴ：オバケと鬼ごっこ', rating]);
          break;
        case 'ダンス・オブ・アイベックス':
          ratings.push(['ヤギたちのダンス', rating]);
          break;
        default:
          ratings.push([title, rating]);
          break;
      }
    }
    Utilities.sleep(1000);
    page++;
  }
  ratings.sort((a, b) => (a[0] > b[0] ? 1 : a[0] < b[0] ? -1 : 0));
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  sheet.getRange(2, 1, ratings.length, 2).setValues(ratings);
}

interface Array<T> {
  findAttribute(name: string, value: string): any;
  sortAttribute(name: string): any[];
}
interface Date {
  addDays(days: number): Date;
}
interface String {
  toNumber(): number | 'N/A';
}

Array.prototype.findAttribute = function (name: string, value: string): any {
  return this.find((element: any) => {
    return element.getAttribute(name).getValue() === value;
  });
};

Array.prototype.sortAttribute = function (name: string): any[] {
  return this.sort((a: any, b: any) => {
    return (
      Number.parseInt(b.getAttribute(name).getValue()) -
      Number.parseInt(a.getAttribute(name).getValue())
    );
  });
};

Date.prototype.addDays = function (days: number): Date {
  let date = new Date(this.getTime());
  date.setDate(date.getDate() + days);
  return date;
};

String.prototype.toNumber = function (): number | 'N/A' {
  let number = Number.parseFloat(this as string);
  return Number.isNaN(number) ? 'N/A' : number;
};
