class UpdateRatings {
  static run(): void {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ratings');
    if (sheet === null) {
      return;
    }
    const bodogeUserId = getBodogeUserId();
    if (!bodogeUserId) {
      return;
    }
    let base = `https://bodoge.hoobby.net/friends/${bodogeUserId}/boardgames/played?page=`;
    let page = 1;
    let ratings: any[][] = [];
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
        let titleMatch = (matches[index] || '').match(
          '<div class="list--interests-item-title-japanese">(.*?)</div>'
        );
        if (!titleMatch?.[1]) {
          continue;
        }
        let title = titleMatch[1]
          .split('/')[0]
          .replace(new RegExp('（.*）'), '')
          .replace('：新版', '')
          .replace('（拡張）', '')
          .replace('&amp;', '＆')
          .trim();
        let ratingMatch = (matches[index] || '').match(
          '<div class="rating--result-stars" data-rating-mode="result" data-rating-result="(.*?)">'
        );
        let rating = ratingMatch?.[1];
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
}

function updateRatings(): void {
  UpdateRatings.run();
}
