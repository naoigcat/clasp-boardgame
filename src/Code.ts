function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Functions')
    .addItem('Update Games', 'updateGames')
    .addItem('Update Ratings', 'updateRatings')
    .addToUi();
}

function updateGames() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Games');
  if (sheet === null) {
    return;
  }
  let rows: any[][] = sheet.getRange('$A$2:$A').getRichTextValues();
  sheet
    .getRange('$B$2:$Z')
    .getValues()
    .forEach((row: any[], index: number) => {
      rows[index] = rows[index].concat(row);
    });
  let last = rows.findIndex((row: any[]) => row[0].getText().length === 0);
  let current = new Date();
  let count = 0;
  rows = rows
    .slice(0, last)
    .map((row: any[], index: number) => {
      row.unshift(index);
      return row;
    })
    .sort((a: any[], b: any[]) => {
      return a[25] < b[25] ? -1 : a[25] > b[25] ? 1 : 0;
    })
    .map((row: any[]) => {
      // Clear columns containing values by ARRAYFORMULA
      [3, 6, 23, 24].forEach((index) => {
        row[index] = '';
      });
      // Reduces the number of API executions because there is a 6 minute timeout
      if (count > 100) {
        return row;
      }
      Logger.log(row[1].getText());
      let url = row[1].getLinkUrl();
      if (url === null) {
        return row;
      }
      let updated = row[25] as Date;
      // Skip if you have been running the API within the past week
      if (updated && updated.withDate(updated.getDate() + 7) > current) {
        return row;
      }
      try {
        let type = url.split('/')[3];
        let id = url.split('/')[4];
        let endpoint = `https://boardgamegeek.com/xmlapi2/thing?type=${type}&stats=1&id=${id}`;
        Logger.log(endpoint);
        let response = UrlFetchApp.fetch(endpoint);
        Utilities.sleep(2000);
        count++;
        if (response.getResponseCode() !== 200) {
          return row;
        }
        let item = XmlService.parse(response.getContentText())
          .getRootElement()
          .getChild('item');
        let numbers = item
          .getChildren('poll')
          .findAttribute('name', 'suggested_numplayers')
          .getChildren('results')
          .reduce((acc, results) => {
            acc[results.getAttribute('numplayers').getValue()] = results
              .getChildren('result')
              .sortAttribute('numvotes')[0]
              .getAttribute('value')
              .getValue();
            return acc;
          }, {});
        Logger.log(numbers);
        let indexes = [...Array(10)].map((v, i) => i + 9);
        indexes.forEach((index) => {
          row[index] = numbers[(index - 7).toString()];
        });
        row[18] = item
          .getChild('statistics')
          .getChild('ratings')
          .getChild('ranks')
          .getChildren('rank')
          .findAttribute('name', 'boardgame')
          .getAttribute('value')
          .getValue()
          .toNumber();
        row[19] = item
          .getChild('statistics')
          .getChild('ratings')
          .getChild('bayesaverage')
          .getAttribute('value')
          .getValue()
          .toNumber();
        row[20] = item
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
        row[21] =
          minplaytime === maxplaytime
            ? minplaytime
            : `${minplaytime}-${maxplaytime}`;
        row[22] = item
          .getChild('yearpublished')
          .getAttribute('value')
          .getValue()
          .toNumber();
        row[25] = current;
        return row;
      } catch (e) {
        Logger.log(e);
        return row;
      }
    })
    .sort((a: any[], b: any[]) => {
      return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0;
    })
    .map((row: any[]) => row.slice(2));
  sheet.getRange(2, 2, rows.length, rows[0].length).setValues(rows);
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
    let html = UrlFetchApp.fetch(base + page.toString()).getContentText();
    let matches = html.match(
      new RegExp('<a class="list--interests-item-title".*?</a>', 'g')
    );
    if (!matches || matches.length === 0) {
      break;
    }
    for (let index = 0; index < matches.length; index++) {
      let title = matches[index]
        .match(
          '<div class="list--interests-item-title-japanese">(.*?)</div>'
        )[1]
        .split('/')[0]
        .replace(new RegExp('（.*）'), '')
        .replace('：新版', '')
        .replace('&amp;', '＆')
        .trim();
      let rating = matches[index].match(
        '<div class="rating--result-stars" data-rating-mode="result" data-rating-result="(.*?)">'
      )[1];
      switch (title) {
        case 'ドミニオン：錬金術＆収穫祭':
          ratings.push(['ドミニオン：錬金術', rating]);
          ratings.push(['ドミニオン：収穫祭', rating]);
        default:
          ratings.push([title, rating]);
      }
    }
    Utilities.sleep(1000);
    page++;
  }
  ratings.sort((a, b) => (a[0] > b[0] ? 1 : a[0] < b[0] ? -1 : 0));
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  sheet.getRange(2, 1, ratings.length, 2).setValues(ratings);
}

export {};

declare global {
  interface Array<T> {
    findAttribute(name: string, value: string): T;
    sortAttribute(name: string): T[];
  }
  interface Date {
    withDate(dayValue: number): Date;
  }
  interface String {
    toNumber(): number | "N/A";
  }
}

Array.prototype.findAttribute = function <
  T extends GoogleAppsScript.XML_Service.Element
>(name: string, value: string): T {
  return this.find((element: GoogleAppsScript.XML_Service.Element) => {
    return element.getAttribute(name).getValue() === value;
  });
};

Array.prototype.sortAttribute = function <
  T extends GoogleAppsScript.XML_Service.Element
>(name: string): T[] {
  return this.sort(
    (
      a: GoogleAppsScript.XML_Service.Element,
      b: GoogleAppsScript.XML_Service.Element
    ) => {
      return (
        Number.parseInt(b.getAttribute(name).getValue()) -
        Number.parseInt(a.getAttribute(name).getValue())
      );
    }
  );
};

Date.prototype.withDate = function (dayValue: number): Date {
  let date = new Date(this.getTime());
  date.setDate(dayValue);
  return date;
};

String.prototype.toNumber = function (): number | "N/A" {
  let number = Number.parseFloat(this);
  return Number.isNaN(number) ? "N/A" : number;
};
