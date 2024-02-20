function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Functions")
    .addItem("Update", "update")
    .addToUi();
}

function update() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Games");
  if (sheet === null) {
    return;
  }
  let rows: any[][] = sheet.getRange("$A$2:$A").getRichTextValues();
  sheet
    .getRange("$B$2:$V")
    .getValues()
    .forEach((row, index) => {
      rows[index] = rows[index].concat(row);
    });
  let last = rows.findIndex((row) => row[0].getText().length === 0);
  let current = new Date();
  let count = 0;
  rows = rows
    .slice(0, last)
    .map((row) => {
      // Clear columns containing values by ARRAYFORMULA
      row[2] = "";
      row[5] = "";
      // Reduces the number of API executions because there is a 6 minute timeout
      if (count > 100) {
        return row;
      }
      Logger.log(row[0].getText());
      let url = row[0].getLinkUrl();
      if (url === null) {
        return row;
      }
      let updated = row[21] as Date;
      // Skip if you have been running the API within the past week
      if (updated && updated.withDate(updated.getDate() + 7) > current) {
        return row;
      }
      try {
        let type = url.split("/")[3];
        let id = url.split("/")[4];
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
          .getChild("item");
        let numbers = item
          .getChildren("poll")
          .findAttribute("name", "suggested_numplayers")
          .getChildren("results")
          .reduce((acc, results) => {
            acc[results.getAttribute("numplayers").getValue()] = results
              .getChildren("result")
              .sortAttribute("numvotes")[0]
              .getAttribute("value")
              .getValue();
            return acc;
          }, {});
        let indexes = [...Array(10)].map((v, i) => i + 7);
        indexes.forEach((index) => {
          row[index] = numbers[(index - 5).toString()];
        });
        row[16] = item
          .getChild("statistics")
          .getChild("ratings")
          .getChild("ranks")
          .getChildren("rank")
          .findAttribute("name", "boardgame")
          .getAttribute("value")
          .getValue()
          .toNumber();
        row[17] = item
          .getChild("statistics")
          .getChild("ratings")
          .getChild("bayesaverage")
          .getAttribute("value")
          .getValue()
          .toNumber();
        row[18] = item
          .getChild("statistics")
          .getChild("ratings")
          .getChild("averageweight")
          .getAttribute("value")
          .getValue()
          .toNumber();
        let minplaytime = item
          .getChild("minplaytime")
          .getAttribute("value")
          .getValue()
          .toNumber();
        let maxplaytime = item
          .getChild("maxplaytime")
          .getAttribute("value")
          .getValue()
          .toNumber();
        row[19] =
          minplaytime === maxplaytime
            ? minplaytime
            : `${minplaytime}-${maxplaytime}`;
        row[20] = item
          .getChild("yearpublished")
          .getAttribute("value")
          .getValue()
          .toNumber();
        row[21] = current;
        return row;
      } catch (e) {
        Logger.log(e);
        return row;
      }
    })
    .map((row) => row.slice(1));
  sheet.getRange(2, 2, rows.length, rows[0].length).setValues(rows);
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

Array.prototype.findAttribute = function <T extends GoogleAppsScript.XML_Service.Element>(name: string, value: string): T {
  return this.find((element: GoogleAppsScript.XML_Service.Element) => {
    return element.getAttribute(name).getValue() === value;
  });
};

Array.prototype.sortAttribute = function <T extends GoogleAppsScript.XML_Service.Element>(name: string): T[] {
  return this.sort((a: GoogleAppsScript.XML_Service.Element, b: GoogleAppsScript.XML_Service.Element) => {
    return Number.parseInt(b.getAttribute(name).getValue()) - Number.parseInt(a.getAttribute(name).getValue());
  });
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
