function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Actions", [
    { name: "Update", functionName: "update" },
  ]);
}

function update() {
  let sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getRange("$A$2:$A").getRichTextValues();
  rows = rows.slice(
    0,
    rows.findIndex((row) => row[0].getText().length === 0)
  );
  let none = [
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
  ];
  let values = rows.map((row) => {
    let url = row[0].getLinkUrl();
    if (url === null) {
      return none;
    }
    Logger.log(url);
    let type = url.split("/")[3];
    let id = url.split("/")[4];
    let response = UrlFetchApp.fetch(
      `https://boardgamegeek.com/xmlapi2/thing?type=${type}&stats=1&id=${id}`
    );
    Utilities.sleep(2000);
    if (response.getResponseCode() !== 200) {
      return none;
    }
    try {
      let xml = XmlService.parse(response.getContentText());
      let item = xml.getRootElement().getChild("item");
      let numbers = item
        .getChildren("poll")
        .find(
          (child) =>
            child.getAttribute("name").getValue() === "suggested_numplayers"
        )
        .getChildren("results")
        .reduce((acc, results) => {
          acc[results.getAttribute("numplayers").getValue()] = results
            .getChildren("result")
            .sort(
              (a, b) =>
                b.getAttribute("numvotes").getValue() -
                a.getAttribute("numvotes").getValue()
            )[0]
            .getAttribute("value")
            .getValue();
          return acc;
        }, {});
      let ratings = item.getChild("statistics").getChild("ratings");
      let minplaytime = Number.parseInt(
        item.getChild("minplaytime").getAttribute("value").getValue()
      );
      let maxplaytime = Number.parseInt(
        item.getChild("maxplaytime").getAttribute("value").getValue()
      );
      let playtime =
        minplaytime === maxplaytime
          ? minplaytime
          : `${minplaytime}-${maxplaytime}`;
      let yearpublished = Number.parseInt(
        item.getChild("yearpublished").getAttribute("value").getValue()
      );
      return [
        ...["2", "3", "4", "5", "6", "7", "8", "9", "10"].map(
          (number) => numbers[number]
        ),
        ...[
          ratings
            .getChild("ranks")
            .getChildren("rank")
            .find(
              (child) => child.getAttribute("name").getValue() === "boardgame"
            )
            .getAttribute("value")
            .getValue(),
          ratings.getChild("bayesaverage").getAttribute("value").getValue(),
          ratings.getChild("averageweight").getAttribute("value").getValue(),
        ]
          .map((value) => Number.parseFloat(value))
          .map((value) => (Number.isNaN(value) ? "N/A" : value)),
        playtime,
        yearpublished,
      ];
    } catch (e) {
      Logger.log(e);
      return none;
    }
  });
  sheet.getRange(2, 7, values.length, none.length).setValues(values);
}
