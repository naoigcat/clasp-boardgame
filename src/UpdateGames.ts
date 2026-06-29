class UpdateGames {
  private static readonly HANDLER = 'updateGames';
  private static readonly BATCH_SIZE = 50;
  private static readonly TRIGGER_INTERVAL_MINUTES = 5;

  private static readonly GAME_OVERRIDES: {
    [gameId: string]: { [playerCount: string]: string };
  } = {
    '8172': {
      '7': 'Recommended',
      '8': 'Recommended',
      '9': 'Recommended',
      '10': 'Recommended',
    },
  };

  private static countPendingRows(rows: any[][], current: Date): number {
    const emptyIndex = rows.findIndex(
      (row: any[]) => row[$._A].getText().length === 0,
    );
    const end = emptyIndex === -1 ? rows.length : emptyIndex;
    let count = 0;
    for (let index = 0; index < end; index++) {
      const row = rows[index];
      const url = row[$._A].getLinkUrl();
      if (url === null) {
        continue;
      }
      const updated = row[$._Z] as Date;
      if (updated && updated.addDays(7) > current) {
        continue;
      }
      count++;
    }
    return count;
  }

  static run(): void {
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
    const pendingCount = UpdateGames.countPendingRows(rows, current);
    if (pendingCount === 0) {
      Triggers.deleteAll(UpdateGames.HANDLER);
      return;
    }
    let processedCount = 0;
    let count = 0;
    let errors: string[] = [];
    let remainingCount = 0;
    try {
      const emptyIndex = rows.findIndex(
        (row: any[]) => row[$._A].getText().length === 0,
      );
      const end = emptyIndex === -1 ? rows.length : emptyIndex;
      rows = rows
        .slice(0, end)
        .sort((a: any[], b: any[]) => {
          return a[$._Z] < b[$._Z] ? -1 : a[$._Z] > b[$._Z] ? 1 : 0;
        })
        .map((row: any[]) => {
          // Even when no processing occurs, clear the cells so ARRAYFORMULA can continue expanding values
          [$._C, $._F, $._W, $._X, $._Y].forEach((index) => {
            row[index] = '';
          });
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
          if (processedCount >= UpdateGames.BATCH_SIZE) {
            return row;
          }
          processedCount++;
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
            const item = XmlService.parse(body)
              .getRootElement()
              .getChild('item');
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
            const overrides = UpdateGames.GAME_OVERRIDES[id];
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
              `Error processing ${rowIdentifier} (URL: ${url}): ${errorMessage}`,
            );
            errors.push(
              `Error processing ${rowIdentifier} (URL: ${url}): ${errorMessage}`,
            );
            return row;
          }
        })
        .sort((a: any[], b: any[]) => {
          return a[$.__] < b[$.__] ? -1 : a[$.__] > b[$.__] ? 1 : 0;
        });
      remainingCount = UpdateGames.countPendingRows(rows, current);
      rows = rows.map((row: any[]) => row.slice($._B));
      if (rows.length === 0) {
        return;
      }
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
    if (remainingCount > 0) {
      Triggers.ensure(
        UpdateGames.HANDLER,
        UpdateGames.TRIGGER_INTERVAL_MINUTES,
      );
    } else {
      Triggers.deleteAll(UpdateGames.HANDLER);
    }
  }
}

function updateGames(): void {
  UpdateGames.run();
}
