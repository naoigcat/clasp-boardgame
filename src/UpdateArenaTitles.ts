class UpdateArenaTitles {
  private static readonly HANDLER = 'updateArenaTitles';
  private static readonly BATCH_SIZE = 100;
  private static readonly TRIGGER_INTERVAL_MINUTES = 5;

  private static countPendingRows(rows: any[][]): number {
    return rows.filter((row: any[]) => row[$._A] && !row[$._C]).length;
  }

  private static normalizeTitle(title: string): string {
    if (typeof title.replace !== 'function') {
      return title;
    }
    title = title.replace(/-.*-/g, '');
    title = title.replace(/&amp;/g, '＆');
    title = title.replace(/!/g, '！');
    title = title.replace(/ - /g, ' － ');
    title = title.replace(/《?新版》?/g, '');
    title = title.replace(/第\d+版/g, '');
    title = title.replace(/\s*･\s*/g, '・');
    title = title.replace(/\s*:\s*/g, '：');
    title = title.replace(/^\s+|\s+$/g, '');
    title = title.replace(
      /^テラフォーミング・マーズ$/,
      'テラフォーミングマーズ',
    );
    title = title.replace(/^チケット・トゥ・ライド/, 'チケットトゥライド');
    title = title.replace(/^ブルゴーニュの城$/, 'ブルゴーニュ');
    title = title.replace(/^サイズ$/, 'サイズ -大鎌戦役-');
    title = title.replace(
      /^ザ・クルー 深海に眠る遺跡$/,
      'ザ・クルー：深海に眠る遺跡',
    );
    title = title.replace(/^パンデミック$/, 'パンデミック：新たなる試練');
    title = title.replace(
      /^ドラフト＆ライトレコーズ$/,
      'ドラフト・アンド・ライト・レコード',
    );
    title = title.replace(/^ラッキーナンバー$/, 'ラッキー・ナンバー');
    title = title.replace(
      /^ガイアプロジェクト$/,
      'テラミスティカ：ガイアプロジェクト',
    );
    title = title.replace(/^タペストリー ～文明の錦の御旗～$/, 'タペストリー');
    title = title.replace(/^メモワール44$/, "メモワール'44");
    title = title.replace(
      /^レイルロード・インク$/,
      'レイルロード・インク：ディープブルー・エディション',
    );
    title = title.replace(/^キャプテン・フリップ$/, 'キャプテンフリップ');
    title = title.replace(/^リビング・フォレスト$/, 'リビングフォレスト');
    title = title.replace(/^アルハンブラ$/, 'アルハンブラの宮殿');
    title = title.replace(/^バニーキングダム$/, 'バニー・キングダム');
    title = title.replace(
      /^アイル・オブ・キャッツ ～ネコたちの楽園～$/,
      'アイル・オブ・キャッツ',
    );
    title = title.replace(
      /^センチュリー：スパイスロード$/,
      'センチュリー；ゴーレム',
    );
    return title;
  }

  static run(): void {
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
        }),
    );
    rows = rows.map((row: any[], index: number) => {
      row.unshift(index);
      return row;
    });
    const pendingCount = UpdateArenaTitles.countPendingRows(rows);
    if (pendingCount === 0) {
      Triggers.deleteAll(UpdateArenaTitles.HANDLER);
      return;
    }
    let processedCount = 0;
    let count = 0;
    rows = rows
      .map((row: any[]) => {
        if (row[$._C]) {
          return row;
        }
        if (processedCount >= UpdateArenaTitles.BATCH_SIZE) {
          return row;
        }
        processedCount++;
        let url = row[$._A];
        if (!row[$._B]) {
          try {
            row[$._B] = (fetch(url)
              .getContentText()
              .match(
                /id="game_name" class="block gamename"\n\s*>(.*?)(\(.*?\))?<\/a/m,
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
        }
        row[$._C] = UpdateArenaTitles.normalizeTitle(row[$._B].toString());
        return row;
      })
      .sort((a: any[], b: any[]) => {
        return a[$.__] < b[$.__] ? -1 : a[$.__] > b[$.__] ? 1 : 0;
      });
    const remainingCount = UpdateArenaTitles.countPendingRows(rows);
    rows = rows.map((row: any[]) => row.slice($._A));
    if (rows.length === 0) {
      return;
    }
    titles.getRange(2, $._A, rows.length, rows[0].length).setValues(rows);
    if (remainingCount > 0) {
      Triggers.ensure(
        UpdateArenaTitles.HANDLER,
        UpdateArenaTitles.TRIGGER_INTERVAL_MINUTES,
      );
    } else {
      Triggers.deleteAll(UpdateArenaTitles.HANDLER);
    }
  }
}

function updateArenaTitles(): void {
  UpdateArenaTitles.run();
}
