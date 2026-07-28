/**
 * Exact title substitutions applied after generic title normalization.
 *
 * Board Game Arena and Bodoge often use marketing or edition-specific spellings
 * that differ from the canonical names already used in the spreadsheet. These
 * deterministic aliases keep matching stable without requiring sheet edits each
 * time a source renames a title.
 */
const TITLE_NORMALIZATION_ALIASES: Readonly<Record<string, string>> = {
  'テラフォーミング・マーズ': 'テラフォーミングマーズ',
  'ブルゴーニュの城': 'ブルゴーニュ',
  'サイズ': 'サイズ -大鎌戦役-',
  'ザ・クルー 深海に眠る遺跡': 'ザ・クルー：深海に眠る遺跡',
  'パンデミック': 'パンデミック：新たなる試練',
  'ドラフト＆ライトレコーズ': 'ドラフト・アンド・ライト・レコード',
  'ラッキーナンバー': 'ラッキー・ナンバー',
  'ガイアプロジェクト': 'テラミスティカ：ガイアプロジェクト',
  'タペストリー ～文明の錦の御旗～': 'タペストリー',
  'メモワール44': "メモワール'44",
  'レイルロード・インク':
    'レイルロード・インク：ディープブルー・エディション',
  'キャプテン・フリップ': 'キャプテンフリップ',
  'リビング・フォレスト': 'リビングフォレスト',
  'アルハンブラ': 'アルハンブラの宮殿',
  'バニーキングダム': 'バニー・キングダム',
  'アイル・オブ・キャッツ ～ネコたちの楽園～': 'アイル・オブ・キャッツ',
} as const;

/**
 * Prefix substitutions that apply to a family of related titles.
 *
 * Series entries share a prefix but keep their distinctive suffixes, so a
 * prefix rewrite avoids enumerating every expansion as an exact alias.
 */
const TITLE_NORMALIZATION_PREFIX_ALIASES: Readonly<Record<string, string>> = {
  'チケット・トゥ・ライド': 'チケットトゥライド',
} as const;

/**
 * Bodoge title substitutions. One source title can intentionally become more
 * than one rating row when it represents a bundled product that the spreadsheet
 * tracks as separate games.
 */
const RATING_TITLE_ALIASES: Readonly<Record<string, readonly string[]>> = {
  '#hashtag': ['ハッシュタグ'],
  'ドミニオン：錬金術＆収穫祭': ['ドミニオン：錬金術', 'ドミニオン：収穫祭'],
  'ハートオブクラウン：セカンドエディション': ['ハートオブクラウン'],
  'ヒューゴ オバケと鬼ごっこ': ['ヒューゴ：オバケと鬼ごっこ'],
  'ダンス・オブ・アイベックス': ['ヤギたちのダンス'],
} as const;

/**
 * Bodoge titles intentionally omitted because their rating belongs to another
 * row or does not map to a standalone game in the spreadsheet.
 */
const EXCLUDED_RATING_TITLES: readonly string[] = [
  'ドミニオン：基本カードセット',
];

/**
 * Game-specific player-count recommendations that correct known upstream data.
 *
 * BoardGameGeek polls are community-voted and occasionally disagree with how a
 * game is actually played. Overrides are keyed by BoardGameGeek ID so they
 * survive title renames.
 */
const GAME_PLAYER_COUNT_OVERRIDES: Readonly<
  Record<string, Readonly<Record<string, string>>>
> = {
  '8172': {
    '7': 'Recommended',
    '8': 'Recommended',
    '9': 'Recommended',
    '10': 'Recommended',
  },
};
