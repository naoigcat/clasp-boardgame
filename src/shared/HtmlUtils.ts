/**
 * Extracts and parses an array-valued JSON property embedded in an HTML page.
 *
 * Board Game Arena embeds JSON inside its page source, so bracket matching is
 * used instead of a greedy regular expression that could consume a later field.
 */
function extractEmbeddedJsonArray(page: string, propertyName: string): unknown[] {
  const propertyToken = `"${propertyName}"`;
  const propertyIndex = page.indexOf(propertyToken);
  if (propertyIndex === -1) {
    throw new Error(`JSON property "${propertyName}" was not found`);
  }

  const valueSeparatorIndex = page.indexOf(
    ':',
    propertyIndex + propertyToken.length,
  );
  const arrayStartIndex = page.indexOf('[', valueSeparatorIndex);
  if (valueSeparatorIndex === -1 || arrayStartIndex === -1) {
    throw new Error(`JSON array "${propertyName}" was not found`);
  }

  const arrayEndIndex = findJsonArrayEnd(page, arrayStartIndex);
  return JSON.parse(page.slice(arrayStartIndex, arrayEndIndex + 1)) as unknown[];
}

/**
 * Finds the inclusive end index of a JSON array while respecting quoted text.
 */
function findJsonArrayEnd(source: string, arrayStartIndex: number): number {
  let depth = 0;
  let isInsideString = false;
  let isEscaped = false;

  for (let index = arrayStartIndex; index < source.length; index += 1) {
    const character = source[index];

    if (isInsideString) {
      if (isEscaped) {
        isEscaped = false;
      } else if (character === '\\') {
        isEscaped = true;
      } else if (character === '"') {
        isInsideString = false;
      }
      continue;
    }

    if (character === '"') {
      isInsideString = true;
    } else if (character === '[') {
      depth += 1;
    } else if (character === ']') {
      depth -= 1;
      if (depth === 0) {
        return index;
      }
    }
  }

  throw new Error('JSON array is not closed');
}
