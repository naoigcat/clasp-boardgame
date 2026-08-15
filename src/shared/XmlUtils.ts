/**
 * XML types supplied by Apps Script's XmlService.
 */
type XmlDocument = GoogleAppsScript.XML_Service.Document;
type XmlElement = GoogleAppsScript.XML_Service.Element;

/**
 * Numeric value displayed when BoardGameGeek omits a usable number.
 *
 * The API uses the literal string "N/A" for missing ranks and similar fields;
 * preserving that token avoids inventing zeros that would look like real data.
 */
type NumericDisplayValue = number | 'N/A';

/**
 * Returns the document root element or raises a descriptive parsing error.
 *
 * XmlService may yield a null root for empty or malformed documents. Failing
 * here keeps callers aligned with the required-child helpers below.
 */
function getRequiredXmlRootElement(document: XmlDocument): XmlElement {
  const rootElement = document.getRootElement();
  if (rootElement === null) {
    throw new Error('Required XML root element was not found');
  }
  return rootElement;
}

/**
 * Returns a required child element or raises a descriptive parsing error.
 *
 * BoardGameGeek responses are expected to contain a fixed shape. Failing early
 * with the missing child name is easier to diagnose than a later null dereference.
 */
function getRequiredXmlChild(
  parent: XmlElement,
  childName: string,
): XmlElement {
  const child = parent.getChild(childName);
  if (child === null) {
    throw new Error(`Required XML child "${childName}" was not found`);
  }
  return child;
}

/**
 * Returns a required XML attribute value or raises a descriptive parsing error.
 */
function getRequiredXmlAttributeValue(
  element: XmlElement,
  attributeName: string,
): string {
  const attribute = element.getAttribute(attributeName);
  if (attribute === null) {
    throw new Error(`Required XML attribute "${attributeName}" was not found`);
  }
  return attribute.getValue();
}

/**
 * Finds the first element whose named attribute equals the requested value.
 *
 * Used for polls and ranks that appear as sibling elements distinguished only
 * by attributes such as `name="boardgame"`.
 */
function findXmlElementByAttribute(
  elements: readonly XmlElement[],
  attributeName: string,
  attributeValue: string,
): XmlElement {
  const element = elements.find(
    (candidate) =>
      candidate.getAttribute(attributeName)?.getValue() === attributeValue,
  );
  if (element === undefined) {
    throw new Error(
      `Element with ${attributeName}="${attributeValue}" was not found`,
    );
  }
  return element;
}

/**
 * Sorts XML elements by a numeric attribute in descending order.
 *
 * BoardGameGeek player-count polls store vote totals on each result. The
 * highest vote count is treated as the community preference for that count.
 */
function sortXmlElementsByNumericAttribute(
  elements: readonly XmlElement[],
  attributeName: string,
): XmlElement[] {
  return elements.slice().sort((first, second) => {
    return (
      Number.parseInt(getRequiredXmlAttributeValue(second, attributeName), 10) -
      Number.parseInt(getRequiredXmlAttributeValue(first, attributeName), 10)
    );
  });
}

/**
 * Parses a BoardGameGeek number while preserving its explicit "N/A" state.
 */
function parseDisplayNumber(value: string): NumericDisplayValue {
  const number = Number.parseFloat(value);
  return Number.isNaN(number) ? 'N/A' : number;
}
