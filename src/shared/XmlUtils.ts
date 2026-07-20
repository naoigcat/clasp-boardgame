/**
 * XML element type supplied by Apps Script's XmlService.
 */
type XmlElement = GoogleAppsScript.XML_Service.Element;

/**
 * Numeric value displayed when BoardGameGeek omits a usable number.
 */
type NumericDisplayValue = number | 'N/A';

/**
 * Returns a required child element or raises a descriptive parsing error.
 */
function getRequiredXmlChild(parent: XmlElement, childName: string): XmlElement {
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
