function findAttribute(
  elements: any[],
  name: string,
  value: string,
): any {
  const element = elements.find((element: any) => {
    return element.getAttribute(name).getValue() === value;
  });
  if (element === undefined) {
    throw new Error(`Element with ${name}="${value}" not found`);
  }
  return element;
}

function sortAttribute(elements: any[], name: string): any[] {
  return elements.slice().sort((a: any, b: any) => {
    return (
      Number.parseInt(b.getAttribute(name).getValue()) -
      Number.parseInt(a.getAttribute(name).getValue())
    );
  });
}

function addDays(date: Date, days: number): Date {
  let result = new Date(date.getTime());
  result.setDate(result.getDate() + days);
  return result;
}

function toNumber(value: string): number | 'N/A' {
  let number = Number.parseFloat(value);
  return Number.isNaN(number) ? 'N/A' : number;
}
