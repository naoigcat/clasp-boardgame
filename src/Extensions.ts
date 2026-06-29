interface Array<T> {
  findAttribute(name: string, value: string): any;
  sortAttribute(name: string): any[];
}
interface Date {
  addDays(days: number): Date;
}
interface String {
  toNumber(): number | 'N/A';
}

Array.prototype.findAttribute = function (name: string, value: string): any {
  return this.find((element: any) => {
    return element.getAttribute(name).getValue() === value;
  });
};

Array.prototype.sortAttribute = function (name: string): any[] {
  return this.sort((a: any, b: any) => {
    return (
      Number.parseInt(b.getAttribute(name).getValue()) -
      Number.parseInt(a.getAttribute(name).getValue())
    );
  });
};

Date.prototype.addDays = function (days: number): Date {
  let date = new Date(this.getTime());
  date.setDate(date.getDate() + days);
  return date;
};

String.prototype.toNumber = function (): number | 'N/A' {
  let number = Number.parseFloat(this as string);
  return Number.isNaN(number) ? 'N/A' : number;
};
