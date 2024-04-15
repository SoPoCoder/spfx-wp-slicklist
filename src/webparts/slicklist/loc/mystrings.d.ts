declare interface ISlicklistWebPartStrings {
  PropertyPaneDescription: string;
  Table1GroupName: string;
  Table1TitleFieldLabel:string;
  Table1SiteNameFieldLabel: string;
  Table1ListNameFieldLabel: string;
  Table1VisibleColsMobileFieldLabel: string;
  Table1VisibleColsTabletFieldLabel: string;
  Table1VisibleColsDesktopFieldLabel: string;
  LookupColumnFieldLabel:string;
  Table2GroupName: string;
  Table2TitleFieldLabel:string;
  Table2SiteNameFieldLabel: string;
  Table2ListNameFieldLabel: string;
  Table2VisibleColsMobileFieldLabel: string;
  Table2VisibleColsTabletFieldLabel: string;
  Table2VisibleColsDesktopFieldLabel: string;
}

declare module 'SlicklistWebPartStrings' {
  const strings: ISlicklistWebPartStrings;
  export = strings;
}
