declare interface ISlicklistWebPartStrings {
  PropertyPaneDescription: string;
  Table1GroupName: string;
  Table1TitleFieldLabel:string;
  Table1SiteNameFieldLabel: string;
  Table1ListNameFieldLabel: string;
  Table1VisibleColsMobileFieldLabel: string;
  Table1VisibleColsTabletFieldLabel: string;
  Table1VisibleColsDesktopFieldLabel: string;
}

declare module 'SlicklistWebPartStrings' {
  const strings: ISlicklistWebPartStrings;
  export = strings;
}
