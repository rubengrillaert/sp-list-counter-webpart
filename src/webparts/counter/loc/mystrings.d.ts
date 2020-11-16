declare interface ICounterWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  SelectSiteContentList: string;
  FilterSiteContentList: string;
  SelectList: string;
  Title:string;
}

declare module 'CounterWebPartStrings' {
  const strings: ICounterWebPartStrings;
  export = strings;
}
