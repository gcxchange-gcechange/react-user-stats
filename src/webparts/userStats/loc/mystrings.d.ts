declare interface IUserStatsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  StorageCapacityLabel: string;
  StorageUnitLabel;
}

declare module 'UserStatsWebPartStrings' {
  const strings: IUserStatsWebPartStrings;
  export = strings;
}
