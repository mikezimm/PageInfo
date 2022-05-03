declare interface IFpsPageInfoWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;

  //Copied from AdvancedPagePropertiesWebPart.ts
  PropertyPaneDescriptionP: string;
  SelectionGroupName: string;
  TitleFieldLabel: string;
  SelectedPropertiesFieldLabel: string;
  PropPaneAddButtonText: string;
  PropPaneDeleteButtonText: string;
  LogAppName: string;
}

declare module 'FpsPageInfoWebPartStrings' {
  const strings: IFpsPageInfoWebPartStrings;
  export = strings;
}
