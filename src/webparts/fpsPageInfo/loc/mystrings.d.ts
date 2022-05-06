declare interface IFpsPageInfoWebPartStrings {
  PropertyPaneDescription: string;
  TOCGroupName: string;
  PinMeGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;

  //Copied from AdvancedPagePropertiesWebPart.ts
  PropertyPaneDescriptionP: string;
  OOTBPropGroupName: string;
  PropertiesGroupName: string;
  PropsTitleFieldLabel: string;
  SelectedPropertiesFieldLabel: string;
  PropPaneAddButtonText: string;
  PropPaneDeleteButtonText: string;
  LogAppName: string;
}

declare module 'FpsPageInfoWebPartStrings' {
  const strings: IFpsPageInfoWebPartStrings;
  export = strings;
}
