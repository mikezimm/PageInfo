declare interface IFpsPageInfoWebPartStrings {
  bannerTitle: string;
  PropertyPaneDescription: string;
  TOCGroupName: string;
  PIStyleGroupName: string;
  PinMeGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;

  // 1 - Analytics options
  analyticsWeb: string;
  analyticsList: string;

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
