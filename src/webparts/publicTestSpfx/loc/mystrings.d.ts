declare interface IPublicTestSpfxWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'PublicTestSpfxWebPartStrings' {
  const strings: IPublicTestSpfxWebPartStrings;
  export = strings;
}
