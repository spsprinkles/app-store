declare interface IAppStoreWebPartStrings {
  AppCatalogUrlFieldLabel: string;
  AppCatalogUrlFieldDescription: string;
  TitleFieldDescription: string;
  TitleFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
}

declare module 'AppStoreWebPartStrings' {
  const strings: IAppStoreWebPartStrings;
  export = strings;
}
