declare interface ISiteAdminWebPartStrings {
  AzureFunctionUrlFieldDescription: string;
  AzureFunctionUrlFieldLabel: string;
  BasicGroupName: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'SiteAdminWebPartStrings' {
  const strings: ISiteAdminWebPartStrings;
  export = strings;
}
