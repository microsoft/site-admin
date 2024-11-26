declare interface ISiteAdminWebPartStrings {
  SitePropCommentsOnSitePagesDisabled: string;
  SitePropContainsAppCatalog: string;
  SitePropCustomScriptsEnabled: string;
  SitePropDisableCompanyWideSharingLinks: string;
  SitePropIncreaseStorage: string;
  SitePropLockState: string;
  SitePropShareByEmailEnabled: string;
  SitePropSocialBarOnSitePagesDisabled: string;
  WebPropCommentsOnSitePagesDisabled: string;
  WebPropSearchPropertyKey: string;
  WebPropSearchPropertyDescription: string;
  WebPropSearchPropertyLabel: string;
  WebPropExcludeFromOfflineClient: string;
  WebPropSearchScope: string;
  WebPropWebTemplate: string;
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
