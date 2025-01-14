declare interface ISiteAdminWebPartStrings {
  AppTitle: string;
  SitePropCommentsOnSitePagesDisabled: string;
  SitePropContainsAppCatalog: string;
  SitePropCreated: string;
  SitePropCustomScriptsEnabled: string;
  SitePropDisableCompanyWideSharingLinks: string;
  SitePropHubSite: string;
  SitePropHubSiteConnected: string;
  SitePropIncreaseStorage: string;
  SitePropLockState: string;
  SitePropShareByEmailEnabled: string;
  SitePropSocialBarOnSitePagesDisabled: string;
  SitePropStorageUsed: string;
  SitePropTemplate: string;
  SitePropTitle: string;
  WebPropCommentsOnSitePagesDisabled: string;
  WebPropSearchPropertyKey: string;
  WebPropSearchPropertyDescription: string;
  WebPropSearchPropertyLabel: string;
  WebPropSearchPropertyTabName: string;
  WebPropSearchPropertyValues: string;
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
