declare interface ISiteAdminWebPartStrings {
  AppTitle: string;
  MaxStorage: string;
  MaxStorageDescription: string;
  ReportsDocRententionYears: string;
  ReportsDocSearchFileExt: string;
  ReportsDocSearchKeywords: string;
  SitePropAttestationDate: string;
  SitePropAttestationUser: string;
  SitePropClientSideAssets: string;
  SitePropCommentsOnSitePagesDisabled: string;
  SitePropContainsAppCatalog: string;
  SitePropCreated: string;
  SitePropCustomScriptsEnabled: string;
  SitePropDisableCompanyWideSharingLinks: string;
  SitePropExcludeFromOfflineClient: string;
  SitePropHubSite: string;
  SitePropHubSiteConnected: string;
  SitePropIncreaseStorage: string;
  SitePropLockState: string;
  SitePropNoCrawl: string;
  SitePropSensitivityLabelId: string;
  SitePropShareByEmailEnabled: string;
  SitePropSocialBarOnSitePagesDisabled: string;
  SitePropStorageUsed: string;
  SitePropTemplate: string;
  SitePropTitle: string;
  WebPropCommentsOnSitePagesDisabled: string;
  WebPropExcludeFromOfflineClient: string;
  WebPropNoCrawl: string;
  WebPropSearchPropertyKey: string;
  WebPropSearchPropertyDescription: string;
  WebPropSearchPropertyLabel: string;
  WebPropSearchPropertyManagedProperty: string;
  WebPropSearchPropertyReportName: string;
  WebPropSearchPropertyTabName: string;
  WebPropSearchPropertyValues: string;
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
