import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration, IPropertyPaneGroup,
  PropertyPaneDropdown, PropertyPaneHorizontalRule,
  PropertyPaneLabel, PropertyPaneLink, PropertyPaneTextField, PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'SiteAdminWebPartStrings';

export interface ISiteAdminWebPartProps {
  AppTitle: string;
  CloudEnvironment: string;
  MaxStorage: number;
  MaxStorageDescription: string;
  ReportsDocSearchFileExt: string;
  ReportsDocSearchKeywords: string;
  SitePropCommentsOnSitePagesDisabled: boolean;
  SitePropCommentsOnSitePagesDisabledDescription: string;
  SitePropCommentsOnSitePagesDisabledLabel: string;
  SitePropContainsAppCatalog: boolean;
  SitePropContainsAppCatalogDescription: string;
  SitePropContainsAppCatalogLabel: string;
  SitePropCreated: boolean;
  SitePropCreatedDescription: string;
  SitePropCreatedLabel: string;
  SitePropCustomScriptsEnabled: boolean;
  SitePropCustomScriptsEnabledDescription: string;
  SitePropCustomScriptsEnabledLabel: string;
  SitePropDisableCompanyWideSharingLinks: boolean;
  SitePropDisableCompanyWideSharingLinksDescription: string;
  SitePropDisableCompanyWideSharingLinksLabel: string;
  SitePropExcludeFromOfflineClient: boolean;
  SitePropExcludeFromOfflineClientDescription: string;
  SitePropExcludeFromOfflineClientLabel: string;
  SitePropHubSite: boolean;
  SitePropHubSiteDescription: string;
  SitePropHubSiteLabel: string;
  SitePropHubSiteConnected: boolean;
  SitePropHubSiteConnectedDescription: string;
  SitePropHubSiteConnectedLabel: string;
  SitePropIncreaseStorage: boolean;
  SitePropIncreaseStorageDescription: string;
  SitePropIncreaseStorageLabel: string;
  SitePropLockState: boolean;
  SitePropLockStateDescription: string;
  SitePropLockStateLabel: string;
  SitePropSensitivityLabelIdEnabled: boolean;
  SitePropSensitivityLabelIdEnabledDescription: string;
  SitePropSensitivityLabelIdEnabledLabel: string;
  SitePropShareByEmailEnabled: boolean;
  SitePropShareByEmailEnabledDescription: string;
  SitePropShareByEmailEnabledLabel: string;
  SitePropSocialBarOnSitePagesDisabled: boolean;
  SitePropSocialBarOnSitePagesDisabledDescription: string;
  SitePropSocialBarOnSitePagesDisabledLabel: string;
  SitePropStorageUsed: boolean;
  SitePropStorageUsedDescription: string;
  SitePropStorageUsedLabel: string;
  SitePropTemplate: boolean;
  SitePropTemplateDescription: string;
  SitePropTemplateLabel: string;
  SitePropTitle: boolean;
  SitePropTitleDescription: string;
  SitePropTitleLabel: string;
  WebPropCommentsOnSitePagesDisabled: boolean;
  WebPropCommentsOnSitePagesDisabledDescription: string;
  WebPropCommentsOnSitePagesDisabledLabel: string;
  WebPropExcludeFromOfflineClient: boolean;
  WebPropExcludeFromOfflineClientDescription: string;
  WebPropExcludeFromOfflineClientLabel: string;
  WebPropSearchPropertyDescription: string;
  WebPropSearchPropertyKey: string;
  WebPropSearchPropertyLabel: string;
  WebPropSearchPropertyTabName: string;
  WebPropSearchPropertyValues: string;
  WebPropSearchScope: boolean;
  WebPropSearchScopeDescription: string;
  WebPropSearchScopeLabel: string;
  WebPropTemplate: boolean;
  WebPropTemplateDescription: string;
  WebPropTemplateLabel: string;
  WebPropTitle: boolean;
  WebPropTitleDescription: string;
  WebPropTitleLabel: string;
}

// Reference the solution
import "main-lib";
declare const SiteAdmin: {
  appDescription: string;
  render: (props: {
    cloudEnv?: string;
    context?: WebPartContext;
    el: HTMLElement;
    title?: string;
    maxStorageDesc?: string;
    maxStorageSize?: number;
    reportProps?: {
      docSearchFileExt?: string;
      docSearchKeywords?: string;
    }
    searchProps?: {
      description: string;
      key: string;
      label: string;
      tabName: string;
      values: string;
    }
    siteProps: {
      [key: string]: {
        description: string;
        disabled: boolean;
        label: string;
      }
    }
    webProps: {
      [key: string]: {
        description: string;
        disabled: boolean;
        label: string;
      }
    }
  }) => void;
  updateTheme: (theme: IReadonlyTheme) => void;
};

export default class SiteAdminWebPart extends BaseClientSideWebPart<ISiteAdminWebPartProps> {
  public render(): void {
    const propValues = (this.properties as any);
    const siteProps: {
      [key: string]: {
        description: string;
        disabled: boolean;
        label: string;
      }
    } = {};
    const webProps: {
      [key: string]: {
        description: string;
        disabled: boolean;
        label: string;
      }
    } = {};

    // Parse the site properties
    for (let i = 0; i < this._siteProps.length; i++) {
      const propName = this._siteProps[i];
      const key = propName.replace("SiteProp", "");

      // Add the property
      siteProps[key] = {
        description: propValues[propName + "Description"],
        disabled: propValues[propName],
        label: propValues[propName + "Label"]
      };
    }

    // Parse the web properties
    for (let i = 0; i < this._webProps.length; i++) {
      const propName = this._webProps[i];
      const key = propName.replace("WebProp", "");

      // Add the property
      webProps[key] = {
        description: propValues[propName + "Description"],
        disabled: propValues[propName],
        label: propValues[propName + "Label"]
      };
    }

    // Render the solution
    SiteAdmin.render({
      cloudEnv: this.properties.CloudEnvironment,
      context: this.context,
      el: this.domElement,
      maxStorageDesc: this.properties.MaxStorageDescription,
      maxStorageSize: this.properties.MaxStorage,
      reportProps: {
        docSearchFileExt: this.properties.ReportsDocSearchFileExt,
        docSearchKeywords: this.properties.ReportsDocSearchKeywords
      },
      searchProps: {
        description: this.properties.WebPropSearchPropertyDescription,
        key: this.properties.WebPropSearchPropertyKey,
        label: this.properties.WebPropSearchPropertyLabel,
        tabName: this.properties.WebPropSearchPropertyTabName,
        values: this.properties.WebPropSearchPropertyValues
      },
      siteProps,
      title: this.properties.AppTitle,
      webProps
    });
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // Update the theme
    SiteAdmin.updateTheme(currentTheme);
  }

  protected get dataVersion(): Version {
    return Version.parse(this.context.manifest.version);
  }

  protected get disableReactivePropertyChanges(): boolean { return true; }

  private _siteProps: string[] = [
    "SitePropCommentsOnSitePagesDisabled",
    "SitePropContainsAppCatalog",
    "SitePropCreated",
    "SitePropCustomScriptsEnabled",
    "SitePropDisableCompanyWideSharingLinks",
    "SitePropExcludeFromOfflineClient",
    "SitePropHubSite",
    "SitePropHubSiteConnected",
    "SitePropIncreaseStorage",
    "SitePropLockState",
    "SitePropSensitivityLabelId",
    "SitePropShareByEmailEnabled",
    "SitePropSocialBarOnSitePagesDisabled",
    "SitePropStorageUsed",
    "SitePropTemplate",
    "SitePropTitle"
  ];

  private _webProps: string[] = [
    "WebPropCommentsOnSitePagesDisabled",
    "WebPropExcludeFromOfflineClient",
    "WebPropSearchScope",
    "WebPropTemplate",
    "WebPropTitle"
  ];

  private generateGroup(groupName: string, props: string[]): IPropertyPaneGroup {
    const groups: IPropertyPaneGroup = {
      groupName,
      groupFields: [
      ]
    };

    // Parse the prop names
    for (let i = 0; i < props.length; i++) {
      const propName = props[i];
      const hideProp = (this.properties as any)[propName] as boolean;

      // Add the property fields
      groups.groupFields.push(PropertyPaneToggle(propName, {
        label: (strings as any)[propName],
        checked: hideProp,
        onText: "The admin will not be able to change this value.",
        offText: "The admin can make changes to this property"
      }));
      if (!hideProp) {
        groups.groupFields.push(PropertyPaneTextField(propName + "Label", {
          label: (strings as any)[propName] + " Label",
          description: "The label displayed in the form.",
          value: (this.properties as any)[propName + "Label"]
        }));
        groups.groupFields.push(PropertyPaneTextField(propName + "Description", {
          label: (strings as any)[propName] + " Description",
          description: "The description displayed in the form",
          multiline: true,
          value: (this.properties as any)[propName + "Description"]
        }));
      }
    }

    // Return the properties
    return groups;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "App Settings",
              groupFields: [
                PropertyPaneTextField("AppTitle", {
                  label: strings.AppTitle,
                  description: "The title displayed.",
                  value: this.properties.AppTitle
                }),
                PropertyPaneDropdown("CloudEnvironment", {
                  label: strings.CloudEnvironment,
                  selectedKey: this.properties.CloudEnvironment,
                  options: [
                    { key: "https://graph.microsoft.com", text: "Default" },
                    { key: "https://microsoftgraph.chinacloudapi.cn", text: "China" },
                    { key: "https://graph.microsoft.us", text: "USL4" },
                    { key: "https://dod-graph.microsoft.us", text: "USL5" }
                  ]
                })
              ]
            }
          ]
        },
        {
          groups: [
            this.generateGroup("Site Properties", this._siteProps)
          ]
        },
        {
          groups: [
            this.generateGroup("Web Properties", this._webProps)
          ]
        },
        {
          groups: [
            {
              groupName: "Storage",
              groupFields: [
                PropertyPaneDropdown("MaxStorage", {
                  label: strings.MaxStorage,
                  options: [
                    { key: "5", text: "5 TB" },
                    { key: "10", text: "10 TB" },
                    { key: "15", text: "15 TB" },
                    { key: "20", text: "20 TB" },
                    { key: "25", text: "25 TB" }
                  ]
                }),
                PropertyPaneTextField("MaxStorageDescription", {
                  label: strings.MaxStorageDescription,
                  description: "The description to display when the max storage threshold has been reached.",
                  value: this.properties.MaxStorageDescription
                }),
              ]
            },
            {
              groupName: "Audit Tools",
              groupFields: [
                PropertyPaneTextField("ReportsDocSearchFileExt", {
                  label: strings.ReportsDocSearchFileExt,
                  description: "The default file extensions to search.",
                  value: this.properties.ReportsDocSearchFileExt
                }),
                PropertyPaneTextField("ReportsDocSearchKeywords", {
                  label: strings.ReportsDocSearchKeywords,
                  description: "The default keywords to search for.",
                  value: this.properties.ReportsDocSearchKeywords
                }),
              ]
            },
            {
              groupName: "Search Property",
              groupFields: [
                PropertyPaneTextField("WebPropSearchPropertyKey", {
                  label: strings.WebPropSearchPropertyKey,
                  description: "The custom property key to set for a site.",
                  value: this.properties.WebPropSearchPropertyKey
                }),
                PropertyPaneTextField("WebPropSearchPropertyTabName", {
                  label: strings.WebPropSearchPropertyTabName,
                  description: "The custom tab name for this property.",
                  value: this.properties.WebPropSearchPropertyTabName
                }),
                PropertyPaneTextField("WebPropSearchPropertyLabel", {
                  label: strings.WebPropSearchPropertyLabel,
                  description: "The label to display for the custom property in the form.",
                  value: this.properties.WebPropSearchPropertyLabel
                }),
                PropertyPaneTextField("WebPropSearchPropertyDescription", {
                  label: strings.WebPropSearchPropertyDescription,
                  description: "The description to display for the custom  property in the form.",
                  value: this.properties.WebPropSearchPropertyDescription
                }),
                PropertyPaneTextField("WebPropSearchPropertyValues", {
                  label: strings.WebPropSearchPropertyValues,
                  description: "A CSV formatted string containing the values for the search property. If not set, a textbox will be displayed for the user to enter the value.",
                  value: this.properties.WebPropSearchPropertyValues
                }),
              ]
            }]
        },
        {
          groups: [
            {
              groupName: "About this app:",
              groupFields: [
                PropertyPaneLabel('version', {
                  text: "Version: " + this.context.manifest.version
                }),
                PropertyPaneLabel('description', {
                  text: SiteAdmin.appDescription
                }),
                PropertyPaneLabel('about', {
                  text: "We think adding sprinkles to a donut just makes it better! SharePoint Sprinkles builds apps that are sprinkled on top of SharePoint, making your experience even better. Check out our site below to discover other SharePoint Sprinkles apps, or connect with us on GitHub."
                }),
                PropertyPaneLabel('support', {
                  text: "Are you having a problem or do you have a great idea for this app? Visit our GitHub link below to open an issue and let us know!"
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLink('requestEnhancement', {
                  href: "https://github.com/spsprinkles/site-admin/issues",
                  text: "Request an Enhancement",
                  target: "_blank"
                }),
                PropertyPaneLink('supportLink', {
                  href: "https://github.com/spsprinkles/site-admin/issues",
                  text: "Submit an Issue",
                  target: "_blank"
                }),
                PropertyPaneLink('sourceLink', {
                  href: "https://github.com/spsprinkles/site-admin",
                  text: "View Source on GitHub",
                  target: "_blank"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
