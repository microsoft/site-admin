import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration, IPropertyPaneGroup,
  PropertyPaneDropdown, PropertyPaneHorizontalRule,
  PropertyPaneLabel, PropertyPaneLink, PropertyPaneTextField, PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'SiteAdminWebPartStrings';

// Include the asset files for the page generator
const imageReferences = [
  require("./assets/956444907.png"),
  require("./assets/704156399.png"),
  require("./assets/1439995666.png"),
  require("./assets/2403381788.png"),
  require("./assets/2009954068.png"),
  require("./assets/394511437.png"),
  require("./assets/178240801.png"),
  require("./assets/2492974796.png"),
  require("./assets/878621426.png"),
  require("./assets/3639654909.png"),
  require("./assets/3423846637.png"),
  require("./assets/734992370.png"),
  require("./assets/3614134656.png"),
  require("./assets/1004145298.png"),
  require("./assets/3707825808.png"),
  require("./assets/3194998129.png"),
  require("./assets/782453716.png"),
  require("./assets/926485291.png"),
  require("./assets/2498076700.png"),
  require("./assets/3608935117.png"),
  require("./assets/4193189816.png")
]

export interface ISiteAdminWebPartProps {
  AppTitle: string;
  DisableSensitivityLabelOverride: boolean;
  MaxStorage: number;
  MaxStorageDescription: string;
  ReportsDocRententionYears: string;
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
  SitePropExcludeFromSearch: boolean;
  SitePropExcludeFromSearchDescription: string;
  SitePropExcludeFromSearchLabel: string;
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
  SitePropSensitivityLabelId: boolean;
  SitePropSensitivityLabelIdDescription: string;
  SitePropSensitivityLabelIdLabel: string;
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
  WebPropExcludeFromSearch: boolean;
  WebPropExcludeFromSearchDescription: string;
  WebPropExcludeFromSearchLabel: string;
  WebPropSearchPropertyDescription: string;
  WebPropSearchPropertyKey: string;
  WebPropSearchPropertyLabel: string;
  WebPropSearchPropertyManagedProperty: string;
  WebPropSearchPropertyReportName: string;
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
    context?: WebPartContext;
    el: HTMLElement;
    title?: string;
    disableSensitivityLabelOverride?: boolean;
    imageReferences: string[];
    maxStorageDesc?: string;
    maxStorageSize?: number;
    reportProps?: {
      docRententionYears?: string;
      docSearchFileExt?: string;
      docSearchKeywords?: string;
    }
    searchProps?: {
      description: string;
      key: string;
      label: string;
      managedProperty: string;
      reportName: string;
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
      context: this.context,
      el: this.domElement,
      disableSensitivityLabelOverride: this.properties.DisableSensitivityLabelOverride,
      imageReferences,
      maxStorageDesc: this.properties.MaxStorageDescription,
      maxStorageSize: this.properties.MaxStorage,
      reportProps: {
        docRententionYears: this.properties.ReportsDocRententionYears,
        docSearchFileExt: this.properties.ReportsDocSearchFileExt,
        docSearchKeywords: this.properties.ReportsDocSearchKeywords
      },
      searchProps: {
        description: this.properties.WebPropSearchPropertyDescription,
        key: this.properties.WebPropSearchPropertyKey,
        label: this.properties.WebPropSearchPropertyLabel,
        managedProperty: this.properties.WebPropSearchPropertyManagedProperty,
        reportName: this.properties.WebPropSearchPropertyReportName,
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
    "SitePropExcludeFromSearch",
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

  private _readOnlyProps: string[] = [
    "SitePropCreated",
    "SitePropHubSite",
    "SitePropHubSiteConnected",
    "SitePropStorageUsed",
    "SitePropTemplate",
    "SitePropTitle",
    "WebPropTemplate",
    "WebPropTitle"
  ]

  private _webProps: string[] = [
    "WebPropCommentsOnSitePagesDisabled",
    "WebPropExcludeFromOfflineClient",
    "WebPropExcludeFromSearch",
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

      // See if this is a readonly property
      if (this._readOnlyProps.indexOf(propName) >= 0) {
        // Add the property fields
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
      } else {
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
                })
              ]
            },
            {
              groupName: "Sensitivity Labels",
              groupFields: [
                PropertyPaneToggle("DisableSensitivityLabelOverride", {
                  label: "Disable Sensivitity Label Override:",
                  checked: this.properties.DisableSensitivityLabelOverride,
                  onText: "The admins will only be allowed to apply sensitivity labels to files that haven't been labeled.",
                  offText: "The admins will be allowed to attempt to override sensitivity labels."
                })
              ]
            },
            {
              groupName: "Storage",
              groupFields: [
                PropertyPaneDropdown("MaxStorage", {
                  label: strings.MaxStorage,
                  options: [
                    { key: "1", text: "1 TB" },
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
                PropertyPaneDropdown("ReportsDocRententionYears", {
                  label: strings.ReportsDocRententionYears,
                  selectedKey: this.properties.ReportsDocRententionYears,
                  options: [
                    { key: "1", text: "1" },
                    { key: "2", text: "2" },
                    { key: "3", text: "3" },
                    { key: "4", text: "4" },
                    { key: "5", text: "5" },
                    { key: "6", text: "6" },
                    { key: "7", text: "7" },
                    { key: "8", text: "8" },
                    { key: "9", text: "9" },
                    { key: "10", text: "10" },
                    { key: "15", text: "15" },
                    { key: "20", text: "20" }
                  ]
                }),
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
                PropertyPaneTextField("WebPropSearchPropertyManagedProperty", {
                  label: strings.WebPropSearchPropertyManagedProperty,
                  description: "The mapped managed property for this custom property.",
                  value: this.properties.WebPropSearchPropertyManagedProperty
                }),
                PropertyPaneTextField("WebPropSearchPropertyTabName", {
                  label: strings.WebPropSearchPropertyTabName,
                  description: "The custom tab name for this property.",
                  value: this.properties.WebPropSearchPropertyTabName
                }),
                PropertyPaneTextField("WebPropSearchPropertyReportName", {
                  label: strings.WebPropSearchPropertyReportName,
                  description: "The report name to for this custom property.",
                  value: this.properties.WebPropSearchPropertyReportName
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
                  href: "https://github.com/microsoft/site-admin/issues",
                  text: "Request an Enhancement",
                  target: "_blank"
                }),
                PropertyPaneLink('reportIssue', {
                  href: "https://github.com/microsoft/site-admin/issues",
                  text: "Submit an Issue",
                  target: "_blank"
                }),
                PropertyPaneLink('discussions', {
                  href: "https://github.com/microsoft/site-admin/discussions",
                  text: "Ask a Question",
                  target: "_blank"
                }),
                PropertyPaneLink('sourceLink', {
                  href: "https://github.com/microsoft/site-admin",
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
