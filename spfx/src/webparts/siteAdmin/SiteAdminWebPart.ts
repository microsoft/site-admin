import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration, IPropertyPaneGroup,
  PropertyPaneHorizontalRule, PropertyPaneLabel, PropertyPaneLink, PropertyPaneTextField, PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'SiteAdminWebPartStrings';

export interface ISiteAdminWebPartProps {
  hideProps: string[];
  SitePropCommentsOnSitePagesDisabled: boolean;
  SitePropContainsAppCatalog: boolean;
  SitePropContainsAppCatalogUrl: string;
  SitePropDisableCompanyWideSharingLinks: boolean;
  SitePropShareByEmailEnabled: boolean;
  SitePropSocialBarOnSitePagesDisabled: boolean;
  WebPropCommentsOnSitePagesDisabled: boolean;
  WebPropExcludeFromOfflineClient: boolean;
  WebPropSearchPropertyDescription: string;
  WebPropSearchPropertyKey: string;
  WebPropSearchPropertyLabel: string;
  WebPropSearchScope: boolean;
  WebPropWebTemplate: boolean;
}

// Reference the solution
import "main-lib";
declare const SiteAdmin: {
  appDescription: string;
  render: (props: {
    context?: WebPartContext;
    el: HTMLElement;
    disableSiteProps?: string[];
    disableWebProps?: string[];
    searchProp?: {
      description: string;
      key: string;
      label: string;
    }
    siteUrls?: string[];
    webUrls?: string[];
  }) => void;
  updateTheme: (theme: IReadonlyTheme) => void;
};

export default class SiteAdminWebPart extends BaseClientSideWebPart<ISiteAdminWebPartProps> {
  public render(): void {
    const disableSiteProps: string[] = [];
    const disableWebProps: string[] = [];
    const siteUrls: string[] = [];
    const webUrls: string[] = [];

    // Parse the site properties
    for (let i = 0; i < this._siteProps.length; i++) {
      const propName = this._siteProps[i];

      // See if this one is selected to be hidden
      if ((this.properties as any)[propName]) {
        // Add the property to hide
        disableSiteProps.push(propName.replace("SiteProp", ""));
      }

      // Add the associated api with this property
      siteUrls.push((this.properties as any)[propName + "Url"])
    }

    // Parse the web properties
    for (let i = 0; i < this._webProps.length; i++) {
      const propName = this._webProps[i];

      // See if this one is selected to be hidden
      if ((this.properties as any)[propName]) {
        // Add the property to hide
        disableWebProps.push(propName.replace("WebProp", ""));
      }

      // See if there is an associated api with this property
      webUrls.push((this.properties as any)[propName + "Url"]);
    }

    // Render the solution
    SiteAdmin.render({
      context: this.context,
      el: this.domElement,
      disableSiteProps,
      disableWebProps,
      siteUrls,
      webUrls,
      searchProp: {
        description: this.properties.WebPropSearchPropertyDescription,
        key: this.properties.WebPropSearchPropertyKey,
        label: this.properties.WebPropSearchPropertyLabel
      }
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
    "SitePropDisableCompanyWideSharingLinks",
    "SitePropShareByEmailEnabled",
    "SitePropSocialBarOnSitePagesDisabled"
  ];

  private _webProps: string[] = [
    "WebPropCommentsOnSitePagesDisabled",
    "WebPropExcludeFromOfflineClient",
    "WebPropSearchScope"
  ];

  private generateGroup(groupName: string, props: string[], azureUrl: string[] = []): IPropertyPaneGroup {
    const hideProps = this.properties.hideProps || [];
    const groups: IPropertyPaneGroup = {
      groupName,
      groupFields: [
      ]
    };

    // Parse the prop names
    for (let i = 0; i < props.length; i++) {
      const propName = props[i];
      const hideProp = hideProps.indexOf(propName) >= 0;
      const hasUrl = (strings as any)[propName + "Url"] ? true : false;

      // Add the property
      groups.groupFields.push(PropertyPaneToggle(propName, {
        label: (strings as any)[propName],
        checked: hideProp,
        onText: "The admin will not be able to change this value.",
        offText: "The admin can make changes to this property"
      }));

      // See if this has an associated api to call
      if (hasUrl) {
        groups.groupFields.push(PropertyPaneTextField(propName + "Url", {
          label: `The azure function api to execute for this property.`,
          multiline: true,
          value: (this.properties as any)[propName + "Url"]
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
            this.generateGroup("Site Properties:", this._siteProps, ["SitePropContainsAppCatalog"]),
          ]
        },
        {
          groups: [
            this.generateGroup("Web Properties:", this._webProps),
            {
              groupName: "Search Property",
              groupFields: [
                PropertyPaneTextField("WebPropSearchPropertyKey", {
                  label: strings.WebPropSearchPropertyKey,
                  description: "The custom property key to set for a site.",
                  value: this.properties.WebPropSearchPropertyKey
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
                })
              ]
            }
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
