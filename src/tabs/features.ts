import { Components } from "gd-sprest-bs";
import { IProp } from "../app";
import { DataSource, IRequest, RequestTypes } from "../ds";
import { Tab } from "./base";
import { IChangeRequest } from "./changes";

/**
 * Features Tab
 */
export class FeaturesTab extends Tab<{
    CommentsOnSitePagesDisabled: boolean;
    DisableCompanyWideSharingLinks: boolean;
    ExcludeFromOfflineClient: string;
    SocialBarOnSitePagesDisabled: boolean;
}, {
    DisableCompanyWideSharingLinks: IRequest
}, {
    CommentsOnSitePagesDisabled: boolean;
    SocialBarOnSitePagesDisabled: boolean;
}> {
    // Constructor
    constructor(el: HTMLElement, props: { [key: string]: IProp; }) {
        super(el, props, "Site");

        // Set the current values
        this._currValues = {
            CommentsOnSitePagesDisabled: DataSource.Site.CommentsOnSitePagesDisabled,
            DisableCompanyWideSharingLinks: DataSource.Site.DisableCompanyWideSharingLinks,
            ExcludeFromOfflineClient: null,
            SocialBarOnSitePagesDisabled: DataSource.Site.SocialBarOnSitePagesDisabled
        }

        // Render the tab
        this.render();
    }

    // Add the custom requests
    onGetRequests(): IChangeRequest[] {
        let requests: IChangeRequest[] = [];

        // Check the search property
        if (this._currValues.ExcludeFromOfflineClient) {
            let hideContent = this._currValues.ExcludeFromOfflineClient == "Hide";

            // Parse the webs
            for (let i = 0; i < DataSource.SiteItems.length; i++) {
                let item = DataSource.SiteItems[i];

                // Add the request
                requests.push({
                    oldValue: "",
                    newValue: hideContent,
                    scope: this._scope,
                    property: "ExcludeFromOfflineClient",
                    url: item.value
                });
            }
        }

        // Return the requests
        return requests;
    }

    // Renders the tab
    private render() {
        // Render the form
        Components.Form({
            el: this._el,
            className: "row",
            groupClassName: "col-4 mb-3",
            controls: [
                {
                    name: "CommentsOnSitePagesDisabled",
                    label: this._props["CommentsOnSitePagesDisabled"].label,
                    description: this._props["CommentsOnSitePagesDisabled"].description,
                    isDisabled: this._props["CommentsOnSitePagesDisabled"].disabled,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.CommentsOnSitePagesDisabled,
                    onChange: item => {
                        let value = item ? true : false;

                        // See if we are changing the value
                        if (this._currValues.CommentsOnSitePagesDisabled != value) {
                            // Set the value
                            this._newValues.CommentsOnSitePagesDisabled = value;
                        } else {
                            // Remove the value
                            delete this._newValues.CommentsOnSitePagesDisabled;
                        }
                    }
                } as Components.IFormControlPropsSwitch,
                {
                    name: "DisableCompanyWideSharingLinks",
                    label: this._props["DisableCompanyWideSharingLinks"].label,
                    description: this._props["DisableCompanyWideSharingLinks"].description,
                    isDisabled: this._props["DisableCompanyWideSharingLinks"].disabled,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.DisableCompanyWideSharingLinks,
                    onChange: item => {
                        let value = item ? true : false;

                        // See if we are changing the value
                        if (this._currValues.DisableCompanyWideSharingLinks != value) {
                            // Set the value
                            this._requestItems.DisableCompanyWideSharingLinks = {
                                key: RequestTypes.DisableCompanyWideSharingLinks,
                                message: `The request to ${value ? "disable" : "enable"} company wide sharing links will be processed within 5 minutes.`,
                                value
                            };
                        } else {
                            // Remove the value
                            delete this._requestItems.DisableCompanyWideSharingLinks;
                        }
                    }
                } as Components.IFormControlPropsSwitch,
                {
                    name: "ExcludeFromOfflineClient",
                    label: this._props["ExcludeFromOfflineClient"].label,
                    description: "This will apply to all sites.",
                    isDisabled: this._props["ExcludeFromOfflineClient"].disabled,
                    type: Components.FormControlTypes.Dropdown,
                    items: [
                        { text: "", value: "" },
                        { text: "Show", value: "Show" },
                        { text: "Hide", value: "Hide" }
                    ],
                    onChange: item => {
                        // Set the value
                        this._currValues.ExcludeFromOfflineClient = item ? item.value : null;
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "SocialBarOnSitePagesDisabled",
                    label: this._props["SocialBarOnSitePagesDisabled"].label,
                    description: this._props["SocialBarOnSitePagesDisabled"].description,
                    isDisabled: this._props["SocialBarOnSitePagesDisabled"].disabled,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.SocialBarOnSitePagesDisabled,
                    onChange: item => {
                        let value = item ? true : false;

                        // See if we are changing the value
                        if (this._currValues.SocialBarOnSitePagesDisabled != value) {
                            // Set the value
                            this._newValues.SocialBarOnSitePagesDisabled = value;
                        } else {
                            // Remove the value
                            delete this._newValues.SocialBarOnSitePagesDisabled;
                        }
                    }
                } as Components.IFormControlPropsSwitch
            ]
        });
    }
}