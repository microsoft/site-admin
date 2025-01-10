import { Components } from "gd-sprest-bs";
import { DataSource, IRequest, RequestTypes } from "../ds";
import { Tab } from "./base";

/**
 * Features Tab
 */
export class FeaturesTab extends Tab<{
    CommentsOnSitePagesDisabled: boolean;
    DisableCompanyWideSharingLinks: boolean;
    ShareByEmailEnabled: boolean;
    SocialBarOnSitePagesDisabled: boolean;
}, {
    DisableCompanyWideSharingLinks: IRequest
}, {
    CommentsOnSitePagesDisabled: boolean;
    ShareByEmailEnabled: boolean;
    SocialBarOnSitePagesDisabled: boolean;
}> {

    // Constructor
    constructor(el: HTMLElement, disableProps: string[] = []) {
        super(el, disableProps, "Site");

        // Set the current values
        this._currValues = {
            CommentsOnSitePagesDisabled: DataSource.Site.CommentsOnSitePagesDisabled,
            DisableCompanyWideSharingLinks: DataSource.Site.DisableCompanyWideSharingLinks,
            ShareByEmailEnabled: DataSource.Site.ShareByEmailEnabled,
            SocialBarOnSitePagesDisabled: DataSource.Site.SocialBarOnSitePagesDisabled
        }

        // Render the tab
        this.render();
    }

    // Renders the tab
    private render() {
        // Render the form
        Components.Form({
            el: this._el,
            className: "row mt-2",
            groupClassName: "col-4 mb-3",
            controls: [
                {
                    name: "CommentsOnSitePagesDisabled",
                    label: "Comments On Site Pages Disabled:",
                    description: "If true, comments on modern site pages will be disabled.",
                    isDisabled: this._disableProps.indexOf("CommentsOnSitePagesDisabled") >= 0,
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
                    label: "Disable Company Wide Sharing Links:",
                    description: "If true, it will hide the comments on the site pages.",
                    isDisabled: this._disableProps.indexOf("DisableCompanyWideSharingLinks") >= 0,
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
                    name: "ShareByEmailEnabled",
                    label: "Enable Guest Access:",
                    isDisabled: this._disableProps.indexOf("ShareByEmailEnabled") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.ShareByEmailEnabled,
                    onChange: item => {
                        let value = item ? true : false;

                        // See if we are changing the value
                        if (this._currValues.ShareByEmailEnabled != value) {
                            // Set the value
                            this._newValues.ShareByEmailEnabled = value;
                        } else {
                            // Remove the value
                            delete this._newValues.ShareByEmailEnabled;
                        }
                    }
                } as Components.IFormControlPropsSwitch,
                {
                    name: "SocialBarOnSitePagesDisabled",
                    label: "Social Bar On Site Pages Disabled:",
                    description: "The search scope for the site to target. Default is 'Site'.",
                    isDisabled: this._disableProps.indexOf("SocialBarOnSitePagesDisabled") >= 0,
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