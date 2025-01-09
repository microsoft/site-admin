import { Components } from "gd-sprest-bs";
import { ITab } from "./tab.d";
import { DataSource, IRequest, RequestTypes } from "../ds";

/**
 * Features Tab
 */
export class FeaturesTab implements ITab {
    private _disableProps: string[] = null;
    private _el: HTMLElement = null;

    // The current values
    private _currValues: {
        CommentsOnSitePagesDisabled: boolean;
        DisableCompanyWideSharingLinks: boolean;
        ShareByEmailEnabled: boolean;
        SocialBarOnSitePagesDisabled: boolean;
    } = null;

    // The new request items to create
    private _newRequestItems: {
        DisableCompanyWideSharingLinks?: IRequest;
    } = {};

    // The new values requested
    private _newSiteValues: {
        CommentsOnSitePagesDisabled?: boolean;
        ShareByEmailEnabled?: boolean;
        SocialBarOnSitePagesDisabled?: boolean;
    } = {};

    // Constructor
    constructor(el: HTMLElement, disableProps: string[] = []) {
        this._disableProps = disableProps;
        this._el = el;

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

    // Returns the site properties to update
    getProps(): { [key: string]: string | number | boolean; } { return this._newSiteValues; }

    // Returns the new request items to create
    getRequests(): IRequest[] {
        // Get the request items to create
        let requests: IRequest[] = [];
        for (let key in this._newRequestItems) {
            // Append the request
            requests.push(this._newRequestItems[key]);
        }

        // Return the requests
        return requests;
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
                            this._newSiteValues.CommentsOnSitePagesDisabled = value;
                        } else {
                            // Remove the value
                            delete this._newSiteValues.CommentsOnSitePagesDisabled;
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
                            this._newRequestItems.DisableCompanyWideSharingLinks = {
                                key: RequestTypes.DisableCompanyWideSharingLinks,
                                message: `The request to ${value ? "disable" : "enable"} company wide sharing links will be processed within 5 minutes.`,
                                value
                            };
                        } else {
                            // Remove the value
                            delete this._newRequestItems.DisableCompanyWideSharingLinks;
                        }
                    }
                } as Components.IFormControlPropsSwitch,
                {
                    name: "ShareByEmailEnabled",
                    label: "Share By Email Enabled:",
                    description: "Disables the offline sync feature in all libraries.",
                    isDisabled: this._disableProps.indexOf("ShareByEmailEnabled") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.ShareByEmailEnabled,
                    onChange: item => {
                        let value = item ? true : false;

                        // See if we are changing the value
                        if (this._currValues.ShareByEmailEnabled != value) {
                            // Set the value
                            this._newSiteValues.ShareByEmailEnabled = value;
                        } else {
                            // Remove the value
                            delete this._newSiteValues.ShareByEmailEnabled;
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
                            this._newSiteValues.SocialBarOnSitePagesDisabled = value;
                        } else {
                            // Remove the value
                            delete this._newSiteValues.SocialBarOnSitePagesDisabled;
                        }
                    }
                } as Components.IFormControlPropsSwitch
            ]
        });
    }
}