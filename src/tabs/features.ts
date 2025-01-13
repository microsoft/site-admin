import { Components } from "gd-sprest-bs";
import { IProp } from "../app";
import { DataSource, IRequest, RequestTypes } from "../ds";
import { Tab } from "./base";

/**
 * Features Tab
 */
export class FeaturesTab extends Tab<{
    CommentsOnSitePagesDisabled: boolean;
    DisableCompanyWideSharingLinks: boolean;
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