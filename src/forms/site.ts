import { LoadingDialog } from "dattatable";
import { Components, Helper, Types } from "gd-sprest-bs";

/**
 * Site Form
 */
export class Site {
    private _el: HTMLElement = null;
    private _form: Components.IForm = null;
    private _site: Types.SP.SiteOData = null;

    // The current values
    private _currValues: {
        CommentsOnSitePagesDisabled: boolean;
        DisableCompanyWideSharingLinks: boolean;
        ShareByEmailEnabled: boolean;
        SocialBarOnSitePagesDisabled: boolean;
    } = null;

    constructor(site: Types.SP.SiteOData, el: HTMLElement) {
        // Save the properties
        this._el = el;
        this._site = site;

        // Set the current values
        this._currValues = {
            CommentsOnSitePagesDisabled: this._site.CommentsOnSitePagesDisabled,
            DisableCompanyWideSharingLinks: this._site.DisableCompanyWideSharingLinks,
            ShareByEmailEnabled: this._site.ShareByEmailEnabled,
            SocialBarOnSitePagesDisabled: this._site.SocialBarOnSitePagesDisabled
        }

        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Render the header
        this.renderHeader();

        // Renders the form
        this.renderForm();

        // Renders the footer
        this.renderFooter();
    }

    // Renders the footer
    private renderFooter() {
        // Create the element
        let elFooter = document.createElement("div");
        elFooter.classList.add("mt-5");
        elFooter.classList.add("d-flex");
        elFooter.classList.add("justify-content-end");
        this._el.appendChild(elFooter);

        // Render the tooltips
        Components.Tooltip({
            el: elFooter,
            content: "Click to save the changes.",
            btnProps: {
                text: "Save Changes",
                onClick: () => {
                    let props = {};
                    let updateFl = false;
                    let values = this._form.getValues();

                    // See if an update is needed
                    if (this._currValues.CommentsOnSitePagesDisabled != values["CommentsOnSitePagesDisabled"]) { props["CommentsOnSitePagesDisabled"] = values["CommentsOnSitePagesDisabled"]; updateFl = true; }
                    if (this._currValues.DisableCompanyWideSharingLinks != values["DisableCompanyWideSharingLinks"]) { props["DisableCompanyWideSharingLinks"] = values["DisableCompanyWideSharingLinks"]; updateFl = true; }
                    if (this._currValues.ShareByEmailEnabled != values["ShareByEmailEnabled"]) { props["ShareByEmailEnabled"] = values["ShareByEmailEnabled"]; updateFl = true; }
                    if (this._currValues.SocialBarOnSitePagesDisabled != values["SocialBarOnSitePagesDisabled"]) { props["SocialBarOnSitePagesDisabled"] = values["SocialBarOnSitePagesDisabled"]; updateFl = true; }

                    // See if an update is needed
                    if (updateFl) {
                        // Show a loading dialog
                        LoadingDialog.setHeader("Updating Web");
                        LoadingDialog.setBody("This will close after the changes complete.");
                        LoadingDialog.show();

                        // Save the changes
                        this._site.update(props).execute(() => {
                            // Close the dialog
                            LoadingDialog.hide();
                        });
                    }
                }
            }
        });
    }

    // Renders the form
    private renderForm() {
        // Render the form
        this._form = Components.Form({
            el: this._el,
            value: this._currValues,
            groupClassName: "mb-3",
            controls: [
                {
                    name: "CommentsOnSitePagesDisabled",
                    label: "Comments On Site Pages Disabled:",
                    description: "The type of web.",
                    type: Components.FormControlTypes.Dropdown,
                    items: [
                        {
                            text: "true",
                            value: "true"
                        },
                        {
                            text: "false",
                            value: "false"
                        }
                    ]
                } as Components.IFormControlPropsDropdown,
                {
                    name: "DisableCompanyWideSharingLinks",
                    label: "Disable Company Wide Sharing Links:",
                    description: "If true, it will hide the comments on the site pages.",
                    type: Components.FormControlTypes.Dropdown,
                    items: [
                        {
                            text: "true",
                            value: "true"
                        },
                        {
                            text: "false",
                            value: "false"
                        }
                    ]
                } as Components.IFormControlPropsDropdown,
                {
                    name: "ShareByEmailEnabled",
                    label: "Share By Email Enabled:",
                    description: "Disables the offline sync feature in all libraries.",
                    type: Components.FormControlTypes.Dropdown,
                    items: [
                        {
                            text: "true",
                            value: "true"
                        },
                        {
                            text: "false",
                            value: "false"
                        }
                    ]
                } as Components.IFormControlPropsDropdown,
                {
                    name: "SocialBarOnSitePagesDisabled",
                    label: "Social Bar On Site Pages Disabled:",
                    description: "The search scope for the site to target. Default is 'Site'.",
                    type: Components.FormControlTypes.Dropdown,
                    items: [
                        {
                            text: "true",
                            value: "true"
                        },
                        {
                            text: "false",
                            value: "false"
                        }
                    ]
                } as Components.IFormControlPropsDropdown
            ]
        });
    }

    // Renders the header
    private renderHeader() {
        // Render the header
        Components.Jumbotron({
            el: this._el,
            lead: "Site Collection Settings",
            size: Components.JumbotronSize.Small,
            type: Components["JumbotronTypes"].Primary
        });
    }
}