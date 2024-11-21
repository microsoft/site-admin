import { LoadingDialog } from "dattatable";
import { Components, Helper, SPTypes, Types } from "gd-sprest-bs";
import { DataSource, IRequest, RequestTypes } from "../ds";
import { APIResponseModal } from "./response";

/**
 * Site Form
 */
export class Site {
    private _el: HTMLElement = null;
    private _form: Components.IForm = null;
    private _disableProps: string[] = null;
    private _site: Types.SP.SiteOData = null;

    // The current values
    private _currValues: {
        CommentsOnSitePagesDisabled: boolean;
        ContainsAppCatalog: boolean;
        CustomScriptsEnabled: boolean;
        DisableCompanyWideSharingLinks: boolean;
        LockState: string;
        ShareByEmailEnabled: boolean;
        SocialBarOnSitePagesDisabled: boolean;
    } = null;

    constructor(site: Types.SP.SiteOData, el: HTMLElement, disableProps: string[] = []) {
        // Save the properties
        this._el = el;
        this._disableProps = disableProps;
        this._site = site;

        // Set the current values
        this._currValues = {
            CommentsOnSitePagesDisabled: this._site.CommentsOnSitePagesDisabled,
            ContainsAppCatalog: false,
            CustomScriptsEnabled: Helper.hasPermissions(site.RootWeb.EffectiveBasePermissions, SPTypes.BasePermissionTypes.AddAndCustomizePages),
            DisableCompanyWideSharingLinks: this._site.DisableCompanyWideSharingLinks,
            LockState: this._site.ReadOnly && this._site.WriteLocked ? "ReadOnly" : "Unlock",
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
                    let values = this._form.getValues();

                    // Save the site properties
                    this.save(values);
                }
            }
        });
    }

    // Renders the form
    private renderForm() {
        // Render the form
        this._form = Components.Form({
            el: this._el,
            groupClassName: "mb-3",
            controls: [
                {
                    name: "CommentsOnSitePagesDisabled",
                    label: "Comments On Site Pages Disabled:",
                    description: "The type of web.",
                    isDisabled: this._disableProps.indexOf("CommentsOnSitePagesDisabled") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.CommentsOnSitePagesDisabled
                } as Components.IFormControlPropsSwitch,
                {
                    name: "DisableCompanyWideSharingLinks",
                    label: "Disable Company Wide Sharing Links:",
                    description: "If true, it will hide the comments on the site pages.",
                    isDisabled: this._disableProps.indexOf("DisableCompanyWideSharingLinks") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.DisableCompanyWideSharingLinks
                } as Components.IFormControlPropsSwitch,
                {
                    name: "ContainsAppCatalog",
                    label: "App Catalog Enabled:",
                    description: "True if this has a site collection app catalog available.",
                    isDisabled: this._disableProps.indexOf("ContainsAppCatalog") >= 0,
                    type: Components.FormControlTypes.Switch,
                    onControlRendering: ctrl => {
                        // Return a promise
                        return new Promise((resolve) => {
                            // See if an app catalog exists on this site
                            DataSource.hasAppCatalog(this._site.Url).then(hasAppCatalog => {
                                // Set the value
                                this._currValues.ContainsAppCatalog = hasAppCatalog;
                                ctrl.value = hasAppCatalog

                                // Resolve the request
                                resolve(ctrl);
                            });
                        });
                    }
                } as Components.IFormControlPropsSwitch,
                {
                    name: "CustomScriptsEnabled",
                    label: "Custom Scripts Enabled:",
                    description: "Enables the custom scripts feature for this site collection.",
                    isDisabled: this._disableProps.indexOf("CustomScriptsEnabled") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.CustomScriptsEnabled
                } as Components.IFormControlPropsSwitch,
                {
                    name: "LockState",
                    label: "Lock State:",
                    description: `<ul>
                        <li><b>Unlock:</b> To unlock the site and make it available to users.</li>
                        <li><b>Read Only:</b> To prevent users from adding/updating/deleting content.</li>
                        <li><b>No Access</b> To prevent users from accessing the site and its content.</li>
                    </ul>`,
                    isDisabled: this._disableProps.indexOf("LockState") >= 0,
                    type: Components.FormControlTypes.Dropdown,
                    value: this._currValues.LockState,
                    items: [
                        {
                            text: "Unlocked",
                            value: "Unlock"
                        },
                        {
                            text: "Read Only",
                            value: "ReadOnly"
                        },
                        {
                            text: "No Access",
                            value: "NoAccess"
                        }
                    ]
                } as Components.IFormControlPropsDropdown,
                {
                    name: "ShareByEmailEnabled",
                    label: "Share By Email Enabled:",
                    description: "Disables the offline sync feature in all libraries.",
                    isDisabled: this._disableProps.indexOf("ShareByEmailEnabled") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.ShareByEmailEnabled
                } as Components.IFormControlPropsSwitch,
                {
                    name: "SocialBarOnSitePagesDisabled",
                    label: "Social Bar On Site Pages Disabled:",
                    description: "The search scope for the site to target. Default is 'Site'.",
                    isDisabled: this._disableProps.indexOf("SocialBarOnSitePagesDisabled") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.SocialBarOnSitePagesDisabled
                } as Components.IFormControlPropsSwitch
            ]
        });
    }

    // Renders the header
    private renderHeader() {
        // Render the header
        Components.Jumbotron({
            el: this._el,
            className: "mb-2",
            lead: "Site Collection Settings",
            size: Components.JumbotronSize.Small,
            type: Components.JumbotronTypes.Primary
        });
    }

    // Saves the properties
    private save(values): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            let props = {};
            let requests: IRequest[] = [];
            let updateFlags: { [key: string]: boolean } = {};

            // Parse the keys
            for (let key in this._currValues) {
                let value = typeof (values[key]) === "boolean" ? values[key] : values[key].data;
                if (this._currValues[key] != value) {
                    // Set the flag
                    updateFlags[key] = false;

                    // See if we need to create a request
                    switch (key) {
                        case "ContainsAppCatalog":
                            // Add a request for this request
                            requests.push({
                                key: RequestTypes.AppCatalog,
                                message: `The request to ${values[key] ? "enable" : "disable"} the app catalog will be processed within 5 minutes.`,
                                value: values[key]
                            });
                            break;
                        case "CustomScriptsEnabled":
                            // Add a request for this request
                            requests.push({
                                key: RequestTypes.CustomScript,
                                message: `The request to ${values[key] ? "enable" : "disable"} custom scripts will be processed within 5 minutes.`,
                                value: values[key]
                            });
                            break;
                        case "DisableCompanyWideSharingLinks":
                            // Add a request for this request
                            requests.push({
                                key: RequestTypes.DisableCompanyWideSharingLinks,
                                message: `The request to ${values[key] ? "disable" : "enable"} company wide sharing links will be processed within 5 minutes.`,
                                value: values[key]
                            });
                            break;
                        case "LockState":
                            // Add a request for this request
                            requests.push({
                                key: RequestTypes.LockState,
                                message: `The request to make the site collection ${values[key].value == "NoAccess" ? "have" : "be"} '${values[key].text}' will be processed within 5 minutes.`,
                                value: values[key].value
                            });
                            break;
                        // We are updating a property
                        default:
                            // Add the property
                            props[key] = value;

                            // Set the flag
                            updateFlags[key] = true;
                            break;
                    }
                }
            }

            // Add the requests
            DataSource.addRequest(this._site.Url, requests).then((responses) => {
                // See if an update is needed
                if (updateFlags.CommentsOnSitePagesDisabled || updateFlags.ShareByEmailEnabled || updateFlags.SocialBarOnSitePagesDisabled) {
                    // Show a loading dialog
                    LoadingDialog.setHeader("Updating Site Collection");
                    LoadingDialog.setBody("This will close after the changes complete.");
                    LoadingDialog.show();

                    // Save the changes
                    this._site.update(props).execute(() => {
                        // Parse the keys
                        for (let key in updateFlags) {
                            // Update the current values
                            this._currValues[key] = values[key];

                            // Add the response
                            responses.push({
                                errorFl: false,
                                key,
                                message: "The property was updated successfully.",
                                value: values[key]
                            })
                        }

                        // Close the dialog
                        LoadingDialog.hide();

                        // Show the responses
                        new APIResponseModal(responses);

                        // Resolve the request
                        resolve();
                    });
                } else {
                    // Show the responses
                    new APIResponseModal(responses);

                    // Resolve the request
                    resolve();
                }
            });
        });
    }
}