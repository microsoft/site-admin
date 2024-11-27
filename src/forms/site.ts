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

    // The current values
    private _currValues: {
        CommentsOnSitePagesDisabled: boolean;
        ContainsAppCatalog: boolean;
        CustomScriptsEnabled: boolean;
        DisableCompanyWideSharingLinks: boolean;
        IncreaseStorage: boolean;
        LockState: string;
        ShareByEmailEnabled: boolean;
        SocialBarOnSitePagesDisabled: boolean;
        TeamsConnected: boolean;
        UsageSize: string;
        UsageUsed: string;
    } = null;

    constructor(el: HTMLElement, disableProps: string[] = []) {
        // Save the properties
        this._el = el;
        this._disableProps = disableProps;

        // Set the current values
        this._currValues = {
            CommentsOnSitePagesDisabled: DataSource.Site.CommentsOnSitePagesDisabled,
            ContainsAppCatalog: false,
            CustomScriptsEnabled: Helper.hasPermissions(DataSource.Site.RootWeb.EffectiveBasePermissions, SPTypes.BasePermissionTypes.AddAndCustomizePages),
            DisableCompanyWideSharingLinks: DataSource.Site.DisableCompanyWideSharingLinks,
            IncreaseStorage: false,
            LockState: DataSource.Site.ReadOnly && DataSource.Site.WriteLocked ? "ReadOnly" : "Unlock",
            ShareByEmailEnabled: DataSource.Site.ShareByEmailEnabled,
            SocialBarOnSitePagesDisabled: DataSource.Site.SocialBarOnSitePagesDisabled,
            TeamsConnected: DataSource.Site.GroupId && DataSource.Site.GroupId != "00000000-0000-0000-0000-000000000000" && DataSource.Web.AllProperties["TeamifyHidden"] != "TRUE",
            UsageSize: DataSource.formatBytes(DataSource.Site.Usage.Storage),
            UsageUsed: DataSource.Site.Usage.StoragePercentageUsed + "%"
        }

        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

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
            className: "row",
            groupClassName: "col-4 mb-5",
            controls: [
                {
                    name: "Created",
                    label: "Created:",
                    description: "The date the site was created.",
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Web.Created
                },
                {
                    name: "Title",
                    label: "Title:",
                    description: "The title of the site collection.",
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Web.Title
                },
                {
                    name: "CommentsOnSitePagesDisabled",
                    label: "Comments On Site Pages Disabled:",
                    description: "If true, comments on modern site pages will be disabled.",
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
                            DataSource.hasAppCatalog(DataSource.Site.Url).then(hasAppCatalog => {
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
                    description: `<ul class="mt-3">
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
                    name: "IncreaseStorage",
                    label: "Increase Storage:",
                    description: "Enable to increase the site collection storage size.",
                    isDisabled: this._disableProps.indexOf("IncreaseStorage") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.IncreaseStorage
                } as Components.IFormControlPropsSwitch,
                {
                    name: "UsageSize",
                    label: "Usage Size:",
                    description: "The total size of the site collection.",
                    type: Components.FormControlTypes.Readonly,
                    value: this._currValues.UsageSize
                },
                {
                    name: "UsageUsed",
                    label: "Usage Used:",
                    description: "The total percent used for the site collection.",
                    type: Components.FormControlTypes.Readonly,
                    value: this._currValues.UsageUsed
                },
                {
                    name: "SocialBarOnSitePagesDisabled",
                    label: "Social Bar On Site Pages Disabled:",
                    description: "The search scope for the site to target. Default is 'Site'.",
                    isDisabled: this._disableProps.indexOf("SocialBarOnSitePagesDisabled") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.SocialBarOnSitePagesDisabled
                } as Components.IFormControlPropsSwitch,
                {
                    name: "TeamsConnected",
                    label: "Teams Connected:",
                    description: "The search scope for the site to target. Default is 'Site'.",
                    isDisabled: this._currValues.TeamsConnected || this._disableProps.indexOf("TeamsConnected") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.TeamsConnected
                } as Components.IFormControlPropsSwitch
            ]
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
                // Skip disabled and read-only controls
                let ctrl = this._form.getControl(key);
                if (ctrl.props.isDisabled || ctrl.props.isReadonly) { continue; }

                // Compare the values for a change
                let value = typeof (values[key]) === "boolean" ? values[key] : values[key].data || values[key].value;
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
                        case "IncreaseStorage":
                            // Add a request for this request
                            requests.push({
                                key: RequestTypes.IncreaseStorage,
                                message: `The request to increase storage for the site collection will be processed within 5 minutes.`,
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
                        case "TeamsConnected":
                            // Add a request for this request
                            requests.push({
                                key: RequestTypes.TeamsConnected,
                                message: `The request to connect the site to teams will be processed within 5 minutes.`,
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
            DataSource.addRequest(DataSource.Site.Url, requests).then((responses) => {
                // See if an update is needed
                if (updateFlags.CommentsOnSitePagesDisabled || updateFlags.ShareByEmailEnabled || updateFlags.SocialBarOnSitePagesDisabled) {
                    // Show a loading dialog
                    LoadingDialog.setHeader("Updating Site Collection");
                    LoadingDialog.setBody("This will close after the changes complete.");
                    LoadingDialog.show();

                    // Save the changes
                    DataSource.Site.update(props).execute(() => {
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