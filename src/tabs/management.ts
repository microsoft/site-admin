import { Components, Helper, SPTypes } from "gd-sprest-bs";
import { ITab } from "./tab.d";
import { DataSource, IRequest, RequestTypes } from "../ds";

/**
 * Management Tab
 */
export class ManagementTab implements ITab {
    private _disableProps: string[] = null;
    private _el: HTMLElement = null;

    // The current values
    private _currValues: {
        ContainsAppCatalog: boolean;
        CustomScriptsEnabled: boolean;
        IncreaseStorage: boolean;
        LockState: string;
        SensitivityLabel: string;
        TeamsConnected: boolean;
    } = null;

    // The new request items to create
    private _newRequestItems: {
        ContainsAppCatalog?: IRequest;
        CustomScriptsEnabled?: IRequest;
        IncreaseStorage?: IRequest;
        LockState?: IRequest;
        TeamsConnected?: IRequest;
    } = {};

    // The new values requested
    private _newSiteValues: {
        SensitivityLabel?: string;
    } = {};

    // Constructor
    constructor(el: HTMLElement, disableProps: string[] = []) {
        this._disableProps = disableProps;
        this._el = el;

        // Set the current values
        this._currValues = {
            ContainsAppCatalog: false,
            CustomScriptsEnabled: Helper.hasPermissions(DataSource.Site.RootWeb.EffectiveBasePermissions, SPTypes.BasePermissionTypes.AddAndCustomizePages),
            IncreaseStorage: false,
            LockState: DataSource.Site.ReadOnly && DataSource.Site.WriteLocked ? "ReadOnly" : "Unlock",
            SensitivityLabel: DataSource.Site.SensitivityLabel == "00000000-0000-0000-0000-000000000000" ? "" : DataSource.Site.SensitivityLabel,
            TeamsConnected: DataSource.Site?.GroupId != "00000000-0000-0000-0000-000000000000" && DataSource.Web.AllProperties["TeamifyHidden"] != "TRUE",
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
                    name: "CustomScriptsEnabled",
                    label: "Custom Scripts Enabled:",
                    description: "Enables the custom scripts feature for this site collection.",
                    isDisabled: this._disableProps.indexOf("CustomScriptsEnabled") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.CustomScriptsEnabled,
                    onChange: item => {
                        let value = item ? true : false;

                        // See if we are changing the value
                        if (this._currValues.CustomScriptsEnabled != value) {
                            // Set the value
                            this._newRequestItems.CustomScriptsEnabled = {
                                key: RequestTypes.CustomScript,
                                message: `The request to ${value ? "enable" : "disable"} custom scripts will be processed within 5 minutes.`,
                                value
                            };
                        } else {
                            // Remove the value
                            delete this._newRequestItems.CustomScriptsEnabled;
                        }
                    }
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
                    ],
                    onChange: item => {
                        let value: string = item?.value;

                        // See if we are changing the value
                        if (this._currValues.LockState != value) {
                            // Set the value
                            this._newRequestItems.LockState = {
                                key: RequestTypes.LockState,
                                message: `The request to make the site collection ${value == "NoAccess" ? "have" : "be"} '${item.text}' will be processed within 5 minutes.`,
                                value
                            };
                        } else {
                            // Remove the value
                            delete this._newRequestItems.LockState;
                        }
                    }
                } as Components.IFormControlPropsDropdown,
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
                    },
                    onChange: item => {
                        let value = item ? true : false;

                        // See if we are changing the value
                        if (this._currValues.ContainsAppCatalog != value) {
                            // Set the value
                            this._newRequestItems.ContainsAppCatalog = {
                                key: RequestTypes.AppCatalog,
                                message: `The request to ${value ? "enable" : "disable"} the app catalog will be processed within 5 minutes.`,
                                value
                            };
                        } else {
                            // Remove the value
                            delete this._newRequestItems.ContainsAppCatalog;
                        }
                    }
                } as Components.IFormControlPropsSwitch,
                {
                    name: "IncreaseStorage",
                    label: "Increase Storage:",
                    description: "Enable to increase the site collection storage size.",
                    isDisabled: this._disableProps.indexOf("IncreaseStorage") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.IncreaseStorage,
                    onChange: item => {
                        let value = item ? true : false;

                        // See if we are changing the value
                        if (this._currValues.IncreaseStorage != value) {
                            // Set the value
                            this._newRequestItems.IncreaseStorage = {
                                key: RequestTypes.IncreaseStorage,
                                message: `The request to increase storage for the site collection will be processed within 5 minutes.`,
                                value
                            };
                        } else {
                            // Remove the value
                            delete this._newRequestItems.IncreaseStorage;
                        }
                    }
                } as Components.IFormControlPropsSwitch,
                {
                    name: "TeamsConnected",
                    label: "Teams Connected:",
                    description: "The search scope for the site to target. Default is 'Site'.",
                    isDisabled: this._currValues.TeamsConnected || this._disableProps.indexOf("TeamsConnected") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.TeamsConnected,
                    onChange: item => {
                        let value = item ? true : false;

                        // See if we are changing the value
                        if (this._currValues.TeamsConnected != value) {
                            // Set the value
                            this._newRequestItems.TeamsConnected = {
                                key: RequestTypes.TeamsConnected,
                                message: `The request to connect the site to teams will be processed within 5 minutes.`,
                                value
                            };
                        } else {
                            // Remove the value
                            delete this._newRequestItems.TeamsConnected;
                        }
                    }
                } as Components.IFormControlPropsSwitch,
                {
                    name: "SensitivityLabel",
                    label: "Sensitivity Label",
                    description: "The sensitivity label guid value for this site.",
                    isDisabled: this._disableProps.indexOf("SensitivityLabel") >= 0,
                    type: Components.FormControlTypes.TextField,
                    value: this._currValues.SensitivityLabel,
                    onChange: value => {
                        // See if we are changing the value
                        if (this._currValues.SensitivityLabel != value) {
                            // Set the value
                            this._newSiteValues.SensitivityLabel = value;
                        } else {
                            // Remove the value
                            delete this._newSiteValues.SensitivityLabel;
                        }
                    }
                }
            ]
        });
    }
}