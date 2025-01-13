import { Components, Helper, SPTypes } from "gd-sprest-bs";
import { Tab } from "./base";
import { IProp } from "../app";
import { DataSource, IRequest, RequestTypes } from "../ds";

/**
 * Management Tab
 */
export class ManagementTab extends Tab<{
    ContainsAppCatalog: boolean;
    CustomScriptsEnabled: boolean;
    IncreaseStorage: boolean;
    LockState: string;
    SensitivityLabel: string;
    ShareByEmailEnabled: boolean;
    TeamsConnected: boolean;
}, {
    ContainsAppCatalog: IRequest;
    CustomScriptsEnabled: IRequest;
    IncreaseStorage: IRequest;
    LockState: IRequest;
    TeamsConnected: IRequest;
}, {
    ShareByEmailEnabled: boolean;
}> {
    // Constructor
    constructor(el: HTMLElement, props: { [key: string]: IProp; }) {
        super(el, props, "Site");

        // Set the current values
        this._currValues = {
            ContainsAppCatalog: false,
            CustomScriptsEnabled: Helper.hasPermissions(DataSource.Site.RootWeb.EffectiveBasePermissions, SPTypes.BasePermissionTypes.AddAndCustomizePages),
            IncreaseStorage: false,
            LockState: DataSource.Site.ReadOnly && DataSource.Site.WriteLocked ? "ReadOnly" : "Unlock",
            SensitivityLabel: DataSource.Site.SensitivityLabel == "00000000-0000-0000-0000-000000000000" ? "" : DataSource.Site.SensitivityLabel,
            ShareByEmailEnabled: DataSource.Site.ShareByEmailEnabled,
            TeamsConnected: DataSource.Site?.GroupId != "00000000-0000-0000-0000-000000000000" && DataSource.Web.AllProperties["TeamifyHidden"] != "TRUE",
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
                    name: "CustomScriptsEnabled",
                    label: this._props["CustomScriptsEnabled"].label,
                    description: this._props["CustomScriptsEnabled"].description,
                    isDisabled: this._props["CustomScriptsEnabled"].disabled,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.CustomScriptsEnabled,
                    onChange: item => {
                        let value = item ? true : false;

                        // See if we are changing the value
                        if (this._currValues.CustomScriptsEnabled != value) {
                            // Set the value
                            this._requestItems.CustomScriptsEnabled = {
                                key: RequestTypes.CustomScript,
                                message: `The request to ${value ? "enable" : "disable"} custom scripts will be processed within 5 minutes.`,
                                value
                            };
                        } else {
                            // Remove the value
                            delete this._requestItems.CustomScriptsEnabled;
                        }
                    }
                } as Components.IFormControlPropsSwitch,
                {
                    name: "LockState",
                    label: this._props["LockState"].label,
                    description: this._props["LockState"].description,
                    isDisabled: this._props["LockState"].disabled,
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
                            this._requestItems.LockState = {
                                key: RequestTypes.LockState,
                                message: `The request to make the site collection ${value == "NoAccess" ? "have" : "be"} '${item.text}' will be processed within 5 minutes.`,
                                value
                            };
                        } else {
                            // Remove the value
                            delete this._requestItems.LockState;
                        }
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "ContainsAppCatalog",
                    label: this._props["ContainsAppCatalog"].label,
                    description: this._props["ContainsAppCatalog"].description,
                    isDisabled: this._props["ContainsAppCatalog"].disabled,
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
                            this._requestItems.ContainsAppCatalog = {
                                key: RequestTypes.AppCatalog,
                                message: `The request to ${value ? "enable" : "disable"} the app catalog will be processed within 5 minutes.`,
                                value
                            };
                        } else {
                            // Remove the value
                            delete this._requestItems.ContainsAppCatalog;
                        }
                    }
                } as Components.IFormControlPropsSwitch,
                {
                    name: "IncreaseStorage",
                    label: this._props["IncreaseStorage"].label,
                    description: this._props["IncreaseStorage"].description + ` The current usage is: ${DataSource.formatBytes(DataSource.Site.Usage.Storage)} of ${DataSource.formatBytes(DataSource.Site.Usage.Storage / DataSource.Site.Usage.StoragePercentageUsed)}`,
                    isDisabled: this._props["IncreaseStorage"].disabled,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.IncreaseStorage,
                    onChange: item => {
                        let value = item ? true : false;

                        // See if we are changing the value
                        if (this._currValues.IncreaseStorage != value) {
                            // Set the value
                            this._requestItems.IncreaseStorage = {
                                key: RequestTypes.IncreaseStorage,
                                message: `The request to increase storage for the site collection will be processed within 5 minutes.`,
                                value
                            };
                        } else {
                            // Remove the value
                            delete this._requestItems.IncreaseStorage;
                        }
                    }
                } as Components.IFormControlPropsSwitch,
                {
                    name: "ShareByEmailEnabled",
                    label: this._props["ShareByEmailEnabled"].label,
                    description: this._props["ShareByEmailEnabled"].description,
                    isDisabled: this._props["ShareByEmailEnabled"].disabled,
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
                    name: "TeamsConnected",
                    label: this._props["TeamsConnected"].label,
                    description: this._props["TeamsConnected"].description,
                    isDisabled: this._props["TeamsConnected"].disabled,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.TeamsConnected,
                    onChange: item => {
                        let value = item ? true : false;

                        // See if we are changing the value
                        if (this._currValues.TeamsConnected != value) {
                            // Set the value
                            this._requestItems.TeamsConnected = {
                                key: RequestTypes.TeamsConnected,
                                message: `The request to connect the site to teams will be processed within 5 minutes.`,
                                value
                            };
                        } else {
                            // Remove the value
                            delete this._requestItems.TeamsConnected;
                        }
                    }
                } as Components.IFormControlPropsSwitch
            ]
        });
    }
}