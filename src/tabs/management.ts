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
    ShareByEmailEnabled: boolean;
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
    constructor(el: HTMLElement, props: { [key: string]: IProp; }, maxStorageSize: number) {
        super(el, props, "Site");

        // Set the current values
        this._currValues = {
            ContainsAppCatalog: DataSource.hasAppCatalog(),
            CustomScriptsEnabled: Helper.hasPermissions(DataSource.Site.RootWeb.EffectiveBasePermissions, SPTypes.BasePermissionTypes.AddAndCustomizePages),
            IncreaseStorage: false,
            LockState: DataSource.Site.ReadOnly && DataSource.Site.WriteLocked ? "ReadOnly" : "Unlock",
            ShareByEmailEnabled: DataSource.Site.ShareByEmailEnabled
        }

        // Render the tab
        this.render(maxStorageSize);
    }

    // Renders the tab
    private render(maxStorageSize: number = 0) {
        // Render the form
        Components.Form({
            el: this._el,
            className: "row",
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
                    value: this._currValues.ContainsAppCatalog,
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
                    onControlRendering: ctrl => {
                        // See if the max threshold has been reached
                        // 1 TB = 1024GB = 1024*1024MB = 1024*1024*1024KB = 1024*1024*1024*1024 = 1099511627776
                        if (Math.round(DataSource.Site.Usage.Storage / DataSource.Site.Usage.StoragePercentageUsed) >= maxStorageSize * 1099511627776) {
                            // Update the props
                            ctrl.isDisabled = true;
                            ctrl.description = "The max threshold has been met. Please contact the helpdesk to request an increase.";
                        }

                        // Return the control
                        return ctrl;
                    },
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
                } as Components.IFormControlPropsSwitch
            ]
        });
    }
}