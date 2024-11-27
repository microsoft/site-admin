import { LoadingDialog } from "dattatable";
import { Components, Helper } from "gd-sprest-bs";
import { DataSource, IResponse } from "../ds";
import { APIResponseModal } from "./response";

export interface ISearchProp {
    description: string;
    key: string;
    label: string;
}

/**
 * Web Form
 */
export class Web {
    private _el: HTMLElement = null;
    private _form: Components.IForm = null;
    private _disableProps: string[] = null;
    private _searchProp: ISearchProp;

    // The current values
    private _currValues: {
        CommentsOnSitePagesDisabled: boolean;
        ExcludeFromOfflineClient: boolean;
        SearchProp: string;
        SearchScope: number;
        WebTemplate: string;
        WebTitle: string;
    } = null;

    constructor(el: HTMLElement, disableProps: string[] = [], searchProp: ISearchProp = {} as any) {
        // Save the properties
        this._el = el;
        this._disableProps = disableProps;
        this._searchProp = searchProp;

        // Set the current values
        this._currValues = {
            CommentsOnSitePagesDisabled: DataSource.Web.CommentsOnSitePagesDisabled,
            ExcludeFromOfflineClient: DataSource.Web.ExcludeFromOfflineClient,
            SearchProp: DataSource.Web.AllProperties[this._searchProp.key],
            SearchScope: DataSource.Web.SearchScope,
            WebTemplate: DataSource.Web.WebTemplate,
            WebTitle: DataSource.Web.Title
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

                    // Save the properties
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
            onControlRendered: ctrl => {
                // Set the class name
                ctrl.el.classList.add("col-4");
            },
            controls: [
                {
                    name: "WebTemplate",
                    label: "Web Template:",
                    description: "The type of web.",
                    type: Components.FormControlTypes.Readonly,
                    value: this._currValues.WebTemplate
                },
                {
                    name: "WebTitle",
                    label: "Web Title:",
                    description: "The title of web.",
                    type: Components.FormControlTypes.Readonly,
                    value: this._currValues.WebTitle
                },
                {
                    name: "CommentsOnSitePagesDisabled",
                    label: "Comments On Site Pages Disabled:",
                    description: "If true, it will hide the comments on the site pages.",
                    isDisabled: this._disableProps.indexOf("CommentsOnSitePagesDisabled") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.CommentsOnSitePagesDisabled
                } as Components.IFormControlPropsSwitch,
                {
                    name: "ExcludeFromOfflineClient",
                    label: "Exclude From Offline Client:",
                    description: "Disables the offline sync feature in all libraries.",
                    isDisabled: this._disableProps.indexOf("ExcludeFromOfflineClient") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.ExcludeFromOfflineClient
                } as Components.IFormControlPropsSwitch,
                {
                    name: "SearchScope",
                    label: "Search Scope:",
                    description: "The search scope for the site to target. Default is 'Site'.",
                    isDisabled: this._disableProps.indexOf("SearchScope") >= 0,
                    type: Components.FormControlTypes.Dropdown,
                    value: this._currValues.SearchScope,
                    items: [
                        {
                            text: "Default",
                            data: 0,
                            value: "0"
                        },
                        {
                            text: "Tenant",
                            data: 1,
                            value: "1"
                        },
                        {
                            text: "Hub",
                            data: 2,
                            value: "2"
                        },
                        {
                            text: "Site",
                            data: 3,
                            value: "3"
                        }
                    ]
                } as Components.IFormControlPropsDropdown,
                {
                    className: this._searchProp.key ? "" : "d-none",
                    name: "SearchProp",
                    label: this._searchProp.label || "Search Property",
                    description: this._searchProp.description || "The custom property to set for search.",
                    isDisabled: this._disableProps.indexOf("SearchProp") >= 0,
                    type: Components.FormControlTypes.TextField,
                    value: this._currValues.SearchProp
                }
            ]
        });
    }

    // Saves the properties
    private save(values): PromiseLike<void> {
        return new Promise((resolve, reject) => {
            let props = {};
            let responses: IResponse[] = [];
            let updateFlags: { [key: string]: boolean } = {};

            // Parse the keys
            for (let key in this._currValues) {
                // Skip disabled and read-only controls
                let ctrl = this._form.getControl(key);
                if (ctrl.props.isDisabled || ctrl.props.isReadonly) { continue; }

                // Set the flag
                updateFlags[key] = false;

                // See if an update is needed
                let value = typeof (values[key]) === "boolean" ? values[key] : values[key].data || values[key].value;
                if (this._currValues[key] != value) {
                    // Add the property
                    props[key] = value;

                    // Set the flag
                    updateFlags[key] = true;
                }
            }

            // Update the search property
            this.updateSearchProp(values["SearchProp"]).then(() => {
                // See if an update is needed
                if (updateFlags.CommentsOnSitePagesDisabled || updateFlags.ExcludeFromOfflineClient || updateFlags.SearchScope) {
                    // Show a loading dialog
                    LoadingDialog.setHeader("Updating Site");
                    LoadingDialog.setBody("This will close after the changes complete.");
                    LoadingDialog.show();

                    // Update the web
                    DataSource.Web.update(props).execute(() => {
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
                            });
                        }

                        // Close the dialog
                        LoadingDialog.hide();

                        // Show the responses
                        new APIResponseModal(responses);

                        // Resolve the request
                        resolve();
                    }, reject);
                } else {
                    // Show the responses
                    new APIResponseModal(responses);

                    // Resolve the request
                    resolve();
                }
            });
        });
    }

    // Method to update the search property
    private updateSearchProp(value: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Ensure a property is set and an update is required
            if (this._searchProp.key && this._currValues.SearchProp != value) {
                // Show a loading dialog
                LoadingDialog.setHeader("Updating Site Property");
                LoadingDialog.setBody("This will close after the update completes.");
                LoadingDialog.show();

                // Update the property
                Helper.setWebProperty(this._searchProp.key, value, true, DataSource.Web.Url).then(() => {
                    // Update the current value
                    this._currValues.SearchProp = value;

                    // Hide the dialog
                    LoadingDialog.hide();

                    // Resolve the request
                    resolve();
                }, resolve);
            } else {
                // Resolve the request
                resolve();
            }
        });
    }
}