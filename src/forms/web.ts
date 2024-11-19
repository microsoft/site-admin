import { LoadingDialog } from "dattatable";
import { Components, Helper, Types } from "gd-sprest-bs";
import { DataSource } from "../ds";

/**
 * Web Form
 */
export class Web {
    private _apiUrls: string[] = null;
    private _el: HTMLElement = null;
    private _form: Components.IForm = null;
    private _disableProps: string[] = null;
    private _web: Types.SP.WebOData = null;

    // The current values
    private _currValues: {
        CommentsOnSitePagesDisabled: boolean;
        DesignatedMAJCOM: string;
        ExcludeFromOfflineClient: boolean;
        SearchScope: number;
        WebTemplate: string;
    } = null;

    constructor(web: Types.SP.WebOData, el: HTMLElement, disableProps: string[], apiUrls?: string[]) {
        // Save the properties
        this._apiUrls = apiUrls;
        this._el = el;
        this._disableProps = disableProps;
        this._web = web;

        // Set the current values
        this._currValues = {
            CommentsOnSitePagesDisabled: this._web.CommentsOnSitePagesDisabled,
            DesignatedMAJCOM: this._web.AllProperties["DesignatedMAJCOM"],
            ExcludeFromOfflineClient: this._web.ExcludeFromOfflineClient,
            SearchScope: this._web.SearchScope,
            WebTemplate: this._web.WebTemplate
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

                    // Save the properties
                    this.save(values).then(() => {
                        // See if we are updating the property bag
                        if (this._currValues.DesignatedMAJCOM != values["DesignatedMAJCOM"]) {
                            // Show a loading dialog
                            LoadingDialog.setHeader("Updating Site Property");
                            LoadingDialog.setBody("This will close after the update completes.");
                            LoadingDialog.show();

                            // Update the property
                            Helper.setWebProperty("DesignatedMAJCOM", values["DesignatedMAJCOM"], true, this._web.Url).then(() => {
                                // Update the current value
                                this._currValues.DesignatedMAJCOM = values["DesignatedMAJCOM"];

                                // Hide the dialog
                                LoadingDialog.hide();
                            });
                        }
                    });
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
                    name: "WebTemplate",
                    label: "Web Template:",
                    description: "The type of web.",
                    type: Components.FormControlTypes.Readonly,
                    value: this._currValues.WebTemplate
                },
                {
                    name: "CommentsOnSitePagesDisabled",
                    label: "Comments On Site Pages Disabled:",
                    description: "If true, it will hide the comments on the site pages.",
                    isDisabled: this._disableProps.indexOf("CommentsOnSitePagesDisabled") >= 0,
                    type: Components.FormControlTypes.Dropdown,
                    value: this._currValues.CommentsOnSitePagesDisabled ? "true" : "false",
                    items: [
                        {
                            text: "true",
                            data: true
                        },
                        {
                            text: "false",
                            data: false
                        }
                    ]
                } as Components.IFormControlPropsDropdown,
                {
                    name: "ExcludeFromOfflineClient",
                    label: "Exclude From Offline Client:",
                    description: "Disables the offline sync feature in all libraries.",
                    isDisabled: this._disableProps.indexOf("ExcludeFromOfflineClient") >= 0,
                    type: Components.FormControlTypes.Dropdown,
                    value: this._currValues.ExcludeFromOfflineClient ? "true" : "false",
                    items: [
                        {
                            text: "true",
                            data: true
                        },
                        {
                            text: "false",
                            data: false
                        }
                    ]
                } as Components.IFormControlPropsDropdown,
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
                    name: "DesignatedMAJCOM",
                    label: "Designated MAJCOM",
                    description: "The designated MAJCOM for this site.",
                    isDisabled: this._disableProps.indexOf("DesignatedMAJCOM") >= 0,
                    type: Components.FormControlTypes.TextField,
                    value: this._currValues.DesignatedMAJCOM
                }
            ]
        });
    }

    // Renders the header
    private renderHeader() {
        // Render the header
        Components.Jumbotron({
            el: this._el,
            className: "mb-2",
            lead: "Site Settings",
            size: Components.JumbotronSize.Small,
            type: Components.JumbotronTypes.Primary
        });
    }

    // Saves the properties
    private save(values): PromiseLike<void> {
        return new Promise((resolve, reject) => {
            let apis: { api: string; key: string; value: any; }[] = [];
            let props = {};
            let updateFl = false;

            // Parse the keys
            for (let key in this._currValues) {
                let value = values[key].data;
                if (this._currValues[key] != value) {
                    // See if there is an associated api
                    let keyIdx = this._apiUrls.indexOf(key);
                    if (keyIdx >= 0) {
                        // Append the url
                        apis.push({ key, value, api: this._apiUrls[keyIdx] });
                    } else {
                        // Add the property
                        props[key] = values[key].data;

                        // Set the flag
                        updateFl = true;
                    }
                }
            }

            // See if an update is needed
            if (this._currValues.CommentsOnSitePagesDisabled != values["CommentsOnSitePagesDisabled"].data) { props["CommentsOnSitePagesDisabled"] = values["CommentsOnSitePagesDisabled"].data; updateFl = true; }
            if (this._currValues.ExcludeFromOfflineClient != values["ExcludeFromOfflineClient"].data) { props["ExcludeFromOfflineClient"] = values["ExcludeFromOfflineClient"].data; updateFl = true; }
            if (this._currValues.SearchScope != values["SearchScope"].data) { props["SearchScope"] = values["SearchScope"].data; updateFl = true; }

            // Process the requests
            DataSource.processRequests(apis).then(() => {
                // See if an update is needed
                if (updateFl) {
                    // Show a loading dialog
                    LoadingDialog.setHeader("Updating Site");
                    LoadingDialog.setBody("This will close after the changes complete.");
                    LoadingDialog.show();

                    // Update the web
                    this._web.update(props).execute((responses) => {
                        // Update the current values
                        this._currValues.CommentsOnSitePagesDisabled = values["CommentsOnSitePagesDisabled"].data;
                        this._currValues.DesignatedMAJCOM = values["DesignatedMAJCOM"];
                        this._currValues.ExcludeFromOfflineClient = values["ExcludeFromOfflineClient"].data;
                        this._currValues.SearchScope = values["SearchScope"].data;

                        // Close the dialog
                        LoadingDialog.hide();

                        // Resolve the request
                        resolve();
                    }, reject);
                } else {
                    // Resolve the request
                    resolve();
                }
            });
        });
    }
}