import { LoadingDialog } from "dattatable";
import { Components, Helper, Types } from "gd-sprest-bs";

/**
 * Web Form
 */
export class Web {
    private _el: HTMLElement = null;
    private _form: Components.IForm = null;
    private _web: Types.SP.WebOData = null;

    // The current values
    private _currValues: {
        CommentsOnSitePagesDisabled: boolean;
        DesignatedMAJCOM: string;
        ExcludeFromOfflineClient: boolean;
        SearchScope: number;
        WebTemplate: string;
    } = null;

    constructor(web: Types.SP.WebOData, el: HTMLElement) {
        // Save the properties
        this._el = el;
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
                    let props = {};
                    let updateFl = false;
                    let values = this._form.getValues();

                    // See if an update is needed
                    if (this._currValues.CommentsOnSitePagesDisabled != values["CommentsOnSitePagesDisabled"]) { props["CommentsOnSitePagesDisabled"] = values["CommentsOnSitePagesDisabled"]; updateFl = true; }
                    if (this._currValues.ExcludeFromOfflineClient != values["ExcludeFromOfflineClient"]) { props["ExcludeFromOfflineClient"] = values["ExcludeFromOfflineClient"]; updateFl = true; }
                    if (this._currValues.SearchScope != values["SearchScope"]) { props["SearchScope"] = values["SearchScope"]; updateFl = true; }

                    // See if we are updating the property bag
                    if (this._currValues.DesignatedMAJCOM != values["DesignatedMAJCOM"]) {
                        // Update the property
                        Helper.setWebProperty("DesignatedMAJCOM", values["DesignatedMAJCOM"], true);
                    }

                    // See if an update is needed
                    if (updateFl) {
                        // Show a loading dialog
                        LoadingDialog.setHeader("Updating Web");
                        LoadingDialog.setBody("This will close after the changes complete.");
                        LoadingDialog.show();

                        // Update the web
                        this._web.update(props).execute(() => {
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
                    name: "WebTemplate",
                    label: "Web Template:",
                    description: "The type of web.",
                    type: Components.FormControlTypes.Readonly
                },
                {
                    name: "CommentsOnSitePagesDisabled",
                    label: "Comments On Site Pages Disabled:",
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
                    name: "ExcludeFromOfflineClient",
                    label: "Exclude From Offline Client:",
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
                    name: "SearchScope",
                    label: "Search Scope:",
                    description: "The search scope for the site to target. Default is 'Site'.",
                    type: Components.FormControlTypes.Dropdown,
                    items: [
                        {
                            text: "Default",
                            value: "0"
                        },
                        {
                            text: "Tenant",
                            value: "1"
                        },
                        {
                            text: "Hub",
                            value: "2"
                        },
                        {
                            text: "Site",
                            value: "3"
                        }
                    ]
                } as Components.IFormControlPropsDropdown,
                {
                    name: "DesignatedMAJCOM",
                    label: "Designated MAJCOM",
                    description: "The designated MAJCOM for this site.",
                    type: Components.FormControlTypes.TextField
                }
            ]
        });
    }

    // Renders the header
    private renderHeader() {
        // Render the header
        Components.Jumbotron({
            el: this._el,
            lead: "Site Settings",
            size: Components.JumbotronSize.Small,
            type: Components["JumbotronTypes"].Primary
        });
    }
}