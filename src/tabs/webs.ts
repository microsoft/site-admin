import { Components } from "gd-sprest-bs";
import { Tab } from "./base";
import { DataSource } from "../ds";

/**
 * Webs Tab
 * Displays the root and sub-webs of the site collection.
 */
export class WebsTab extends Tab<{
    CommentsOnSitePagesDisabled: boolean;
    ExcludeFromOfflineClient: boolean;
    SearchScope: number;
    WebTemplate: string;
    WebTitle: string;
}, {}, {
    CommentsOnSitePagesDisabled?: boolean;
    ExcludeFromOfflineClient?: boolean;
    SearchScope?: number;
}> {
    // Constructor
    constructor(el: HTMLElement, disableProps: string[] = []) {
        super(el, disableProps, "Web");

        // Set the current values
        this._currValues = {
            CommentsOnSitePagesDisabled: DataSource.Web.CommentsOnSitePagesDisabled,
            ExcludeFromOfflineClient: DataSource.Web.ExcludeFromOfflineClient,
            SearchScope: DataSource.Web.SearchScope,
            WebTemplate: DataSource.Web.WebTemplate,
            WebTitle: DataSource.Web.Title
        }

        // Render the tab
        this.render();
    }

    // Renders the tab
    private render() {
        Components.Form({
            el: this._el,
            className: "row mt-2",
            groupClassName: "col-4 mb-3",
            onControlRendered: ctrl => {
                // Set the class name
                ctrl.el.classList.add("col-4");
            },
            controls: [
                {
                    name: "WebTemplate",
                    label: "Site Template:",
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
                    name: "ExcludeFromOfflineClient",
                    label: "Remove Site From Search Results:",
                    isDisabled: this._disableProps.indexOf("ExcludeFromOfflineClient") >= 0,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.ExcludeFromOfflineClient,
                    onChange: item => {
                        let value = item ? true : false;

                        // See if we are changing the value
                        if (this._currValues.ExcludeFromOfflineClient != value) {
                            // Set the value
                            this._newValues.ExcludeFromOfflineClient = value;
                        } else {
                            // Remove the value
                            delete this._newValues.ExcludeFromOfflineClient;
                        }
                    }
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
                    ],
                    onChange: item => {
                        let value = item?.data;

                        // See if we are changing the value
                        if (this._currValues.SearchScope != value) {
                            // Set the value
                            this._newValues.SearchScope = value;
                        } else {
                            // Remove the value
                            delete this._newValues.SearchScope;
                        }
                    }
                } as Components.IFormControlPropsDropdown
            ]
        });
    }
}