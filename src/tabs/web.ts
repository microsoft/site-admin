import { Components } from "gd-sprest-bs";
import { Tab } from "./base";
import { IProp } from "../app";
import { DataSource, IRequest, RequestTypes } from "../ds";

/**
 * Webs Tab
 * Displays the root and sub-webs of the site collection.
 */
export class WebTab extends Tab<{
    CommentsOnSitePagesDisabled: boolean;
    ExcludeFromOfflineClient: boolean;
    NoCrawl?: boolean;
    SearchScope: number;
    WebTemplate: string;
    WebTitle: string;
}, {
    NoCrawl?: IRequest;
}, {
    CommentsOnSitePagesDisabled?: boolean;
    ExcludeFromOfflineClient?: boolean;
    SearchScope?: number;
}> {
    // Constructor
    constructor(el: HTMLElement, props: { [key: string]: IProp; }) {
        super(el, props, "Web");

        // Set the current values
        this._currValues = {
            CommentsOnSitePagesDisabled: DataSource.Web.CommentsOnSitePagesDisabled,
            ExcludeFromOfflineClient: DataSource.Web.ExcludeFromOfflineClient,
            NoCrawl: DataSource.Web.NoCrawl,
            SearchScope: DataSource.Web.SearchScope,
            WebTemplate: DataSource.Web.WebTemplate,
            WebTitle: DataSource.Web.Title
        }

        // Render the tab
        this.render();
    }


    // Method to refresh the tab
    refresh() {
        // Update the current values
        this._currValues = {
            CommentsOnSitePagesDisabled: DataSource.Web.CommentsOnSitePagesDisabled,
            ExcludeFromOfflineClient: DataSource.Web.ExcludeFromOfflineClient,
            NoCrawl: DataSource.Web.NoCrawl,
            SearchScope: DataSource.Web.SearchScope,
            WebTemplate: DataSource.Web.WebTemplate,
            WebTitle: DataSource.Web.Title
        }

        // Clear the changes
        this._newValues = {};
        this._requestItems = {};

        // Clear the tab
        while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

        // Render the tab
        this.render();
    }

    // Renders the tab
    private render() {
        Components.Form({
            el: this._el,
            className: "row",
            groupClassName: "col-4 mb-3",
            onControlRendered: ctrl => {
                // Set the class name
                ctrl.el.classList.add("col-4");
            },
            controls: [
                {
                    name: "Template",
                    label: this._props["Template"].label,
                    description: this._props["Template"].description,
                    isDisabled: this._props["Template"].disabled,
                    type: Components.FormControlTypes.Readonly,
                    value: this._currValues.WebTemplate,
                    onControlRendering: ctrl => {
                        // Return a promise
                        return new Promise(resolve => {
                            // Get the web template
                            DataSource.getWebTemplate(this._currValues.WebTemplate).then(template => {
                                // Set the value
                                ctrl.value = template;

                                // Resolve the request
                                resolve(ctrl);
                            });
                        });
                    }
                },
                {
                    name: "Title",
                    label: this._props["Title"].label,
                    description: this._props["Title"].description,
                    isDisabled: this._props["Title"].disabled,
                    type: Components.FormControlTypes.Readonly,
                    value: this._currValues.WebTitle
                },
                {
                    name: "CommentsOnSitePagesDisabled",
                    label: this._props["CommentsOnSitePagesDisabled"].label,
                    description: this._props["CommentsOnSitePagesDisabled"].description,
                    isDisabled: this._props["CommentsOnSitePagesDisabled"].disabled,
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
                    label: this._props["ExcludeFromOfflineClient"].label,
                    description: this._props["ExcludeFromOfflineClient"].description,
                    isDisabled: this._props["ExcludeFromOfflineClient"].disabled,
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
                    name: "NoCrawl",
                    label: this._props["NoCrawl"].label,
                    description: this._props["NoCrawl"].description,
                    isDisabled: this._props["NoCrawl"].disabled,
                    type: Components.FormControlTypes.Switch,
                    value: this._currValues.NoCrawl,
                    onChange: item => {
                        let value = item ? true : false;

                        // See if we are changing the value
                        if (this._currValues.NoCrawl != value) {
                            // Set the value
                            this._requestItems.NoCrawl = {
                                key: RequestTypes.NoCrawl,
                                message: `The request to ${value ? "hide" : "show"} content from search will be processed within 5 minutes.`,
                                value
                            };
                        } else {
                            // Remove the value
                            delete this._requestItems.NoCrawl;
                        }
                    }
                } as Components.IFormControlPropsSwitch,
                {
                    name: "SearchScope",
                    label: this._props["SearchScope"].label,
                    description: this._props["SearchScope"].description,
                    isDisabled: this._props["SearchScope"].disabled,
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