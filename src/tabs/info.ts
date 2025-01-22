import { Components } from "gd-sprest-bs";
import * as moment from "moment";
import { IProp } from "../app";
import { DataSource } from "../ds";
import { Tab } from "./base";

/**
 * Information Tab
 */
export class InfoTab extends Tab {
    // Constructor
    constructor(el: HTMLElement, props: { [key: string]: IProp; }) {
        super(el, props, "Site");

        // Render the tab
        this.render();
    }

    // Renders the tab
    private render() {
        // Render the form
        Components.Form({
            el: this._el,
            className: "row",
            groupClassName: "col-4 mb-3",
            controls: [
                {
                    name: "Created",
                    label: this._props["Created"].label,
                    description: this._props["Created"].description,
                    type: Components.FormControlTypes.Readonly,
                    value: moment(DataSource.Site.RootWeb.Created).format("LLLL")
                },
                {
                    name: "Title",
                    label: this._props["Title"].label,
                    description: this._props["Title"].description,
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Site.RootWeb.Title
                },
                {
                    name: "Template",
                    label: this._props["Template"].label,
                    description: this._props["Template"].description,
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Site.RootWeb.WebTemplate,
                    onControlRendering: ctrl => {
                        // Return a promise
                        return new Promise(resolve => {
                            // Get the web template
                            DataSource.getWebTemplate(DataSource.Site.RootWeb.WebTemplate).then(template => {
                                // Set the value
                                ctrl.value = template;

                                // Resolve the request
                                resolve(ctrl);
                            });
                        });
                    }
                },
                {
                    name: "HubSite",
                    label: this._props["HubSite"].label,
                    description: this._props["HubSite"].description,
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Site.IsHubSite ? "Yes" : "No"
                },
                {
                    name: "HubSiteConnected",
                    label: this._props["HubSiteConnected"].label,
                    description: this._props["HubSiteConnected"].description,
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Site.HubSiteId != "00000000-0000-0000-0000-000000000000" ? "Yes" : "No"
                },
                {
                    name: "StorageUsed",
                    label: this._props["StorageUsed"].label,
                    description: this._props["StorageUsed"].description,
                    type: Components.FormControlTypes.Readonly,
                    value: `${DataSource.formatBytes(DataSource.Site.Usage.Storage)} of ${DataSource.formatBytes(DataSource.Site.Usage.Storage / DataSource.Site.Usage.StoragePercentageUsed)} (${Math.round(DataSource.Site.Usage.StoragePercentageUsed * 100) / 100 + "%"} Used)`
                }
            ]
        });
    }
}