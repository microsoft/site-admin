import { Components } from "gd-sprest-bs";
import * as moment from "moment";
import { Tab } from "./base";
import { DataSource } from "../ds";

/**
 * Information Tab
 */
export class InfoTab extends Tab {
    // Constructor
    constructor(el: HTMLElement, disabledProps:string[]) {
        super(el, disabledProps, "Site");

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
                    name: "Created",
                    label: "Date Created:",
                    type: Components.FormControlTypes.Readonly,
                    value: moment(DataSource.Site.RootWeb.Created).format("LLLL")
                },
                {
                    name: "Title",
                    label: "Title:",
                    description: "The title of the site collection.",
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Site.RootWeb.Title
                },
                {
                    name: "SiteTemplate",
                    label: "Site Template:",
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Site.RootWeb.WebTemplate
                },
                {
                    name: "IsHubSite",
                    label: "Hub Site:",
                    description: "If true, indicates this is a hub site.",
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Site.IsHubSite ? "Yes" : "No"
                },
                {
                    name: "IsHubSiteConnected",
                    label: "Connected to Hub:",
                    description: "If true, indicates this is connected to a hub site.",
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Site.HubSiteId != "00000000-0000-0000-0000-000000000000" ? "Yes" : "No"
                },
                {
                    name: "UsageSize",
                    label: "Storage Used:",
                    type: Components.FormControlTypes.Readonly,
                    value: `${DataSource.formatBytes(DataSource.Site.Usage.Storage)} of ${DataSource.formatBytes(DataSource.Site.Usage.Storage/DataSource.Site.Usage.StoragePercentageUsed)} (${Math.round(DataSource.Site.Usage.StoragePercentageUsed * 100) / 100 + "%"} Used)`
                }
            ]
        });
    }
}