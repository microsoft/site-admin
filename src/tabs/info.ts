import { Components } from "gd-sprest-bs";
import * as moment from "moment";
import { ITab } from "./tab.d";
import { DataSource, IRequest } from "../ds";

/**
 * Information Tab
 */
export class InfoTab implements ITab {
    private _el: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement) {
        this._el = el;

        // Render the tab
        this.render();
    }

    // This tab only holds read-only properties, so no changes will exist
    getProps(): { [key: string]: string | number | boolean; } { return {}; }

    // This tab only holds read-only properties, so no changes will exist
    getRequests(): IRequest[] { return []; }

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
                    label: "Created:",
                    description: "The date the site was created.",
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
                    description: "The site template of the root web.",
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Site.RootWeb.WebTemplate
                },
                {
                    name: "IsHubSite",
                    label: "Is Hub Site:",
                    description: "If true, indicates this is a hub site.",
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Site.IsHubSite ? "Yes" : "No"
                },
                {
                    name: "IsHubSiteConnected",
                    label: "Is Connected to Hub Site:",
                    description: "If true, indicates this is connected to a hub site.",
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Site.HubSiteId != "00000000-0000-0000-0000-000000000000" ? "Yes" : "No"
                },
                {
                    name: "UsageSize",
                    label: "Storage Used:",
                    description: "The current storage of the site collection.",
                    type: Components.FormControlTypes.Readonly,
                    value: `${DataSource.formatBytes(DataSource.Site.Usage.Storage)} of ${DataSource.formatBytes(DataSource.Site.Usage.Storage/DataSource.Site.Usage.StoragePercentageUsed)} (${Math.round(DataSource.Site.Usage.StoragePercentageUsed * 100) / 100 + "%"} Used)`
                }
            ]
        });
    }
}