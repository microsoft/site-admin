import { Components } from "gd-sprest-bs";
import * as moment from "moment";
import { IProp } from "../app";
import { DataSource } from "../ds";
import { Tab } from "./base";
import { M365Groups } from "../m365Groups";

/**
 * Information Tab
 */
export class InfoTab extends Tab {
    // Constructor
    constructor(el: HTMLElement, siteAttestation: boolean, props: { [key: string]: IProp; }) {
        super(el, props, "Site");

        // Render the tab
        this.render(siteAttestation);
    }

    // Returns the site template
    private getSiteTemplate(): string {
        return `${DataSource.Site.RootWeb.WebTemplate}#${DataSource.Site.RootWeb.Configuration}`;
    }

    // Returns the site type, based on the template
    private getSiteType(): string {
        if (DataSource.Site.RootWeb.WebTemplate?.startsWith("GROUP")) { return "Teams"; }
        if (DataSource.Site.RootWeb.WebTemplate?.startsWith("TEAMCHANNEL")) { return "Teams Channel"; }
        return "SharePoint";
    }

    // Renders the tab
    private render(siteAttestation: boolean) {
        let dtAttestation = DataSource.Web.AllProperties["AttestationDate"] || "";
        if (dtAttestation) {
            // Set the date/time
            dtAttestation = moment(dtAttestation).format("LLLL");
        }

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
                    name: "Group",
                    label: this._props["Group"].label,
                    description: this._props["Group"].description,
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Web.AllProperties["GroupId"],
                    onControlRendered: ctrl => {
                        // Get the m365 group
                        if (DataSource.Web.AllProperties["GroupId"]) {
                            // Get the group information
                            M365Groups.getGroupInfo([DataSource.Web.AllProperties["GroupId"]], group => {
                                // Set the group name
                                ctrl.setValue(group.displayName);
                                ctrl.setDescription("The group id: " + group.id);
                            });
                        }
                    }
                },
                {
                    name: "SiteType",
                    label: "Site Type:",
                    type: Components.FormControlTypes.Readonly,
                    value: this.getSiteType()
                },
                {
                    name: "Template",
                    label: this._props["Template"].label,
                    description: this._props["Template"].description,
                    type: Components.FormControlTypes.Readonly,
                    value: this.getSiteTemplate()
                },
                {
                    name: "TemplateName",
                    label: this._props["TemplateName"].label,
                    description: this._props["TemplateName"].description,
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Site.RootWeb.WebTemplate,
                    onControlRendering: ctrl => {
                        // Return a promise
                        return new Promise(resolve => {
                            // Get the web template
                            DataSource.getWebTemplate(this.getSiteTemplate()).then(template => {
                                // Set the value
                                ctrl.value = template;

                                // Resolve the request
                                resolve(ctrl);
                            });
                        });
                    }
                },
                {
                    name: "StorageUsed",
                    label: this._props["StorageUsed"].label,
                    description: this._props["StorageUsed"].description,
                    type: Components.FormControlTypes.Readonly,
                    value: `${DataSource.formatBytes(DataSource.Site.Usage.Storage)} of ${DataSource.formatBytes(DataSource.Site.Usage.Storage / DataSource.Site.Usage.StoragePercentageUsed)} (${Math.round(DataSource.Site.Usage.StoragePercentageUsed * 100) + "%"} Used)`
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
                    name: "AttestationDate",
                    className: siteAttestation ? "" : "d-none",
                    label: this._props["AttestationDate"].label,
                    description: this._props["AttestationDate"].description,
                    type: Components.FormControlTypes.Readonly,
                    value: dtAttestation
                },
                {
                    name: "AttestationUser",
                    className: siteAttestation ? "" : "d-none",
                    label: this._props["AttestationUser"].label,
                    description: this._props["AttestationUser"].description,
                    type: Components.FormControlTypes.Readonly,
                    value: DataSource.Web.AllProperties["AttestationUser"] || ""
                }
            ]
        });
    }
}

/**
 * Mapper for Template Name to Human-Readable Description
 */
const TemplateNameMapper: { [key: string]: string } = {

}