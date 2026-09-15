import { CanvasForm } from "dattatable";
import { Components, Helper } from "gd-sprest-bs";
import * as moment from "moment";
import { IAppProps } from "../app";
import * as Reports from "../reports";
import { ReportTypes } from "./reports";

/**
 * Site Audit
 */
export class SiteAudit {
    private _appProps: IAppProps = null;
    private _el: HTMLElement = null;
    private _elLog: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement, appProps: IAppProps) {
        this._appProps = appProps;
        this._el = el;

        // Render the tab
        this.render();
    }

    // Returns the reports to view
    private getReports(): Components.ICheckboxGroupItem[] {
        // Parse the report types
        let items: Components.ICheckboxGroupItem[] = [];

        // Add the reports
        if (typeof (this._appProps.hideReports.dlp) === "undefined" || this._appProps.hideReports.dlp != true) {
            items.push({
                name: ReportTypes.DLP,
                label: this._appProps.reportNames.dlp || "Data Loss Prevention",
                data: "Finds files that has DLP applied to it."
            });
        }
        if (typeof (this._appProps.hideReports.docRetention) === "undefined" || this._appProps.hideReports.docRetention != true) {
            items.push({
                name: ReportTypes.DocRetention,
                label: this._appProps.reportNames.docRetention || "Document Retention",
                data: "Find documents older than a specified date."
            });
        }
        if (typeof (this._appProps.hideReports.externalShares) === "undefined" || this._appProps.hideReports.externalShares != true) {
            items.push({
                name: ReportTypes.ExternalShares,
                label: this._appProps.reportNames.externalShares || "External Shares",
                data: "Scans for documents that have been shared externally."
            });
        }
        if (typeof (this._appProps.hideReports.externalUsers) === "undefined" || this._appProps.hideReports.externalUsers != true) {
            items.push({
                name: ReportTypes.ExternalUsers,
                label: this._appProps.reportNames.externalUsers || "External Users",
                data: "Scans the user information list for 'external' user accounts."
            });
        }
        if (typeof (this._appProps.hideReports.permissions) === "undefined" || this._appProps.hideReports.permissions != true) {
            items.push({
                name: ReportTypes.Permissions,
                label: this._appProps.reportNames.permissions || "Permissions",
                data: "Scans all users/groups that have permissions to the site."
            });
        }
        if (typeof (this._appProps.hideReports.searchAgents) === "undefined" || this._appProps.hideReports.searchAgents != true) {
            items.push({
                name: ReportTypes.SearchAgents,
                label: this._appProps.reportNames.searchAgents || "Search Agents",
                data: "Searches all libraries for agent files in the site."
            });
        }
        if (typeof (this._appProps.hideReports.searchDocs) === "undefined" || this._appProps.hideReports.searchDocs != true) {
            items.push({
                name: ReportTypes.SearchDocs,
                label: this._appProps.reportNames.searchDocs || "Search Documents",
                data: "Find documents by keywords."
            });
        }
        if (typeof (this._appProps.hideReports.searchEEEU) === "undefined" || this._appProps.hideReports.searchEEEU != true) {
            items.push({
                name: ReportTypes.SearchEEEU,
                label: this._appProps.reportNames.searchEEEU || "Search EEEU",
                data: "Search for the 'Every' and 'Everyone exception external users' accounts."
            });
        }
        if (typeof (this._appProps.hideReports.sensitivityLabels) === "undefined" || this._appProps.hideReports.sensitivityLabels != true) {
            items.push({
                name: ReportTypes.SensitivityLabels,
                label: this._appProps.reportNames.sensitivityLabels || "Sensitivity Labels",
                data: "Search for files that have sensitivity labels."
            });
        }
        if (typeof (this._appProps.hideReports.sharingLinks) === "undefined" || this._appProps.hideReports.sharingLinks != true) {
            items.push({
                name: ReportTypes.SharingLinks,
                label: this._appProps.reportNames.sharingLinks || "Sharing Links",
                data: "Scans for any 'Sharing Link' groups."
            });
        }
        if (typeof (this._appProps.hideReports.uniquePermissions) === "undefined" || this._appProps.hideReports.uniquePermissions != true) {
            items.push({
                name: ReportTypes.UniquePermissions,
                label: this._appProps.reportNames.uniquePermissions || "Unique Permissions",
                data: "Scans for items that have unique permissions."
            });
        }

        // Return the items
        return items;
    }

    // Renders the audit form
    private render() {
        // Clear the element
        while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

        // Render the form
        let form = Components.Form({
            el: this._el,
            controls: [
                {
                    name: "Reports",
                    label: "Reports",
                    description: "Select the reports to run in this report.",
                    type: Components.FormControlTypes.MultiSwitch,
                    items: this.getReports(),
                    required: true,
                    //value: this._currValues.WebTemplate
                } as Components.IFormControlPropsMultiSwitch,
                {
                    name: "SkipLargeLists",
                    label: "Skip Large Lists",
                    description: "Skips lists/libraries that have more than the items specified.",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    value: 5000,
                    items: [
                        { text: "All Lists", value: "0", isSelected: true },
                        { text: ">500 Items", value: "500" },
                        { text: ">1000 Items", value: "1000" },
                        { text: ">5000 Items", value: "5000" },
                        { text: ">10000 Items", value: "10000" },
                        { text: ">25000 Items", value: "25000" },
                        { text: ">50000 Items", value: "50000" },
                        { text: ">100000 Items", value: "100000" },
                    ]
                } as Components.IFormControlPropsDropdown
            ]
        });

        // Render a footer
        let elFooter = document.createElement("div");
        elFooter.classList.add("mt-2");
        elFooter.classList.add("d-flex");
        elFooter.classList.add("justify-content-end");
        this._el.appendChild(elFooter);

        // Add a button
        Components.Button({
            el: elFooter,
            text: "Run",
            onClick: () => {
                // Ensure the form is required
                if (!form.isValid()) { return; }

                // Run the site audit reports against the sites
                // Set the default values to use
                let formValues = form.getValues();
                formValues["ShowSearch"] = false;
                this.run(formValues["Reports"]);
            }
        });
    }

    // Runs the site audit
    private run(reports: Components.ICheckboxGroupItem[]) {
        // Clear the element
        while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

        // Render a navigation
        Components.Navbar({
            el: this._el,
            brand: "Site Audit",
            items: [{
                text: "New Audit",
                className: "btn-outline-light",
                isButton: true,
                onClick: () => {
                    // Render the site audit
                    this.render();
                }
            }]
        });

        // Generate the tabs
        let elTabs: { [key: string]: HTMLElement } = {};
        let items: Components.INavLinkProps[] = [];
        reports.forEach(report => {
            // Add the tab
            items.push({
                data: report.name,
                title: report.label,
                isDisabled: true,
                onRenderTab: (el, item) => { elTabs[item.data] = el; }
            });
        });

        // Set the first tab to be active
        items[0].isActive = true;
        items[0].isDisabled = false;

        // Render tabs
        let nav = Components.Nav({
            el: this._el,
            isPills: true,
            isTabs: true,
            items
        });

        // Parse the reports
        Helper.Executor(reports, report => {
            // Return a promise
            return new Promise(resolve => {
                // Set the default form values
                let formValues = {};

                // Ensure this tab is enabled and show it
                nav.getTab(report.label).enable();
                nav.showTab(report.name);

                // See which report we are running
                switch (report.name) {
                    case ReportTypes.DLP:
                        formValues["FileTypes"] = this._appProps.reportProps.dlpFileExt;
                        return Reports.DLP.run(elTabs[report.name], this._appProps.auditOnly, formValues, null, () => { resolve(null); });
                    case ReportTypes.DocRetention:
                        formValues["SelectedDate"] = moment(Date.now()).subtract(this._appProps.reportProps.docRententionYears, "months").toISOString();
                        return Reports.DocRetention.run(elTabs[report.name], this._appProps.auditOnly, formValues, null, () => { resolve(null); });
                    case ReportTypes.ExternalShares:
                        return Reports.ExternalShares.run(elTabs[report.name], this._appProps.auditOnly, formValues, null, () => { resolve(null); });
                    case ReportTypes.ExternalUsers:
                        return Reports.ExternalUsers.run(elTabs[report.name], this._appProps.auditOnly, formValues, null, () => { resolve(null); });
                    case ReportTypes.Permissions:
                        return Reports.Permissions.run(elTabs[report.name], this._appProps.auditOnly, formValues, null, () => { resolve(null); });
                    case ReportTypes.SearchAgents:
                        return Reports.SearchAgents.run(elTabs[report.name], this._appProps.auditOnly, formValues, null, () => { resolve(null); });
                    case ReportTypes.SearchDocs:
                        return Reports.SearchDocs.run(elTabs[report.name], this._appProps.auditOnly, formValues, null, () => { resolve(null); });
                    case ReportTypes.SearchEEEU:
                        return Reports.SearchEEEU.run(elTabs[report.name], this._appProps.auditOnly, formValues, null, () => { resolve(null); });
                    case ReportTypes.SensitivityLabels:
                        formValues["SearchType"] = [{ name: "WithLabels" }, { name: "WithoutLabels" }];
                        return Reports.SensitivityLabels.run(elTabs[report.name], this._appProps.auditOnly, formValues, null, () => { resolve(null); });
                    case ReportTypes.SharingLinks:
                        return Reports.SharingLinks.run(elTabs[report.name], this._appProps.auditOnly, formValues, null, () => { resolve(null); });
                    case ReportTypes.UniquePermissions:
                        return Reports.UniquePermissions.run(elTabs[report.name], this._appProps.auditOnly, formValues, null, () => { resolve(null); });
                }
            });
        }).then(() => {
        });
    }
}