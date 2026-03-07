import { Components } from "gd-sprest-bs";
import { DataSource } from "../ds";
import { IAppProps } from "../app";
import * as Reports from "../reports";
import { ISearchProps } from "./searchProp";

// Report Properties
export interface IReportProps {
    dlpFileExt?: string;
    docRententionYears?: string;
    docSearchFileExt?: string;
    docSearchKeywords?: string;
    sensitivityLabelFileExt?: string;
}

// Report Types
enum ReportTypes {
    DLP = "DLP",
    DocRetention = "DocRetention",
    ExternalShares = "ExternalShares",
    ExternalUsers = "ExternalUsers",
    Permissions = "Permissions",
    SearchDocs = "SearchDocs",
    SearchEEEU = "SearchEEEU",
    SearchProp = "SearchProp",
    SearchUsers = "SearchUsers",
    SensitivityLabels = "SensitivityLabels",
    SharingLinks = "SharingLinks",
    UniquePermissions = "UniquePermissions"
}

/**
 * Reports Tab
 */
export class ReportsTab {
    private _appProps: IAppProps = null;
    private _auditOnly: boolean = null;
    private _el: HTMLElement = null;
    private _disableSensitivityLabelOverride: boolean = null;
    private _form: Components.IForm = null;
    private _loadOneDrive: boolean = false;
    private _reportProps: IReportProps = null;
    private _searchProps: ISearchProps = null;
    private _selectedReport: string = null;

    // Constructor
    constructor(el: HTMLElement, appProps: IAppProps, loadOneDrive: boolean) {
        this._appProps = appProps;
        this._auditOnly = !DataSource.IsAdmin || (appProps.auditOnly ? true : false);
        this._el = el;
        this._disableSensitivityLabelOverride = appProps.disableSensitivityLabelOverride;
        this._loadOneDrive = loadOneDrive;
        this._reportProps = appProps.reportProps;
        this._searchProps = appProps.searchProps;

        // Render the tab
        this.render();
    }

    // Renders the tab
    private render(selectedReport: string = ReportTypes.DLP) {
        // Set the selected report
        this._selectedReport = selectedReport;

        // Clear the element
        while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

        // Set the reports to display
        let items: Components.IDropdownItem[] = [];

        // Add the reports
        if (typeof (this._appProps.hideReports.dlp) === "undefined" || this._appProps.hideReports.dlp != true) {
            items.push({
                text: "Data Loss Prevention",
                data: "Finds files that has DLP applied to it.",
                value: ReportTypes.DLP
            });
        }
        if (typeof (this._appProps.hideReports.docRetention) === "undefined" || this._appProps.hideReports.docRetention != true) {
            items.push({
                text: "Document Retention",
                data: "Find documents older than a specified date.",
                value: ReportTypes.DocRetention
            });
        }
        if (typeof (this._appProps.hideReports.externalShares) === "undefined" || this._appProps.hideReports.externalShares != true) {
            items.push({
                text: "External Shares",
                data: "Scans for documents that have been shared externally.",
                value: ReportTypes.ExternalShares
            });
        }
        if (typeof (this._appProps.hideReports.externalUsers) === "undefined" || this._appProps.hideReports.externalUsers != true) {
            items.push({
                text: "External Users",
                data: "Scans the user information list for 'external' user accounts.",
                value: ReportTypes.ExternalUsers
            });
        }
        if (typeof (this._appProps.hideReports.permissions) === "undefined" || this._appProps.hideReports.permissions != true) {
            items.push({
                text: "Permissions",
                data: "Scans all users/groups that have permissions to the site.",
                value: ReportTypes.Permissions
            });
        }
        if (typeof (this._appProps.hideReports.searchDocs) === "undefined" || this._appProps.hideReports.searchDocs != true) {
            items.push({
                text: "Search Documents",
                data: "Find documents by keywords.",
                value: ReportTypes.SearchDocs
            });
        }
        if (typeof (this._appProps.hideReports.searchEEEU) === "undefined" || this._appProps.hideReports.searchEEEU != true) {
            items.push({
                text: "Search EEEU",
                data: "Search for the 'Every' and 'Everyone exception external users' accounts.",
                value: ReportTypes.SearchEEEU
            });
        }
        if (typeof (this._appProps.hideReports.searchProp) === "undefined" || this._appProps.hideReports.searchProp != true) {
            // Ensure the property is set
            if (this._searchProps.reportName) {
                items.push({
                    text: this._searchProps.reportName || "Search Property",
                    data: "Find sites by search property.",
                    value: ReportTypes.SearchProp,
                    isDisabled: this._searchProps.managedProperty && DataSource.SearchPropItems ? false : true
                });
            }
        }
        if (typeof (this._appProps.hideReports.searchUsers) === "undefined" || this._appProps.hideReports.searchUsers != true) {
            items.push({
                text: "Search Users",
                data: "Search users by keyword or account.",
                value: ReportTypes.SearchUsers
            });
        }
        if (typeof (this._appProps.hideReports.sensitivityLabels) === "undefined" || this._appProps.hideReports.sensitivityLabels != true) {
            items.push({
                text: "Sensitivity Labels",
                data: "Search for files that have sensitivity labels.",
                value: ReportTypes.SensitivityLabels
            });
        }
        if (typeof (this._appProps.hideReports.sharingLinks) === "undefined" || this._appProps.hideReports.sharingLinks != true) {
            items.push({
                text: "Sharing Links",
                data: "Scans for any 'Sharing Link' groups.",
                value: ReportTypes.SharingLinks
            });
        }
        if (typeof (this._appProps.hideReports.uniquePermissions) === "undefined" || this._appProps.hideReports.uniquePermissions != true) {
            items.push({
                text: "Unique Permissions",
                data: "Scans for items that have unique permissions.",
                value: ReportTypes.UniquePermissions
            });
        }


        // See if this is onedrive
        if (this._loadOneDrive) {
            // Remove some of the reports
            items.splice(8, 1);
            items.splice(7, 1);
        }

        // Render a form
        this._form = Components.Form({
            el: this._el,
            controls: [
                {
                    name: "ReportType",
                    label: "Select Report",
                    description: "Select a report to run against this site.",
                    type: Components.FormControlTypes.Dropdown,
                    value: this._selectedReport,
                    items,
                    onChange: item => {
                        // Render the form for this report
                        this.render(item?.value);
                    }
                } as Components.IFormControlPropsDropdown
            ]
        });

        // See if webs exist
        if (DataSource.SiteItems && DataSource.SiteItems.length > 1) {
            // See if this is a report that can be run against a specified web
            switch (this._selectedReport) {
                case ReportTypes.DLP:
                case ReportTypes.ExternalUsers:
                case ReportTypes.Permissions:
                case ReportTypes.SearchEEEU:
                case ReportTypes.SearchUsers:
                case ReportTypes.SensitivityLabels:
                case ReportTypes.UniquePermissions:
                    // Add a sub-web control
                    this._form.insertControl(1, {
                        name: "TargetWeb",
                        label: "Target Web",
                        description: "Select a web to run the report.",
                        type: Components.FormControlTypes.Dropdown,
                        items: [{ text: "All Sites", value: null } as Components.IDropdownItem].concat(DataSource.SiteItems)
                    } as Components.IFormControlPropsDropdown);
                    break;
            }
        }

        // Add the controls
        switch (this._selectedReport) {
            case ReportTypes.DLP:
                this._form.appendControls(Reports.DLP.getFormFields(this._reportProps?.dlpFileExt));
                break;
            case ReportTypes.DocRetention:
                this._form.appendControls(Reports.DocRetention.getFormFields(this._reportProps.docRententionYears));
                break;
            case ReportTypes.ExternalShares:
                this._form.appendControls(Reports.ExternalShares.getFormFields());
                break;
            case ReportTypes.ExternalUsers:
                this._form.appendControls(Reports.ExternalUsers.getFormFields());
                break;
            case ReportTypes.Permissions:
                this._form.appendControls(Reports.Permissions.getFormFields());
                break;
            case ReportTypes.SearchDocs:
                this._form.appendControls(Reports.SearchDocs.getFormFields(this._reportProps?.docSearchFileExt, this._reportProps?.docSearchKeywords));
                break;
            case ReportTypes.SearchEEEU:
                this._form.appendControls(Reports.SearchEEEU.getFormFields());
                break;
            case ReportTypes.SearchProp:
                this._form.appendControls(Reports.SearchProp.getFormFields(DataSource.Site.RootWeb.AllProperties[this._searchProps.key]));
                break;
            case ReportTypes.SearchUsers:
                this._form.appendControls(Reports.SearchUsers.getFormFields());
                break;
            case ReportTypes.SensitivityLabels:
                this._form.appendControls(Reports.SensitivityLabels.getFormFields());
                break;
            case ReportTypes.SharingLinks:
                this._form.appendControls(Reports.SharingLinks.getFormFields());
                break;
            case ReportTypes.UniquePermissions:
                this._form.appendControls(Reports.UniquePermissions.getFormFields());
                break;
        }

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
                if (!this._form.isValid()) { return; }

                // Set the form values
                let formValues = this._form.getValues();
                formValues["LoadOneDrive"] = this._loadOneDrive ? "true" : "false";

                // Run the report
                switch (this._selectedReport) {
                    case ReportTypes.DLP:
                        Reports.DLP.run(this._el, this._auditOnly, formValues, () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.DocRetention:
                        Reports.DocRetention.run(this._el, this._auditOnly, formValues, () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.ExternalShares:
                        Reports.ExternalShares.run(this._el, this._auditOnly, formValues, () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.ExternalUsers:
                        Reports.ExternalUsers.run(this._el, this._auditOnly, formValues, () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.Permissions:
                        Reports.Permissions.run(this._el, this._auditOnly, formValues, () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.SearchDocs:
                        Reports.SearchDocs.run(this._el, this._auditOnly, formValues, () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.SearchEEEU:
                        Reports.SearchEEEU.run(this._el, this._auditOnly, formValues, () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.SearchProp:
                        let searchValue = this._form.getValues()["value"];
                        Reports.SearchProp.run(this._el, this._auditOnly, this._searchProps.managedProperty, typeof (searchValue) === "string" ? searchValue : searchValue?.value, () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.SearchUsers:
                        // Ensure the values exist
                        if (formValues.UserName || formValues.PeoplePicker.length > 0) {
                            Reports.SearchUsers.run(this._el, this._auditOnly, formValues, () => {
                                // Render this component
                                this.render(this._selectedReport);
                            });
                        } else {
                            // Update the validation
                            let ctrl = this._form.getControl("UserName");
                            ctrl.updateValidation(ctrl.el, {
                                isValid: false,
                                invalidMessage: "A keyword or account is required to perform a search."
                            });
                            ctrl = this._form.getControl("PeoplePicker");
                            ctrl.updateValidation(ctrl.el, {
                                isValid: false,
                                invalidMessage: "A keyword or account is required to perform a search."
                            });
                        }
                        break;
                    case ReportTypes.SensitivityLabels:
                        Reports.SensitivityLabels.run(this._el, this._auditOnly, formValues, () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.SharingLinks:
                        Reports.SharingLinks.run(this._el, this._auditOnly, formValues, () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.UniquePermissions:
                        Reports.UniquePermissions.run(this._el, this._auditOnly, formValues, () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                }
            }
        });
    }

    // Method to add the component
    setWebs(webs: Components.IDropdownItem[]) {
        // See if this is a report that can be run against a specified web
        switch (this._selectedReport) {
            case ReportTypes.DLP:
            case ReportTypes.ExternalUsers:
            case ReportTypes.Permissions:
            case ReportTypes.SearchEEEU:
            case ReportTypes.SearchUsers:
            case ReportTypes.SensitivityLabels:
            case ReportTypes.UniquePermissions:
                // See if more than one sub-webs are available
                if (webs.length > 1) {
                    // Add a sub-web control
                    this._form.insertControl(1, {
                        name: "TargetWeb",
                        label: "Target Web",
                        description: "Select a web to run the report.",
                        type: Components.FormControlTypes.Dropdown,
                        items: [{ text: "All Sites", value: null } as Components.IDropdownItem].concat(webs)
                    } as Components.IFormControlPropsDropdown);
                }
                break;
        }
    }
}