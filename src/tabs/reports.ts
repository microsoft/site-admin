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
    private _auditOnly: boolean = null;
    private _el: HTMLElement = null;
    private _disableSensitivityLabelOverride: boolean = null;
    private _form: Components.IForm = null;
    private _reportProps: IReportProps = null;
    private _searchProps: ISearchProps = null;
    private _selectedReport: string = null;

    // Constructor
    constructor(el: HTMLElement, appProps: IAppProps) {
        this._auditOnly = !DataSource.IsAdmin || (appProps.auditOnly ? true : false);
        this._el = el;
        this._disableSensitivityLabelOverride = appProps.disableSensitivityLabelOverride;
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
                    items: [
                        {
                            text: "Data Loss Prevention",
                            data: "Finds files that has DLP applied to it.",
                            value: ReportTypes.DLP
                        },
                        {
                            text: "Document Retention",
                            data: "Find documents older than a specified date.",
                            value: ReportTypes.DocRetention
                        },
                        {
                            text: "External Shares",
                            data: "Scans for documents that have been shared externally.",
                            value: ReportTypes.ExternalShares
                        },
                        {
                            text: "External Users",
                            data: "Scans the user information list for 'external' user accounts.",
                            value: ReportTypes.ExternalUsers
                        },
                        {
                            text: "Permissions",
                            data: "Scans all users/groups that have permissions to the site.",
                            value: ReportTypes.Permissions
                        },
                        {
                            text: "Search Documents",
                            data: "Find documents by keywords.",
                            value: ReportTypes.SearchDocs
                        },
                        {
                            text: "Search EEEU",
                            data: "Search for the 'Every' and 'Everyone exception external users' accounts.",
                            value: ReportTypes.SearchEEEU
                        },
                        {
                            text: this._searchProps.reportName || "Search Property",
                            data: "Find sites by search property.",
                            value: ReportTypes.SearchProp,
                            isDisabled: this._searchProps.managedProperty && DataSource.SearchPropItems ? false : true
                        },
                        {
                            text: "Search Users",
                            data: "Search users by keyword or account.",
                            value: ReportTypes.SearchUsers
                        },
                        {
                            text: "Sensitivity Labels",
                            data: "Search for files that have sensitivity labels.",
                            value: ReportTypes.SensitivityLabels
                        },
                        {
                            text: "Sharing Links",
                            data: "Scans for any 'Sharing Link' groups.",
                            value: ReportTypes.SharingLinks
                        },
                        {
                            text: "Unique Permissions",
                            data: "Scans for items that have unique permissions.",
                            value: ReportTypes.UniquePermissions
                        }
                    ],
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

                // Run the report
                switch (this._selectedReport) {
                    case ReportTypes.DLP:
                        Reports.DLP.run(this._el, this._auditOnly, this._form.getValues(), () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.DocRetention:
                        Reports.DocRetention.run(this._el, this._auditOnly, this._form.getValues(), () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.ExternalShares:
                        Reports.ExternalShares.run(this._el, this._auditOnly, this._form.getValues(), () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.ExternalUsers:
                        Reports.ExternalUsers.run(this._el, this._auditOnly, this._form.getValues(), () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.Permissions:
                        Reports.Permissions.run(this._el, this._auditOnly, this._form.getValues(), () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.SearchDocs:
                        Reports.SearchDocs.run(this._el, this._auditOnly, this._form.getValues(), () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.SearchEEEU:
                        Reports.SearchEEEU.run(this._el, this._auditOnly, this._form.getValues(), () => {
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
                        let values = this._form.getValues();
                        if (values.UserName || values.PeoplePicker.length > 0) {
                            Reports.SearchUsers.run(this._el, this._auditOnly, values, () => {
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
                        Reports.SensitivityLabels.run(this._el, this._auditOnly, this._form.getValues(), () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.SharingLinks:
                        Reports.SharingLinks.run(this._el, this._auditOnly, this._form.getValues(), () => {
                            // Render this component
                            this.render(this._selectedReport);
                        });
                        break;
                    case ReportTypes.UniquePermissions:
                        Reports.UniquePermissions.run(this._el, this._auditOnly, this._form.getValues(), () => {
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