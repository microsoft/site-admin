import { Components } from "gd-sprest-bs";
import { DataSource } from "../ds";
import { IAppProps } from "../app";
import * as Reports from "../reports";
import { ISearchProps } from "./searchProp";

// Report Properties
export interface IReportProps {
    docRententionYears?: string;
    docSearchFileExt?: string;
    docSearchKeywords?: string;
}

// Report Types
enum ReportTypes {
    DocRetention = "DocRetention",
    ExternalShares = "ExternalShares",
    ExternalUsers = "ExternalUsers",
    Lists = "Lists",
    Permissions = "Permissions",
    SearchDocs = "SearchDocs",
    SearchEEEU = "SearchEEEU",
    SearchProp = "SearchProp",
    SearchUsers = "SearchUsers",
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
    private _reportProps: IReportProps = null;
    private _searchProps: ISearchProps = null;

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
    private render(selectedReport: string = ReportTypes.DocRetention) {
        // Clear the element
        while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

        // Render a form
        let form = Components.Form({
            el: this._el,
            controls: [
                {
                    name: "ReportType",
                    label: "Select Report",
                    description: "Select a report to run against this site.",
                    type: Components.FormControlTypes.Dropdown,
                    value: selectedReport,
                    items: [
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
                            text: "Lists/Libraries",
                            data: "Scans all of the lists/libraries for this site.",
                            value: ReportTypes.Lists
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

        // Add the controls
        switch (selectedReport) {
            case ReportTypes.DocRetention:
                form.appendControls(Reports.DocRetention.getFormFields(this._reportProps.docRententionYears));
                break;
            case ReportTypes.ExternalShares:
                form.appendControls(Reports.ExternalShares.getFormFields());
                break;
            case ReportTypes.ExternalUsers:
                form.appendControls(Reports.ExternalUsers.getFormFields());
                break;
            case ReportTypes.Lists:
                form.appendControls(Reports.Lists.getFormFields());
                break;
            case ReportTypes.Permissions:
                form.appendControls(Reports.Permissions.getFormFields());
                break;
            case ReportTypes.SearchDocs:
                form.appendControls(Reports.SearchDocs.getFormFields(this._reportProps?.docSearchFileExt, this._reportProps?.docSearchKeywords));
                break;
            case ReportTypes.SearchEEEU:
                form.appendControls(Reports.SearchEEEU.getFormFields());
                break;
            case ReportTypes.SearchProp:
                form.appendControls(Reports.SearchProp.getFormFields(DataSource.Site.RootWeb.AllProperties[this._searchProps.key]));
                break;
            case ReportTypes.SearchUsers:
                form.appendControls(Reports.SearchUsers.getFormFields());
                break;
            case ReportTypes.SharingLinks:
                form.appendControls(Reports.SharingLinks.getFormFields());
                break;
            case ReportTypes.UniquePermissions:
                form.appendControls(Reports.UniquePermissions.getFormFields());
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
                if (!form.isValid()) { return; }

                // Run the report
                switch (selectedReport) {
                    case ReportTypes.DocRetention:
                        Reports.DocRetention.run(this._el, this._auditOnly, form.getValues(), () => {
                            // Render this component
                            this.render(selectedReport);
                        });
                        break;
                    case ReportTypes.ExternalShares:
                        Reports.ExternalShares.run(this._el, this._auditOnly, form.getValues(), () => {
                            // Render this component
                            this.render(selectedReport);
                        });
                        break;
                    case ReportTypes.ExternalUsers:
                        Reports.ExternalUsers.run(this._el, this._auditOnly, form.getValues(), () => {
                            // Render this component
                            this.render(selectedReport);
                        });
                        break;
                    case ReportTypes.Lists:
                        Reports.Lists.run(this._el, this._auditOnly, form.getValues(), this._disableSensitivityLabelOverride, () => {
                            // Render this component
                            this.render(selectedReport);
                        });
                        break;
                    case ReportTypes.Permissions:
                        Reports.Permissions.run(this._el, this._auditOnly, form.getValues(), () => {
                            // Render this component
                            this.render(selectedReport);
                        });
                        break;
                    case ReportTypes.SearchDocs:
                        Reports.SearchDocs.run(this._el, this._auditOnly, form.getValues(), () => {
                            // Render this component
                            this.render(selectedReport);
                        });
                        break;
                    case ReportTypes.SearchEEEU:
                        Reports.SearchEEEU.run(this._el, this._auditOnly, form.getValues(), () => {
                            // Render this component
                            this.render(selectedReport);
                        });
                        break;
                    case ReportTypes.SearchProp:
                        let searchValue = form.getValues()["value"];
                        Reports.SearchProp.run(this._el, this._auditOnly, this._searchProps.managedProperty, typeof (searchValue) === "string" ? searchValue : searchValue?.value, () => {
                            // Render this component
                            this.render(selectedReport);
                        });
                        break;
                    case ReportTypes.SearchUsers:
                        // Ensure the values exist
                        let values = form.getValues();
                        if (values.UserName || values.PeoplePicker.length > 0) {
                            Reports.SearchUsers.run(this._el, this._auditOnly, values, () => {
                                // Render this component
                                this.render(selectedReport);
                            });
                        } else {
                            // Update the validation
                            let ctrl = form.getControl("UserName");
                            ctrl.updateValidation(ctrl.el, {
                                isValid: false,
                                invalidMessage: "A keyword or account is required to perform a search."
                            });
                            ctrl = form.getControl("PeoplePicker");
                            ctrl.updateValidation(ctrl.el, {
                                isValid: false,
                                invalidMessage: "A keyword or account is required to perform a search."
                            });
                        }
                        break;
                    case ReportTypes.SharingLinks:
                        Reports.SharingLinks.run(this._el, this._auditOnly, form.getValues(), () => {
                            // Render this component
                            this.render(selectedReport);
                        });
                        break;
                    case ReportTypes.UniquePermissions:
                        Reports.UniquePermissions.run(this._el, this._auditOnly, form.getValues(), () => {
                            // Render this component
                            this.render(selectedReport);
                        });
                        break;
                }
            }
        });
    }
}