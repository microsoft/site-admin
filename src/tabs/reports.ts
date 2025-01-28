import { Components } from "gd-sprest-bs";
import * as Reports from "../reports";

// Report Properties
export interface IReportProps {
    docSearchFileExt?: string;
    docSearchKeywords?: string;
}

// Report Types
enum ReportTypes {
    DocRetention = "DocRetention",
    ExternalUsers = "ExternalUsers",
    SearchDocs = "SearchDocs",
    SearchUsers = "SearchUsers",
    UniquePermissions = "UniquePermissions"
}

/**
 * Reports Tab
 */
export class ReportsTab {
    private _el: HTMLElement = null;
    private _reportProps: IReportProps = null;

    // Constructor
    constructor(el: HTMLElement, reportProps: IReportProps) {
        this._el = el;
        this._reportProps = reportProps;

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
                            text: "External Users",
                            data: "Scans for external user information.",
                            value: ReportTypes.ExternalUsers
                        },
                        {
                            text: "Search Documents",
                            data: "Find documents by keywords.",
                            value: ReportTypes.SearchDocs
                        },
                        {
                            text: "Search Users",
                            data: "Search users by keyword or account.",
                            value: ReportTypes.SearchUsers
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
                form.appendControls(Reports.DocRetention.getFormFields());
                break;
            case ReportTypes.ExternalUsers:
                form.appendControls(Reports.ExternalUsers.getFormFields());
                break;
            case ReportTypes.SearchDocs:
                form.appendControls(Reports.SearchDocs.getFormFields(this._reportProps?.docSearchFileExt, this._reportProps?.docSearchKeywords));
                break;
            case ReportTypes.SearchUsers:
                form.appendControls(Reports.SearchUsers.getFormFields());
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
                        Reports.DocRetention.run(this._el, form.getValues(), () => {
                            // Render this component
                            this.render(selectedReport);
                        });
                        break;
                    case ReportTypes.ExternalUsers:
                        Reports.ExternalUsers.run(this._el, form.getValues(), () => {
                            // Render this component
                            this.render(selectedReport);
                        });
                        break;
                    case ReportTypes.SearchDocs:
                        Reports.SearchDocs.run(this._el, form.getValues(), () => {
                            // Render this component
                            this.render(selectedReport);
                        });
                        break;
                    case ReportTypes.SearchUsers:
                        // Ensure the values exist
                        let values = form.getValues();
                        if (values.UserName || values.PeoplePicker.length > 0) {
                            Reports.SearchUsers.run(this._el, values, () => {
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
                    case ReportTypes.UniquePermissions:
                        Reports.UniquePermissions.run(this._el, form.getValues(), () => {
                            // Render this component
                            this.render(selectedReport);
                        });
                        break;
                }
            }
        });
    }
}