import { Components } from "gd-sprest-bs";
import * as Reports from "../reports";

// Report Types
enum ReportTypes {
    BrokenPermissions = "BrokenPermissions",
    DocRetention = "DocRetention",
    ExternalUsers = "ExternalUsers",
    FindUsers = "FindUsers",
    SearchDocs = "SearchDocs"
}

/**
 * Reports Tab
 */
export class ReportsTab {
    private _el: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement) {
        this._el = el;

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
                            text: "Broken Permissions",
                            data: "Scans for broken inheritance.",
                            value: ReportTypes.BrokenPermissions
                        },
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
                            text: "Find Users",
                            data: "Find users by keyword or account.",
                            value: ReportTypes.FindUsers
                        },
                        {
                            text: "Search Documents",
                            data: "Find documents by keywords.",
                            value: ReportTypes.SearchDocs
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
            case ReportTypes.BrokenPermissions:
                form.appendControls(Reports.BrokenPermissions.getFormFields());
                break;
            case ReportTypes.DocRetention:
                form.appendControls(Reports.DocRetention.getFormFields());
                break;
            case ReportTypes.ExternalUsers:
                break;
            case ReportTypes.FindUsers:
                break;
            case ReportTypes.SearchDocs:
                form.appendControls(Reports.SearchDocs.getFormFields());
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
                // Run the report
                switch (selectedReport) {
                    case ReportTypes.BrokenPermissions:
                        Reports.BrokenPermissions.run(this._el, form.getValues(), () => {
                            // Render this component
                            this.render();
                        });
                        break;
                    case ReportTypes.DocRetention:
                        Reports.DocRetention.run(this._el, form.getValues(), () => {
                            // Render this component
                            this.render();
                        });
                        break;
                    case ReportTypes.ExternalUsers:
                        break;
                    case ReportTypes.FindUsers:
                        break;
                    case ReportTypes.SearchDocs:
                        Reports.SearchDocs.run(this._el, form.getValues(), () => {
                            // Render this component
                            this.render();
                        });
                        break;
                }
            }
        });
    }
}