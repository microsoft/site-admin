import { Components } from "gd-sprest-bs";
import * as moment from "moment";
import { DocRetention } from "../reports";

// Report Types
enum ReportTypes {
    DocRetention = "DocRetention",
    ExternalUsers = "ExternalUsers",
    FindUsers = "FindUsers",
    ListPermissions = "ListPermissions",
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
                            text: "List Permissions",
                            data: "Scans the permissions for all lists and libraries.",
                            value: ReportTypes.ListPermissions
                        },
                        {
                            text: "Search Documents",
                            data: "Find documents by keywords.",
                            value: ReportTypes.SearchDocs
                        }
                    ],
                    onChange: item => {
                        // See which report was selected
                        switch (item?.value) {
                            // Doc Retention
                            case ReportTypes.DocRetention:

                                break;
                        }
                    }
                } as Components.IFormControlPropsDropdown
            ]
        });

        // Add the controls
        switch (selectedReport) {
            case ReportTypes.DocRetention:
                form.appendControls(DocRetention.getFormFields());
                break;
            case ReportTypes.ExternalUsers:
                break;
            case ReportTypes.FindUsers:
                break;
            case ReportTypes.ListPermissions:
                break;
            case ReportTypes.SearchDocs:
                break;
        }

        // Render a footer
        let elFooter = document.createElement("div");
        elFooter.classList.add("d-flex");
        elFooter.classList.add("align-items-end");
        this._el.appendChild(elFooter);

        // Add a button
        Components.Button({
            el: elFooter,
            text: "Run",
            onClick: () => {
                // Run the report
                switch (selectedReport) {
                    case ReportTypes.DocRetention:
                        DocRetention.run(this._el, moment(form.getValues()["SelectedDate"]).format("YYYY-MM-DD"), () => {
                            // Render this component
                            this.render();
                        });
                        break;
                    case ReportTypes.ExternalUsers:
                        break;
                    case ReportTypes.FindUsers:
                        break;
                    case ReportTypes.ListPermissions:
                        break;
                    case ReportTypes.SearchDocs:
                        break;
                }
            }
        });
    }
}