import { CanvasForm, Dashboard, Modal } from "dattatable";
import { Components, Helper } from "gd-sprest-bs";
import { IAppProps } from "./app";

/**
 * Dialog to select the default reports to be selected for the site audit tab.
 */
export class SiteAuditReportDialog {
    private _appProps: IAppProps = null;
    private _onUpdate: (reports: string[]) => void;

    // Constructor
    constructor(appProps: IAppProps, reports: string[], onUpdate: (reports: string[]) => void) {
        // Set the update callback
        this._appProps = appProps;
        this._onUpdate = onUpdate;

        // Initializes the dialog
        this.init(reports);
    }

    // Initializes the dialog
    private init(reports: string[]) {
        // Clear the modal
        Modal.clear();
        Modal.setType(Components.ModalTypes.Large);

        // Set the header
        Modal.setHeader("Site Audit Reports");

        // Render the form
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [
                {
                    name: "Reports",
                    label: "Reports",
                    type: Components.FormControlTypes.MultiSwitch,
                    value: reports,
                    items: [
                        { label: this._appProps.reportNames.dlp || "Data Loss Prevention" },
                        { label: this._appProps.reportNames.docRetention || "Document Retention" },
                        { label: this._appProps.reportNames.externalShares || "External Shares" },
                        { label: this._appProps.reportNames.externalUsers || "External Users" },
                        { label: this._appProps.reportNames.permissions || "Permissions" },
                        { label: this._appProps.reportNames.searchAgents || "Search Agents" },
                        //{ label: this._appProps.reportNames.searchDocs || "Search Documents"},
                        { label: this._appProps.reportNames.searchEEEU || "Search EEEU" },
                        { label: this._appProps.reportNames.sensitivityLabels || "Sensitivity Labels" },
                        { label: this._appProps.reportNames.sharingLinks || "Sharing Links" },
                        { label: this._appProps.reportNames.uniquePermissions || "Unique Permissions" }
                    ]
                } as Components.IFormControlPropsMultiSwitch
            ]
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Click to update the regex patterns.",
                    btnProps: {
                        text: "Update",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Get the reports
                            let reports = []
                            form.getValues()["Reports"].forEach(report => {
                                // Add the report
                                reports.push(report.label);
                            });

                            // Call the event
                            this._onUpdate(reports);

                            // Close the modal
                            Modal.hide();
                        }
                    }
                }
            ]
        });

        // Show the modal
        Modal.show();
    }
}