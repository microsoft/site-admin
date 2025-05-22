import { LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo } from "gd-sprest-bs";
import { DataSource, RequestTypes } from "./ds";

/**
 * Site Attestation Form
 */
export class SiteAttestationForm {
    // Renders the modal
    constructor(formText: string) {
        // Render the form
        this.render(formText);
    }

    // Create the request item
    private createRequest() {
        // Show a loading dialog
        LoadingDialog.setHeader("Creating Request");
        LoadingDialog.setBody("Submitting the attestation for this site...");
        LoadingDialog.show();

        // Create the item
        DataSource.addRequest(DataSource.SiteContext.SiteFullUrl, [{
            key: RequestTypes.SiteAttestation,
            message: "The request to update the attestation for this site will be completed within 5 minutes.",
            value: ContextInfo.userEmail
        }]).then((responses) => {
            // Show the completion form
            this.showCompletionForm(responses[0].message);

            // Hide the dialog
            LoadingDialog.hide();
        });
    }

    // Renders the form
    private render(formText: string) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Site Attestation");

        // Set the body
        Modal.setBody(formText);

        // Set the footer
        Components.ButtonGroup({
            el: Modal.FooterElement,
            buttons: [
                {
                    text: "Close",
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                        // Close the modal
                        Modal.hide();
                    },
                },
                {
                    text: "Submit",
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                        this.createRequest();
                    },
                }
            ]
        });

        // Show the modal
        Modal.show();
    }

    // Displays the completion form
    private showCompletionForm(bodyText: string) {
        // Clear the Modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Site Attestation");

        // Set the body
        Modal.setBody(bodyText);

        // Set the footer
        Components.Button({
            el: Modal.FooterElement,
            text: "Close",
            type: Components.ButtonTypes.OutlinePrimary,
            onClick: () => {
                // Close the modal
                Modal.hide();
            }
        });
    }
}