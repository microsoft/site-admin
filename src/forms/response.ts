import { Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
import { IResponse } from "../ds";

/**
 * API Response Modal
 */
export class APIResponseModal {
    constructor(responses: IResponse[]) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Change Request(s)");

        // Render the responses
        this.renderResponses(responses);

        // Show the modal
        Modal.show();
    }

    // Renders the footer
    private renderFooter() {
        // Render a close button
        Components.Tooltip({
            el: Modal.FooterElement,
            content: "Closes the modal.",
            btnProps: {
                text: "Close",
                type: Components.ButtonTypes.OutlinePrimary,
                onClick: () => {
                    // Close the modal
                    Modal.hide();
                }
            }
        })
    }

    // Renders the reponses
    private renderResponses(responses: IResponse[]) {
        // See if there are no responses
        if (responses.length == 0) {
            // Set the message
            Modal.setBody("No changes were detected.");
            return;
        }

        // Render a table
        Components.Table({
            el: Modal.BodyElement,
            rows: responses,
            columns: [
                {
                    name: "key",
                    title: "Property"
                },
                {
                    name: "",
                    title: "Request Submitted?",
                    onRenderCell: (el, col, item: IResponse) => {
                        // Render the status
                        item.errorFl ? "Error" : "Successful";
                    }
                },
                {
                    name: "message",
                    title: "Message"
                }
            ]
        });

        // Set the footer
        this.renderFooter();
    }
}