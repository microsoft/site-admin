import { Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
import { IResponse } from "../ds";

/**
 * API Response Modal
 */
export class APIResponseModal {
    constructor(responses: IResponse[]) {
        // Return if no responses exist
        if (responses.length == 0) { return; }

        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("API Responses");

        // Render the responses
        this.renderResponses(responses);

        // Set the footer
        this.renderFooter();

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
                    name: "value",
                    title: "New Value"
                },
                {
                    name: "",
                    title: "Status",
                    onRenderCell: (el, col, item: IResponse) => {
                        // Render the status
                        item.errorFl ? "Error" : "Successful";
                    }
                },
                {
                    name: "message",
                    title: "API Reponse"
                }
            ]
        });
    }
}