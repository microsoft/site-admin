import { LoadingDialog, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
import { DataSource } from "./ds";

/**
 * Load Form
 */
export class LoadForm {
    _form: Components.IForm = null;
    _onSuccess: () => void = null;

    // Renders the modal
    constructor(elForm: HTMLElement, elFooter: HTMLElement, onSuccess: () => void) {
        this._onSuccess = onSuccess;

        // Render the form and footer
        this.renderForm(elForm);
        this.renderFooter(elFooter);
    }

    // Renders the footer
    private renderFooter(el: HTMLElement) {
        Components.TooltipGroup({
            el,
            tooltips: [
                {
                    content: "Validates that you are an admin for the site entered.",
                    btnProps: {
                        text: "Load Site",
                        onClick: () => {
                            // Submit the request
                            this.submitForm();
                        }
                    }
                }
            ]
        });
    }

    // Renders the form
    private renderForm(el: HTMLElement) {
        this._form = Components.Form({
            el,
            controls: [
                {
                    name: "url",
                    label: "Site Url:",
                    type: Components.FormControlTypes.TextField,
                    description: "The absolute/relative url to the site. (Example: /sites/dev)",
                    required: true,
                    errorMessage: "The site url is required.",
                    onControlRendered: ctrl => {
                        // Set the key down event
                        ctrl.textbox.elTextbox.addEventListener("keypress", ev => {
                            // See if they hit the enter button
                            if (ev["keyCode"] === 13) {
                                ev.preventDefault();

                                // Submit the request
                                this.submitForm();
                            }
                        });
                    }
                }
            ]
        });
    }

    // Shows the modal
    static showModal(onSuccess: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Load Site");

        // Create the form/footer
        new LoadForm(Modal.BodyElement, Modal.FooterElement, onSuccess);

        // Show the modal
        Modal.show();
    }

    // Submits the form
    private submitForm() {
        // Validate the form
        if (this._form.isValid()) {
            let ctrl = this._form.getControl("url");
            let url = this._form.getValues()["url"];

            // Show a loading dialog
            LoadingDialog.setHeader("Validating Site");
            LoadingDialog.setBody("This will close after the site url is validated...");
            LoadingDialog.show();

            // Validate the url
            DataSource.validate(url).then(
                // Success
                () => {
                    // Close the dialogs
                    LoadingDialog.hide();
                    Modal.hide();

                    // Call the success method
                    this._onSuccess();
                },

                // Error
                (errorMessage) => {
                    // Close the dialog
                    LoadingDialog.hide();

                    // Update the validation
                    ctrl.updateValidation(ctrl.el, {
                        isValid: false,
                        invalidMessage: errorMessage
                    });
                }
            )
        }
    }
}