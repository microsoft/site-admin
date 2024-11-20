import { LoadingDialog, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
import { DataSource, ISiteInfo } from "../ds";

/**
 * Load Form
 */
export class Load {
    private static _form: Components.IForm = null;

    // Renders the modal
    private static render(onSuccess: (siteInfo: ISiteInfo) => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Load Site");

        // Render the form
        this.renderForm(Modal.BodyElement);

        // Render the footer
        this.renderFooter(Modal.FooterElement, onSuccess);

        // Show the modal
        Modal.show();
    }

    // Renders the footer
    static renderFooter(el: HTMLElement, onSuccess: (siteInfo: ISiteInfo) => void) {
        Components.TooltipGroup({
            el,
            tooltips: [
                {
                    content: "Validates that you are a site collection admin for the site entered.",
                    btnProps: {
                        text: "Load",
                        onClick: () => {
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
                                    (siteInfo) => {
                                        // Close the dialogs
                                        LoadingDialog.hide();
                                        Modal.hide();

                                        // Call the success method
                                        onSuccess(siteInfo);
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
                }
            ]
        });
    }

    // Renders the form
    static renderForm(el: HTMLElement) {
        this._form = Components.Form({
            el,
            controls: [
                {
                    name: "url",
                    label: "Site Collection Url:",
                    type: Components.FormControlTypes.TextField,
                    description: "The absolute/relative url to the site collection. (Example: /sites/dev)",
                    required: true,
                    errorMessage: "The site url is required."
                }
            ]
        });
    }

    // Shows the modal
    static show(onSuccess: (siteInfo: ISiteInfo) => void) {
        // Render the modal
        this.render(onSuccess);
    }
}