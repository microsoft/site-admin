import { LoadingDialog, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
import { DataSource, ISiteInfo } from "../ds";

/**
 * Load Form
 */
export class Load {
    // Renders the modal
    private static render(onSuccess: (siteInfo: ISiteInfo) => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Load Site");

        // Set the body
        let form = Components.Form({
            el: Modal.BodyElement,
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

        // Render the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Validates that you are a site collection admin for the site entered.",
                    btnProps: {
                        text: "Load",
                        onClick: () => {
                            // Validate the form
                            if (form.isValid()) {
                                let ctrl = form.getControl("url");
                                let url = form.getValues()["url"];

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

        // Show the modal
        Modal.show();
    }

    // Shows the modal
    static show(onSuccess: (siteInfo: ISiteInfo) => void) {
        // Render the modal
        this.render(onSuccess);
    }
}