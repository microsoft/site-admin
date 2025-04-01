import { LoadingDialog, Modal } from "dattatable";
import { Components, SitePages, SPTypes } from "gd-sprest-bs";
import Strings from "../strings";
import { CanvasContent1 } from "./content";

/**
 * Page Generator
 */
export class PageGenerator {
    constructor() {
        // Show the form
        this.showForm();
    }

    // Creates the page
    private createPage(fileName: string, title: string) {
        // Show a loading dialog
        LoadingDialog.setHeader("Creating Page");
        LoadingDialog.setBody("The page is being created...");
        LoadingDialog.show();

        // Create the page
        SitePages.createPage(fileName, title, SPTypes.ClientSidePageLayout.Article, Strings.SourceUrl).then(page => {
            // Update the loading dialog
            LoadingDialog.setBody("Configuring the page...");

            // Set the content
            page.item.update({ CanvasContent1 }).execute(() => {
                // Update the loading dialog
                LoadingDialog.setBody("Uploading the images...");

                // TODO

                // Show the page in a new tab
                window.open(page.page.Url, "_blank");

                // Hide the dialogs
                LoadingDialog.hide();
                Modal.hide();
            }, () => {
                // Error updating the page
                console.error("There was an error configuring the page.");
                LoadingDialog.hide();

                // Show an error modal in the modal
                // TODO
            });
        });
    }

    // Shows the form to generate a page
    private showForm() {
        // Set the header
        Modal.setHeader("Page Generator");

        // Render the form
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [
                {
                    name: "Name",
                    title: "Page Name",
                    description: "The file name of the page to create.",
                    type: Components.FormControlTypes.TextField,
                    isPlainText: true,
                    required: true,
                    errorMessage: "A file name is required.",
                    value: "Copilot-Data-Readiness"
                },
                {
                    name: "Title",
                    title: "Page Title",
                    description: "The title displayed.",
                    type: Components.FormControlTypes.TextField,
                    isPlainText: true,
                    required: true,
                    errorMessage: "A title is required.",
                    value: "Copilot Data Readiness"
                },
            ]
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            buttonType: Components.ButtonTypes.OutlinePrimary,
            tooltips: [
                {
                    content: "Click to create the page.",
                    btnProps: {
                        text: "Create",
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let values = form.getValues();

                                // Create the page
                                this.createPage(values["Name"].split('.aspx')[0] + ".aspx", values["Title"]);
                            }
                        }
                    }
                },
                {
                    content: "Closes this dialog.",
                    btnProps: {
                        text: "Close",
                        onClick: () => {
                            // Hide the dialog
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