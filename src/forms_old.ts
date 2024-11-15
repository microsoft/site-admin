import { LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Web } from "gd-sprest-bs";
import { DataSource, IListItem } from "./ds";


// Forms
export class Forms {
    // Shows the delete dialog
    static deleteForm(item: IListItem, onDelete: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Delete Request");

        // Set the body
        Modal.setBody("Are you sure you want to delete this request?");

        // Render a delete button
        Components.Button({
            el: Modal.FooterElement,
            type: Components.ButtonTypes.OutlineDanger,
            text: "Delete",
            onClick: () => {
                // Show a loading dialog
                LoadingDialog.setHeader("Deleting Request");
                LoadingDialog.setBody("This will close after the request is removed...");
                LoadingDialog.show();

                // Delete the item
                item.delete().execute(() => {
                    // Refresh the data
                    DataSource.refresh().then(() => {
                        // Call the update event
                        onDelete();

                        // Hide the dialogs
                        LoadingDialog.hide();
                        Modal.hide();
                    });
                }, () => {
                    // Hide the dialogs
                    LoadingDialog.hide();
                    Modal.hide();
                });
            }
        });

        // Show the dialog
        Modal.show();
    }

    // Shows the new form
    static newForm(onUpdate: () => void) {
        // Show the new form
        DataSource.List.newForm({
            onSetHeader: el => {
                // Set the header
                el.firstElementChild.innerHTML = "Create Request";
            },
            onFormButtonsRendering: buttons => {
                // Modify the create button's text
                buttons[0].text = "Submit";

                // Return the buttons
                return buttons;
            },
            onCreateEditForm: props => {
                props.onControlRendering = (ctrl, fld) => {
                    // See if this is the owners field
                    if (fld.InternalName == "Owners") {
                        // Default the owners to the current user
                        ctrl.value = [ContextInfo.userId];
                    }
                }

                // Return the properties
                return props;
            },
            onValidation: (values, isValid) => {
                // See if the form values have been set
                if (isValid) {
                    // Return a promise
                    return new Promise(resolve => {
                        let ctrlSiteUrl = DataSource.List.EditForm.getControl("Title");

                        // Show a loading dialog
                        LoadingDialog.setHeader("Getting Site Information");
                        LoadingDialog.setBody("Validating the site url...");
                        LoadingDialog.show();

                        // Validate the site and permissions
                        let webUrl = values["Title"];
                        DataSource.validate(webUrl).then(
                            // Success
                            siteInfo => {
                                // Web exists, update the site url to be absolute
                                values["Title"] = siteInfo.web.Url;

                                // Resolve the request
                                resolve(true);
                            },

                            errorMsg => {
                                // Update the validation
                                ctrlSiteUrl.updateValidation(ctrlSiteUrl.el, {
                                    isValid: false,
                                    invalidMessage: errorMsg
                                });

                                // Hide the loading dialog
                                LoadingDialog.hide();

                                // Resolve the request
                                resolve(false);
                            }
                        );
                    });
                }

                // Return the flag by default
                return isValid;
            },
            onUpdate: (item: IListItem) => {
                // Show a loading dialog
                LoadingDialog.setHeader("Refreshing Requests");
                LoadingDialog.setBody("Refreshing the data...");
                LoadingDialog.show();

                // Refresh the data
                DataSource.refresh(item.Id).then(() => {
                    // See if the azure function is enabled
                    if (DataSource.AzureFunctionEnabled) {
                        // Update the loading dialog
                        LoadingDialog.setBody("Processing the request...");

                        // Process the request
                        this.processRequest(item.Title, item.Id).then((updateFl) => {
                            // Call the update event
                            updateFl ? onUpdate() : null;

                            // Hide the dialog
                            LoadingDialog.hide();
                        });
                    } else {
                        // Call the update event
                        onUpdate();

                        // Hide the dialog
                        LoadingDialog.hide();
                    }
                });
            }
        });
    }

    // Method to process a request
    static processRequest(webUrl: string, itemId: number): PromiseLike<boolean> {
        // Return a promise
        return new Promise((resolve) => {
            // Show a loading dialog
            LoadingDialog.setHeader("Validating Site");
            LoadingDialog.setBody("The site and permissions are being validated...");
            LoadingDialog.show();

            // Validate the web and permissions
            DataSource.validate(webUrl).then(
                // Success
                web => {
                    // Show a loading dialog
                    LoadingDialog.setHeader("Processing Request");
                    LoadingDialog.setBody("The request is being processed...");

                    // Process the request
                    DataSource.processRequest(itemId).then(
                        // Success
                        (response) => {
                            // Refresh the item
                            DataSource.refresh(itemId).then(() => {
                                // Hide the loading dialog
                                LoadingDialog.hide();

                                // Resolve the request
                                resolve(true);
                            });

                            // Clear the modal
                            Modal.clear();

                            // Set the header
                            Modal.setHeader("Request Processed");

                            // Set the body
                            Modal.setBody(response);

                            // Show the modal
                            Modal.show();
                        },

                        // Error
                        errorMsg => {
                            // Clear the modal
                            Modal.clear();

                            // Set the header
                            Modal.setHeader("Error Processing Request");

                            // Set the body
                            Modal.setBody(errorMsg);

                            // Show the modal
                            Modal.show();

                            // Hide the loading dialog
                            LoadingDialog.hide();

                            // Resolve the request
                            resolve(false);
                        }
                    );
                },

                errorMsg => {
                    // Clear the modal
                    Modal.clear();

                    // Set the header
                    Modal.setHeader("Error Processing Request");

                    // Set the body
                    Modal.setBody(errorMsg);

                    // Show the modal
                    Modal.show();

                    // Hide the loading dialog
                    LoadingDialog.hide();

                    // Resolve the request
                    resolve(false);
                }
            );
        });
    }
}