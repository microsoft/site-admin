import { LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, SitePages, SPTypes, Web } from "gd-sprest-bs";
import Strings from "../strings";
import { Template } from "./template";

/**
 * Page Generator
 */
export class PageGenerator {
    private static _imageReferences: string[];
    static set ImageReferences(values: string[]) { this._imageReferences = values; }

    // Constructor
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
            // Get the folder url of the page
            let pageFolderName = page.file.Name.replace('.aspx', "");

            // Upload the images for the page
            this.uploadImages(pageFolderName, page.item.Id).then(template => {
                // Update the loading dialog
                LoadingDialog.setBody("Configuring the page...");

                // Set the content for the page
                page.item.update(template.CanvasContent1).execute(() => {
                    // Show the page in a new tab
                    window.open(page.page.AbsoluteUrl, "_blank");

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
        });
    }

    // Gets the folder information for the page
    private getFolderInfo(pageName: string, pageId: number): PromiseLike<{ listId: string; folderUrl: string; }> {
        // Return a promise
        return new Promise((resolve) => {
            // Get the folder information
            SitePages().getOrCreateAssetFolder(pageName, true, null, pageId).execute(info => {
                let folderUrl = info["GetOrCreateAssetFolder"];

                // Get the folder
                Web().getFolderByServerRelativeUrl(folderUrl).ListItemAllFields().query({ Expand: ["ParentList"] }).execute(folder => {
                    // Resolve the request
                    resolve({
                        listId: folder.ParentList.Id,
                        folderUrl: info["GetOrCreateAssetFolder"]
                    });
                });
            });
        });
    }

    // Shows the form to generate a page
    private showForm() {
        // Clear the modal
        Modal.clear();

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

    // Uploads the images for the site page
    private uploadImages(pageFolderName: string, pageItemId: number): PromiseLike<Template> {
        // Return a promise
        return new Promise(resolve => {
            // Get the folder info for the page
            this.getFolderInfo(pageFolderName, pageItemId).then(pageFolderInfo => {
                // Create the template
                let template = new Template(ContextInfo.siteId, ContextInfo.webId, pageFolderInfo.listId, pageFolderInfo.folderUrl);

                // Update the loading dialog
                LoadingDialog.setBody("Uploading the images...");

                // Get the web url from the image
                Web.getWebUrlFromPageUrl(PageGenerator._imageReferences[0]).execute(webRef => {
                    // Parse the image references
                    let ctr = 0;
                    Helper.Executor(PageGenerator._imageReferences, imageUrl => {
                        // Update the loading dialog
                        LoadingDialog.setBody(`Uploading image ${++ctr} of ${PageGenerator._imageReferences.length}...`);

                        // Return a promise
                        return new Promise(resolve => {
                            // Get the file
                            Web(webRef.GetWebUrlFromPageUrl).getFileByUrl(imageUrl).execute(file => {
                                // Get the content
                                file.content().execute(data => {
                                    let fileInfo = imageUrl.split('/');
                                    let fileName = fileInfo[fileInfo.length - 1].split('_')[0];

                                    // Upload the file
                                    SitePages().addImage(pageFolderName, fileName + ".png", pageItemId, data).execute(image => {
                                        // Update the template
                                        template.updateImageId(fileName, image.UniqueId);

                                        // Get the next file
                                        resolve(image);
                                    }, resolve);
                                }, resolve);
                            }, resolve);
                        });
                    }).then(() => {
                        // Resolve the request
                        resolve(template);
                    });
                });
            });
        });
    }
}