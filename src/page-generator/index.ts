import { LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, SitePages, SPTypes, Types, Web } from "gd-sprest-bs";
import Strings from "../strings";
import * as Templates from "./templates";
import { } from "./templates/site-configuration";
import { PageTemplate } from "./template";

// Page Mapper
const PageMapper: { [key: string]: { filename: string; title: string; template: PageTemplate; } } = {
    "AppPermissions": { filename: "AppPermissions.aspx", title: "App Permissions", template: Templates.AppPermissions },
    "AuditReports": { filename: "AuditReports.aspx", title: "Audit Reports", template: Templates.AuditReports },
    "DataReadiness": { filename: "DataReadiness.aspx", title: "Data Readiness", template: Templates.DataReadiness },
    "Overview": { filename: "Overview.aspx", title: "Overview", template: Templates.Overview },
    "ReportDLP": { filename: "ReportDLP.aspx", title: "SAT - Data Loss Prevention (DLP)", template: Templates.ReportDLP },
    "ReportDocRetention": { filename: "ReportDocRetention.aspx", title: "SAT - Document Retention", template: Templates.ReportDocRetention },
    "ReportExtShares": { filename: "ReportExtShares.aspx", title: "SAT - External Shares", template: Templates.ReportExtShares },
    "ReportExtUsers": { filename: "ReportExtUsers.aspx", title: "SAT - External Users", template: Templates.ReportExtUsers },
    "ReportPermissions": { filename: "ReportPermissions.aspx", title: "SAT - Permissions", template: Templates.ReportPermissions },
    "ReportSearchAgents": { filename: "ReportSearchAgents.aspx", title: "SAT - Search Agents", template: Templates.ReportSearchAgents },
    "ReportSearchDocs": { filename: "ReportSearchDocs.aspx", title: "SAT - Search Documents", template: Templates.ReportSearchDocs },
    "ReportSearchEEEU": { filename: "ReportSearchEEEU.aspx", title: "SAT - Search External and External Users (EEEU)", template: Templates.ReportSearchEEEU },
    "ReportSearchProp": { filename: "ReportSearchProp.aspx", title: "SAT - Search Properties", template: Templates.ReportSearchProp },
    "ReportSearchUsers": { filename: "ReportSearchUsers.aspx", title: "SAT - Search Users", template: Templates.ReportSearchUsers },
    "ReportSensitivityLabels": { filename: "ReportSensitivityLabels.aspx", title: "SAT - Sensitivity Labels", template: Templates.ReportSensitivityLabels },
    "ReportSharingLinks": { filename: "ReportSharingLinks.aspx", title: "SAT - Sharing Links", template: Templates.ReportSharingLinks },
    "ReportUniquePermissions": { filename: "ReportUniquePermissions.aspx", title: "SAT - Unique Permissions", template: Templates.ReportUniquePermissions },
    "SiteConfiguration": { filename: "SiteConfiguration.aspx", title: "Site Configuration", template: Templates.SiteConfiguration }
}

/**
 * Page Generator
 */
export class PageGenerator {
    private static _imageReferences: string[];
    static get ImageReferences() { return this._imageReferences; }
    static set ImageReferences(values: string[]) { this._imageReferences = values; }

    // Constructor
    constructor() {
        // Show the form
        this.showForm();
    }

    // Creates the page
    private createFolder(folderName: string): PromiseLike<Types.SP.Folder> {
        // Show a loading dialog
        LoadingDialog.setHeader("Creating Folder");
        LoadingDialog.setBody("The folder is being created...");
        LoadingDialog.show();

        // Return a promise
        return new Promise(resolve => {
            // Get the site pages
            Web().Lists().getByTitle("Site Pages").RootFolder().Folders().execute(folders => {
                // Parse the folders
                for (let i = 0; i < folders.results.length; i++) {
                    let folder = folders.results[i];

                    // See if the target folder exists
                    if (folder.Name.toLowerCase() == folderName.toLowerCase()) {
                        // Resolve the request
                        resolve(folder);
                        return;
                    }
                }

                // Create the folder
                Web().Lists().getByTitle("Site Pages").RootFolder().Folders().add(folderName).execute(resolve);
            });
        });
    }

    // Creates the pages
    private createPages(folder: Types.SP.Folder) {
        // Show a loading dialog
        LoadingDialog.setHeader("Creating Pages");
        LoadingDialog.setBody("The pages are being created...");
        LoadingDialog.show();

        // Parse the files to create
        let counter = 0;
        let pages = Object.keys(PageMapper);
        Helper.Executor(pages, key => {
            let pageInfo = PageMapper[key];

            // Update the dialog
            LoadingDialog.setHeader(`Creating Page ${++counter} of ${pages.length}`);

            // Return a promise
            return new Promise(resolve => {
                // Create the page
                SitePages.createPage(`${folder.Name}/${pageInfo.filename}`, pageInfo.title, SPTypes.ClientSidePageLayout.Article, Strings.SourceUrl).then(page => {
                    // Get the folder url of the page
                    let pageFolderName = page.file.Name.replace('.aspx', "");

                    // Update the loading dialog
                    LoadingDialog.setBody("Configuring the page...");

                    // Upload the images
                    this.uploadImages(pageFolderName, page.item.Id, pageInfo.template).then(pageTemplate => {
                        // See if this is the main page
                        let content = pageTemplate.Content;
                        if (pageInfo.template == Templates.Overview) {
                            // Update the link references
                            content = content.replace(/href="\/sites\/demo\/site-admin/g, 'href="' + Strings.SourceUrl);
                        }

                        // Set the content for the page
                        page.item.update({ CanvasContent1: content }).execute(() => {
                            // See if this is the main page
                            if (pageInfo.template == Templates.Overview) {
                                // Show the page in a new tab
                                window.open(page.page.AbsoluteUrl, "_blank");
                            }

                            // Create the next page
                            resolve(null);
                        }, () => {
                            // Error updating the page
                            console.error("There was an error configuring the page.");
                            LoadingDialog.hide();

                            // Show an error modal in the modal
                            // TODO

                            // Create the next page
                            resolve(null);
                        });
                    });
                });
            });
        }).then(() => {
            // Hide the dialogs
            LoadingDialog.hide();
            Modal.hide();
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

        Modal.setBody(`<p>This tool will generate pages to help with data readiness guidance. Enter the folder name and click 'Generate'.</p>`);

        // Render the form
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [
                {
                    name: "Name",
                    title: "Folder Name",
                    description: "The folder name create in the Site Pages library.",
                    type: Components.FormControlTypes.TextField,
                    isPlainText: true,
                    required: true,
                    errorMessage: "A folder name is required.",
                    value: "Site Admin Tool"
                }
            ]
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            buttonType: Components.ButtonTypes.OutlinePrimary,
            tooltips: [
                {
                    content: "Click to create the pages.",
                    btnProps: {
                        text: "Create",
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let values = form.getValues();

                                // Create the folder
                                this.createFolder(values["Name"]).then(folder => {
                                    // Create the pages
                                    this.createPages(folder);
                                });
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
    private uploadImages(pageFolderName: string, pageItemId: number, page: PageTemplate): PromiseLike<PageTemplate> {
        // Return a promise
        return new Promise(resolve => {
            // Get the folder info for the page
            this.getFolderInfo(pageFolderName, pageItemId).then(pageFolderInfo => {
                // Create the data readiness page
                page.updateFolderInfo(ContextInfo.siteId, ContextInfo.webId, pageFolderInfo.listId, pageFolderInfo.folderUrl);

                // Do nothing if no image references exist
                let images = Object.keys(page.Images);
                if (images.length == 0) {
                    resolve(page);
                    return;
                }

                // Update the loading dialog
                LoadingDialog.setBody("Uploading the images...");

                // Get the web reference for the images in the SPFx folder
                Web.getWebUrlFromPageUrl(PageGenerator.ImageReferences[0]).execute(webRef => {
                    // Parse the image references
                    let ctr = 0;
                    Helper.Executor(images, imageKey => {
                        // Find the image
                        let imageRef = PageGenerator.ImageReferences.filter(imageRef => {
                            if (imageRef.indexOf(imageKey) > 0) { return imageRef; }
                        })[0];

                        // Update the loading dialog
                        LoadingDialog.setBody(`Uploading image ${++ctr} of ${images.length}...`);

                        // Return a promise
                        return new Promise(resolve => {
                            // Get the file
                            Web(webRef.GetWebUrlFromPageUrl).getFileByUrl(imageRef).execute(file => {
                                // Get the content
                                file.content().execute(data => {
                                    let fileInfo = imageRef.split('/');
                                    let fileName = fileInfo[fileInfo.length - 1].split('_')[0];

                                    // Upload the file
                                    SitePages().addImage(pageFolderName, fileName + ".png", pageItemId, null, data).execute(image => {
                                        // Update the template
                                        page.updateImageId(fileName, image.UniqueId);

                                        // Get the next file
                                        resolve(image);
                                    }, resolve);
                                }, resolve);
                            }, resolve);
                        });
                    }).then(() => {
                        // Update the loading dialog
                        LoadingDialog.setBody(`${images.length} images uploaded`);

                        // Resolve the request
                        resolve(page);
                    });
                });
            });
        });
    }
}