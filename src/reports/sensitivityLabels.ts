import { Dashboard, Documents, LoadingDialog } from "dattatable";
import { Components, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { fileEarmark } from "gd-sprest-bs/build/icons/svgs/fileEarmark";
import { DataSource } from "../ds";
import { ExportCSV } from "./exportCSV";

interface ISensitivityLabelItem {
    Author: string;
    FileExtension: string;
    FileName: string;
    ListId: string;
    ListTitle: string;
    Path: string;
    SensitivityLabel: string;
    SensitivityLabelId: string;
    WebUrl: string;
    WebId: string;
}

interface IWebItem {
    ListId: string;
    ListTitle: string;
    WebId: string;
    WebUrl: string;
}

const CSVFields = [
    "Author",
    "FileExtension",
    "FileName",
    "ListId",
    "ListTitle",
    "Path",
    "SensitivityLabel",
    "SensitivityLabelId",
    "WebId",
    "WebUrl"
]

export class SensitivityLabels {
    private static _items: ISensitivityLabelItem[] = [];

    // Gets the form fields to display
    static getFormFields(): Components.IFormControlProps[] { return []; }

    // Analyzes the libraries
    private static analyzeLibraries(webId: string, webUrl: string, libraries: Types.SP.ListOData[]) {
        // Return a promise
        return new Promise(resolve => {
            let counter = 0;

            // Parse the libraries
            Helper.Executor(libraries, lib => {
                // Update the dialog
                LoadingDialog.setBody(`Analyzing Library ${lib.Title}<br/>${++counter} of ${libraries.length}`);

                // Return a promise
                return new Promise(resolve => {
                    // Get the files for this library
                    DataSource.loadFiles(webId, lib.Title).then(files => {
                        // Parse the files
                        files.forEach(file => {
                            // Ensure a sensitivity label exists
                            if (file.sensitivityLabel && file.sensitivityLabel.displayName) {
                                let fileInfo = file.name.split('.');
                                let folderPath = file.parentReference.path.split('root:')[1];

                                // Append the data
                                this._items.push({
                                    Author: file.createdBy.user["email"],
                                    FileExtension: fileInfo[fileInfo.length - 1],
                                    FileName: file.name,
                                    ListId: lib.Id,
                                    ListTitle: lib.Title,
                                    Path: `${lib.RootFolder.ServerRelativeUrl}${folderPath}/${file.name}`,
                                    SensitivityLabel: file.sensitivityLabel.displayName,
                                    SensitivityLabelId: file.sensitivityLabel.id,
                                    WebId: webId,
                                    WebUrl: webUrl
                                });
                            }
                        });

                        // Check the next library
                        resolve(null);
                    });
                });
            }).then(resolve);
        });
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, auditOnly: boolean, onClose: () => void) {
        // Render the summary
        new Dashboard({
            el,
            navigation: {
                title: "Sensitivity Labels",
                showFilter: false,
                items: [{
                    text: "New Search",
                    className: "btn-outline-light",
                    isButton: true,
                    onClick: () => {
                        // Call the close event
                        onClose();
                    }
                }],
                itemsEnd: [{
                    text: "Export to CSV",
                    className: "btn-outline-light me-2",
                    isButton: true,
                    onClick: () => {
                        // Export the CSV
                        new ExportCSV("SensitivityLabels.csv", CSVFields, this._items);
                    }
                }]
            },
            table: {
                rows: this._items,
                onRendering: dtProps => {
                    dtProps.columnDefs = [
                        {
                            "targets": 4,
                            "orderable": false,
                            "searchable": false
                        }
                    ];

                    // Order by the 1st column by default; ascending
                    dtProps.order = [[3, "asc"]];

                    // Return the properties
                    return dtProps;
                },
                columns: [
                    {
                        name: "ListTitle",
                        title: "List"
                    },
                    {
                        name: "FileName",
                        title: "File"
                    },
                    {
                        name: "",
                        title: "File Info",
                        onRenderCell: (el, col, item: ISensitivityLabelItem) => {
                            // Set the sort value
                            el.setAttribute("data-order", item.Path);

                            // Show the file info
                            el.innerHTML = `
                                <b>Create By: </b>${item.Author}
                                <br/>
                                <b>Path: </b>${item.Path}
                            `;
                        }
                    },
                    {
                        name: "SensitivityLabel",
                        title: "Sensitivity Label"
                    },
                    {
                        className: "text-end",
                        name: "",
                        title: "",
                        onRenderCell: (el, col, row: ISensitivityLabelItem) => {
                            let btnDelete: Components.IButton = null;

                            // Render the buttons
                            let tooltips = Components.TooltipGroup({
                                el,
                                tooltips: [
                                    {
                                        content: "View Document",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            iconClassName: "mx-1",
                                            iconType: fileEarmark,
                                            iconSize: 24,
                                            text: "View",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // View the document
                                                window.open(Documents.isWopi(`${row.FileName}`) ? row.WebUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc=" + row.Path + "&action=view" : row.Path, "_blank");
                                            }
                                        }
                                    }
                                ]
                            });
                        }
                    }
                ]
            }
        });
    }

    // Runs the report
    static run(el: HTMLElement, auditOnly: boolean, values: { [key: string]: string }, onClose: () => void) {
        let data: IWebItem[] = [];

        // Clear the items
        this._items = [];

        // Show a loading dialog
        LoadingDialog.setHeader("Searching Site");
        LoadingDialog.setBody("Loading the libraries...");
        LoadingDialog.show();

        // Parse the webs
        let counter = 0;
        Helper.Executor(DataSource.SiteItems, siteItem => {
            // Update the dialog
            LoadingDialog.setHeader(`Searching Site ${++counter} of ${DataSource.SiteItems.length}`);

            // Return a promise
            return new Promise(resolve => {
                // Get the libraries for this site
                Web(siteItem.text, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists().query({
                    Filter: `Hidden eq false and BaseTemplate eq ${SPTypes.ListTemplateType.DocumentLibrary} or BaseTemplate eq ${SPTypes.ListTemplateType.PageLibrary}`,
                    Expand: ["RootFolder"],
                    GetAllItems: true,
                    Select: ["Id", "Title", "RootFolder/ServerRelativeUrl"],
                    Top: 5000
                }).execute(libs => {
                    // Add the libraries to analyze
                    libs.results.forEach(lib => {
                        // Append the item
                        data.push({
                            ListId: lib.Id,
                            ListTitle: lib.Title,
                            WebId: siteItem.value,
                            WebUrl: siteItem.text
                        });
                    });

                    // Update the dialog
                    LoadingDialog.setBody("Loading the files for the libraries...");

                    // Analyze the libraries
                    this.analyzeLibraries(siteItem.value, siteItem.text, libs.results).then(resolve);
                });
            });
        }).then(() => {
            // Clear the element
            while (el.firstChild) { el.removeChild(el.firstChild); }

            // Render the summary
            this.renderSummary(el, auditOnly, onClose);

            // Hide the loading dialog
            LoadingDialog.hide();
        });
    }
}