import { Dashboard, Documents, LoadingDialog, Modal } from "dattatable";
import { Components, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { fileEarmark } from "gd-sprest-bs/build/icons/svgs/fileEarmark";
import * as moment from "moment";
import { DataSource } from "../ds";
import Strings from "../strings";
import { ExportCSV } from "./exportCSV";

interface IDLPItem {
    AppliedActionsText: string;
    Author: string;
    ConditionDescription: string;
    FileExtension: string;
    FileName: string;
    GeneralText: string;
    LastProcessedTime: string;
    ListId: string;
    ListTitle: string;
    Path: string;
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
    "WebUrl",
    "ListTitle",
    "FileName",
    "FileExtension",
    "GeneralText",
    "AppliedActionsText",
    "ConditionDescription",
    "LastProcessedTime",
    "Author",
    "Path",
    "ListId",
    "WebId"
]

export class DLP {
    private static _items: IDLPItem[] = [];

    // Gets the form fields to display
    static getFormFields(): Components.IFormControlProps[] { return []; }

    // Analyzes a single library
    static analyzeLibrary(webId: string, webUrl: string, libId: string, libTitle: string) {
        // Clear the items
        this._items = [];

        // Show a loading dialog
        LoadingDialog.setHeader("Analyzing Library");
        LoadingDialog.setBody("Getting all files in this library...");
        LoadingDialog.show();

        // Get the item ids for this library
        Web(webUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists(libTitle).Items().query({
            Expand: ["Author"],
            GetAllItems: true,
            Select: ["Author/Title", "FileLeafRef", "FileRef", "File_x0020_Type", "Id"],
            Top: 5000
        }).execute(items => {
            let batchRequests = 0;

            // Update the dialog
            LoadingDialog.setBody("Creating batch job for files...");

            // Parse the items and create the batch job
            let list = Web(webUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists(libTitle);
            items.results.forEach(item => {
                // Increment the counter
                batchRequests++;

                // Create a batch request to get the dlp policy on this item
                list.Items(item.Id).GetDlpPolicyTip().batch(result => {
                    // Ensure a policy exists
                    if (typeof (result["GetDlpPolicyTip"]) === "undefined") {
                        // Parse the conditions
                        result.MatchedConditionDescriptions.results.forEach(condition => {
                            // Append the data
                            this._items.push({
                                AppliedActionsText: result.AppliedActionsText,
                                Author: item["Author"]?.Title,
                                ConditionDescription: condition,
                                FileExtension: item["File_x0020_Type"],
                                FileName: item["FileLeafRef"],
                                GeneralText: result.GeneralText,
                                LastProcessedTime: result.LastProcessedTime,
                                ListId: libId,
                                ListTitle: libTitle,
                                Path: item["FileRef"],
                                WebId: webId,
                                WebUrl: webUrl
                            });
                        });
                    }
                });
            });

            // Update the dialog
            LoadingDialog.setBody(`Executing Batch Request for ${batchRequests} items...`);

            // Execute the batch request
            list.execute(() => {
                // Set the modal
                Modal.clear();
                Modal.setHeader("Data Loss Prevention Report");

                // Show the results
                this.renderSummary(Modal.BodyElement, false, null);

                // Render the footer
                Components.ButtonGroup({
                    el: Modal.FooterElement,
                    buttons: [
                        {
                            text: "Close",
                            type: Components.ButtonTypes.OutlinePrimary,
                            onClick: () => { Modal.hide(); }
                        }
                    ]
                });

                // Show the modal
                Modal.show();

                // Hide the dialog
                LoadingDialog.hide();
            });
        });
    }

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
                    let batchRequests = 0;

                    // Get the item ids for this library
                    Web(webUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists(lib.Title).Items().query({
                        Expand: ["Author"],
                        GetAllItems: true,
                        Select: ["Author/Title", "FileLeafRef", "FileRef", "File_x0020_Type", "Id"],
                        Top: 5000
                    }).execute(items => {
                        let list = Web(webUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists(lib.Title);

                        // Parse the items and create the batch job
                        items.results.forEach(item => {
                            // Increment the counter
                            batchRequests++;

                            // Create a batch request to get the dlp policy on this item
                            list.Items(item.Id).GetDlpPolicyTip().batch(result => {
                                // Ensure a policy exists
                                if (typeof (result["GetDlpPolicyTip"]) === "undefined") {
                                    // Parse the conditions
                                    result.MatchedConditionDescriptions.results.forEach(condition => {
                                        // Append the data
                                        this._items.push({
                                            AppliedActionsText: result.AppliedActionsText,
                                            Author: item["Author"]?.Title,
                                            ConditionDescription: condition,
                                            FileExtension: item["File_x0020_Type"],
                                            FileName: item["FileLeafRef"],
                                            GeneralText: result.GeneralText,
                                            LastProcessedTime: result.LastProcessedTime,
                                            ListId: lib.Id,
                                            ListTitle: lib.Title,
                                            Path: item["FileRef"],
                                            WebId: webId,
                                            WebUrl: webUrl
                                        });
                                    });
                                }
                            });
                        });

                        // Update the dialog
                        LoadingDialog.setBody(`Executing Batch Request for ${batchRequests} items...`);

                        // Execute the batch request
                        list.execute(resolve);
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
                title: "DLP Report",
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
                        new ExportCSV("dlpReport.csv", CSVFields, this._items);
                    }
                }]
            },
            table: {
                rows: this._items,
                onRendering: dtProps => {
                    dtProps.columnDefs = [
                        {
                            "targets": 5,
                            "orderable": false,
                            "searchable": false
                        }
                    ];

                    // Order by the 1st column by default; ascending
                    dtProps.order = [[2, "asc"]];

                    // Return the properties
                    return dtProps;
                },
                columns: [
                    {
                        name: "ListTitle",
                        title: "List"
                    },
                    {
                        name: "Author",
                        title: "Created By"
                    },
                    {
                        name: "",
                        title: "File",
                        onRenderCell: (el, col, item: IDLPItem) => {
                            // Set the sort value
                            el.setAttribute("data-order", item.Path);

                            // Show the file info
                            el.innerHTML = `
                                <b>Name: </b>${item.FileName}
                                <br/>
                                <b>Path: </b>${item.Path}
                            `;
                        }
                    },
                    {
                        name: "ConditionDescription",
                        title: "Condition"
                    },
                    {
                        name: "LastProcessedTime",
                        title: "Last Processed Time",
                        onRenderCell: (el, col, item: IDLPItem) => {
                            el.innerHTML = item.LastProcessedTime ? moment(item.LastProcessedTime).format(Strings.TimeFormat) : "";
                        }
                    },
                    {
                        className: "text-end",
                        name: "",
                        title: "",
                        onRenderCell: (el, col, row: IDLPItem) => {
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
                    GetAllItems: true,
                    Select: ["Id", "Title"],
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