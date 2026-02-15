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
    private static _dashboard: Dashboard = null;
    private static _elSubNav: HTMLElement = null;
    private static _items: IDLPItem[] = [];

    // Gets the form fields to display
    static getFormFields(fileExt: string = ""): Components.IFormControlProps[] {
        return [
            {
                label: "File Types",
                name: "FileTypes",
                className: "mb-3",
                type: Components.FormControlTypes.TextField,
                value: fileExt
            }
        ];
    }

    // Analyzes a single library
    static analyzeLibrary(webId: string, webUrl: string, libId: string, libTitle: string) {
        // Clear the items
        this._items = [];

        // Set the modal
        Modal.clear();
        Modal.setHeader("Data Loss Prevention Report");

        // Show the results
        this.renderSummary(Modal.BodyElement, false, false);

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

        // Update the status
        this._elSubNav.children[0].innerHTML = "Analyzing Library";
        this._elSubNav.children[1].innerHTML = "Getting all files in this library...";

        // Create the list for the batch requests
        let batchRequests = 0;
        let completed = 0;
        let list = Web(webUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists().getById(libId);

        // Get the item ids for this library
        let itemCounter = 0;
        DataSource.loadItems({
            webUrl,
            listId: libId,
            query: {
                Expand: ["Author"],
                Select: ["Author/Title", "FileLeafRef", "FileRef", "File_x0020_Type", "Id"],
            },
            onItem: item => {
                // Create a batch request to get the dlp policy on this item
                list.Items(item.Id).GetDlpPolicyTip().batch(result => {
                    // Ensure a policy exists
                    if (typeof (result["GetDlpPolicyTip"]) === "undefined") {
                        // Parse the conditions
                        result.MatchedConditionDescriptions.results.forEach(condition => {
                            let dataItem: IDLPItem = {
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
                            };

                            // Append the data
                            this._items.push(dataItem);
                            this._dashboard.Datatable.addRow(dataItem);
                        });
                    }

                    // Increment the counter and update the dialog
                    this._elSubNav.children[1].innerHTML = `Batch Requests Processed ${++completed} of ${batchRequests}...`;
                }, batchRequests++ % 25 == 0);

                // Update the dialog
                this._elSubNav.children[1].innerHTML = `Creating Batch Requests - Processed ${++itemCounter} items...`;
            }
        }).then(() => {
            // Update the dialog
            this._elSubNav.children[1].innerHTML = `Executing Batch Request for ${batchRequests} items...`;

            // Execute the batch request
            list.execute(() => {
                // Hide the sub-nav
                this._elSubNav.classList.add("d-none");
            });
        });
    }

    // Analyzes the libraries
    private static analyzeLibraries(webId: string, webUrl: string, libraries: Types.SP.ListOData[], fileExtensions: string[]) {
        // Return a promise
        return new Promise(resolve => {
            let counter = 0;
            let siteText = this._elSubNav.children[0].innerHTML;

            // Parse the libraries
            Helper.Executor(libraries, lib => {
                // Update the dialog
                this._elSubNav.children[0].innerHTML = `${siteText} [Analyzing Library ${++counter} of ${libraries.length}]: ${lib.Title}`;

                // Return a promise
                return new Promise(resolve => {
                    let batchRequests = 0;
                    let completed = 0;

                    // Set the list
                    let list = Web(webUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists(lib.Title);

                    // Get the item ids for this library
                    let itemCounter = 0;
                    DataSource.loadItems({
                        webUrl,
                        listId: lib.Id,
                        query: {
                            Expand: ["Author"],
                            Select: ["Author/Title", "FileLeafRef", "FileRef", "File_x0020_Type", "Id"]
                        },
                        onItem: item => {
                            let analyzeFile = true;

                            // See if the file extensions are provided
                            if (fileExtensions) {
                                // Default the flag
                                analyzeFile = false

                                // Loop through the file extensions
                                fileExtensions.forEach(fileExt => {
                                    // Set the flag if there is match
                                    if (fileExt.toLowerCase() == item["File_x0020_Type"]?.toLowerCase()) { analyzeFile = true; }
                                });
                            }

                            // See if we are analyzing this file
                            if (analyzeFile) {
                                // Create a batch request to get the dlp policy on this item
                                list.Items(item.Id).GetDlpPolicyTip().batch(result => {
                                    // Ensure a policy exists
                                    if (typeof (result["GetDlpPolicyTip"]) === "undefined") {
                                        // Parse the conditions
                                        result.MatchedConditionDescriptions.results.forEach(condition => {
                                            let dataItem: IDLPItem = {
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
                                            };

                                            // Append the data
                                            this._items.push(dataItem);
                                            this._dashboard.Datatable.addRow(dataItem);
                                        });
                                    }

                                    // Increment the counter and update the dialog
                                    this._elSubNav.children[1].innerHTML = `Batch Requests Processed ${++completed} of ${batchRequests}...`;
                                }, batchRequests++ % 25 == 0);
                            }

                            // Update the dialog
                            this._elSubNav.children[1].innerHTML = `Creating Batch Requests - Processed ${++itemCounter} items...`;
                        }
                    }).then(() => {
                        // Update the dialog
                        this._elSubNav.children[1].innerHTML = `Executing Batch Request for ${batchRequests} items...`;

                        // Execute the batch request
                        list.execute(resolve);
                    }, resolve);
                });
            }).then(resolve);
        });
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, auditOnly: boolean, showSearch?: boolean, onClose?: () => void) {
        // Render the summary
        this._dashboard = new Dashboard({
            el,
            navigation: {
                title: "DLP Report",
                showFilter: false,
                items: showSearch ? [{
                    text: "New Search",
                    className: "btn-outline-light",
                    isButton: true,
                    onClick: () => {
                        // Call the close event
                        onClose();
                    }
                }] : null,
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

        // Set the sub-nav element
        this._elSubNav = el.querySelector("#sub-navigation");
        this._elSubNav.classList.remove("d-none");
        this._elSubNav.classList.add("my-2");
        this._elSubNav.innerHTML = `<div class="h6"></div><div></div>`;
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

        // Get the file extensions
        let fileExtensions: string[] = values["FileTypes"] ? values["FileTypes"].trim().split(' ') : [];

        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Render the summary
        this.renderSummary(el, auditOnly, true, onClose);

        // Hide the loading dialog
        LoadingDialog.hide();

        // Determine the webs to target
        let siteItems: Components.IDropdownItem[] = values["TargetWeb"] && values["TargetWeb"]["value"] ? [values["TargetWeb"]] as any : DataSource.SiteItems;

        // Parse the webs
        let counter = 0;
        Helper.Executor(siteItems, siteItem => {
            // Update the status
            this._elSubNav.children[0].innerHTML = `Searching Site ${++counter} of ${siteItems.length}`;

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
                    this._elSubNav.children[1].innerHTML = "Loading the files for the libraries...";

                    // Analyze the libraries
                    this.analyzeLibraries(siteItem.value, siteItem.text, libs.results, fileExtensions).then(resolve);
                });
            });
        }).then(() => {
            // Hide the sub-nav
            this._elSubNav.classList.add("d-none");
        });
    }
}