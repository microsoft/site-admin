import { Dashboard, Documents, LoadingDialog } from "dattatable";
import { Components, ContextInfo, Helper, Search, SPTypes, Types, Web, v2 } from "gd-sprest-bs";
import { Workbook } from "exceljs";
import { extractRawText } from "mammoth";
import * as moment from "moment";
import { PDFParse } from "pdf-parse";
import { DataSource } from "../ds";
import Strings from "../strings";
import { ExportCSV } from "./exportCSV";
import { SensitivityLabels } from "./sensitivityLabels";

interface ISearchItem {
    _driveItem?: Types.Microsoft.Graph.driveItem;
    Author: string;
    FileExtension: string;
    FileUrl: string;
    HitHighlightedSummary?: string;
    LastModifiedTime: string;
    ListId: string;
    Path: string;
    RegexPattern?: string;
    SensitivityLabel?: string;
    SensitivityLabelId?: string;
    SPSiteUrl: string;
    SPWebUrl: string;
    Title: string;
    ViewUrl: string;
    WebId: string;
}

const CSVFields = [
    "Author", "FileExtension", "HitHighlightedSummary", "RegexPattern", "LastModifiedTime",
    "SensitivityLabel", "SensitivityLabelId", "ListId", "Path", "SPSiteUrl", "SPWebUrl", "Title", "WebId"
]

// The valid file extensions for regex patterns
const FileExtensions = ["csv", "docx", "pptx", "pdf", "txt", "xlsx"];

export class SearchDocs {
    private static _dashboard: Dashboard = null;
    private static _elSubNav: HTMLElement = null;
    private static _items: ISearchItem[] = [];
    private static _loadOneDrive: boolean = false;
    private static _stopFl: boolean = false;

    // Analyzes the file
    private static analyzeFile(item: Types.Microsoft.Graph.driveItem, driveUrl: string, webUrl: string, webId: string, regexPatterns: RegExp[]) {
        // Return a promise
        return new Promise(resolve => {
            // Update the dialog
            this._elSubNav.children[1].innerHTML = `Processing File: ${item.name}`;

            // Convert the buffer to string
            let getContent = (buffer: ArrayBuffer): PromiseLike<string> => {
                return new Promise((resolve, reject) => {
                    switch (item.file["fileExtension"].substring(1)) {
                        case "docx":
                            // Get the content
                            extractRawText({ arrayBuffer: new Uint8Array(buffer) as any }).then(result => {
                                resolve(result.value);
                            }).catch(reject);
                            break;
                        case "pdf":
                            // Get the content
                            let pdf = new PDFParse({ data: buffer });
                            pdf.getText().then(content => { resolve(content.text); }).catch(reject);
                            break;
                        case "xlsx":
                            (new Workbook()).xlsx.load(buffer).then(workbook => {
                                let content = [];
                                // Parse the sheets
                                workbook.eachSheet(sheet => {
                                    // Parse each row
                                    sheet.eachRow(row => {
                                        // Append the content
                                        content.push(row.values.toString());
                                    });
                                });
                                resolve(content.join("\n"));
                            }).catch(reject);
                            break;
                        default:
                            // Convert the buffer to a string
                            let decoder = new TextDecoder("utf-8");
                            resolve(decoder.decode(buffer));
                            break;
                    }
                });
            }

            // Get the content of the file
            v2.drive({
                driveId: item.parentReference.driveId,
                siteId: item.parentReference.siteId,
                webId
            }).items(item.id)["content"]().execute(buffer => {
                // Get the content
                getContent(buffer).then(content => {
                    let patterns = [];

                    // See if the regex pattern exists in the content
                    regexPatterns.forEach(regexPatterns => {
                        // Test the pattern
                        if (regexPatterns.test(content)) {
                            // Add the pattern
                            patterns.push(regexPatterns.source);
                        }
                    });

                    // See if patterns were found
                    if (patterns.length > 0) {
                        // Add the item
                        let itemInfo: ISearchItem = {
                            _driveItem: item,
                            Author: item.createdBy.user["email"],
                            FileExtension: item.file["fileExtension"].substring(1),
                            FileUrl: item.parentReference.path.split("/root:").pop() + "/" + item.name,
                            LastModifiedTime: item.fileSystemInfo.lastModifiedDateTime,
                            ListId: item.parentReference.driveId,
                            Path: driveUrl + item.parentReference.path.split("/root:").pop(),
                            RegexPattern: patterns.join(", "),
                            SensitivityLabel: item.sensitivityLabel?.displayName,
                            SensitivityLabelId: item.sensitivityLabel?.id,
                            SPSiteUrl: item.parentReference.path,
                            SPWebUrl: webUrl,
                            Title: item.name,
                            ViewUrl: item.webUrl,
                            WebId: webId,
                        }

                        // Add it to the dashboard
                        this._items.push(itemInfo);
                        this._dashboard.Datatable.addRow(itemInfo);
                    }

                    // Resolve the request
                    resolve(null);
                }, () => {
                    // Add an error
                    // TODO

                    // Resolve the request
                    resolve(null);
                });
            });
        });
    }

    // Analyzes the libraries of a site
    private static analyzeLibraries(webId: string, webUrl: string, drives: Types.Microsoft.Graph.drive[], fileExt: string[], regexPatterns: RegExp[]) {
        // Return a promise
        return new Promise(resolve => {
            // Set the completed event
            let onCompleted = () => {
                // Clear the sub-nav
                this._elSubNav.classList.add("d-none");

                // Clear the callback events
                ContextInfo.clearRateLimitCallbacks();

                // Resolve the request
                resolve(null);
            };

            // File counters for processing
            let fileCounter = 0;
            let filesToProcess: Types.Microsoft.Graph.driveItem[] = [];
            let processingCounter = 0;
            let processedCounter = 0;

            // Create a worker process
            let worker = Helper.WebWorker(() => {
                // See if we are stopping this process
                if (this._stopFl) {
                    // Stop the process
                    worker.stop();
                }

                // Do nothing if we are processing the max files at once
                if (processingCounter >= Strings.MaxRequests) { return; }

                // Do nothing if we are done
                if (filesToProcess.length == 0) {
                    // Stop the process
                    worker.stop();

                    // Call the event
                    onCompleted ? onCompleted() : null;

                    // Do nothing
                    return;
                }

                // Increment the # of files being processed
                processingCounter++;

                // Analyze the file
                let file = filesToProcess.splice(0, 1).pop();
                this.analyzeFile(file, file.parentReference["driveUrl"], webUrl, webId, regexPatterns).then(() => {
                    // Update the dialog
                    this._elSubNav.children[1].innerHTML = `[Processed ${++processedCounter} of ${fileCounter}] File Labelled: ${file.name}`;

                    // Decrement the # of files being processed
                    processingCounter--;
                });
            }, 100);

            // Parse the libraries
            Helper.Executor(drives, drive => {
                // Set the status
                this._elSubNav.children[0].innerHTML = `Analyzing Library: '${drive.name}'`;
                this._elSubNav.children[1].innerHTML = `Loading the files for this library...`;

                // Parse the file extensions to target
                let filters = [];
                (fileExt || FileExtensions).forEach(ext => {
                    filters.push(`substringof('.${ext}', FileLeafRef)`);
                });

                // Load the files for this drive
                DataSource.loadFiles(webId, webUrl, drive.id, drive.name, null, null, item => {
                    // Ensure the file extension is valid
                    if ((fileExt || FileExtensions).indexOf(item.file["fileExtension"].substring(1)) < 0) { return; }

                    // Add the drive reference
                    item.parentReference["driveUrl"] = drive.webUrl;

                    // Add the file to process
                    filesToProcess.push(item);

                    // Ensure the process is running
                    worker.start();
                });
            }).then(() => {
                // Update the dialog
                this._elSubNav.children[0].innerHTML = `All Files Loaded. Waiting for the processing to complete.`;

                // Ensure the process is running
                worker.start();
            });
        });
    }

    // Deletes a document
    private static deleteDocument(item: ISearchItem) {
        // Display a loading dialog
        LoadingDialog.setHeader("Deleting Document");
        LoadingDialog.setBody("Deleting the document '" + item.Title + "'. This will close after the request completes.");
        LoadingDialog.show();

        // Delete the document
        Web(item.SPWebUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).getFileByServerRelativeUrl(item.Path).delete().execute(
            // Success
            () => {
                // TODO - Display the confirmation

                // Close the dialog
                LoadingDialog.hide();
            },

            // Error
            () => {
                // TODO - Display an error

                // Close the dialog
                LoadingDialog.hide();
            }
        )
    }

    // Gets the form fields to display
    static getFormFields(fileExt: string = "", keywords: string = "", regexPatterns: string = ""): Components.IFormControlProps[] {
        let ctrlRegex: Components.IFormControl;
        let ctrlSearchTerms: Components.IFormControl;
        let ctrlSearchType: Components.IFormControl;
        return [
            {
                label: "Search Type",
                name: "SearchType",
                className: "mb-3",
                type: Components.FormControlTypes.Dropdown,
                required: true,
                items: [
                    { text: "Keyword", value: "Keyword", isSelected: true },
                    { text: "Regex Pattern", value: "RegexPattern" }
                ],
                onControlRendered: ctrl => { ctrlSearchType = ctrl; },
                onChange: (item) => {
                    // See which one is selected
                    if (item.value == "Keyword") {
                        // Show/Hide the controls
                        ctrlRegex.hide();
                        ctrlSearchTerms.show();
                    } else {
                        // Show/Hide the controls
                        ctrlRegex.show();
                        ctrlSearchTerms.hide();
                    }
                }
            } as Components.IFormControlPropsDropdown,
            {
                label: "Search Terms",
                name: "SearchTerms",
                className: "mb-3",
                description: "Enter the search terms using quotes for phrases [Ex: movie \"social media\" show]",
                type: Components.FormControlTypes.TextField,
                required: true,
                value: keywords,
                errorMessage: "You must enter at least 1 search term.",
                onControlRendered: ctrl => { ctrlSearchTerms = ctrl; }
            },
            {
                label: "Regex Pattern",
                name: "RegexPattern",
                className: "mb-3 d-none",
                description: "Enter the regular expression pattern to search for.",
                type: Components.FormControlTypes.TextField,
                required: true,
                value: regexPatterns,
                errorMessage: "You must enter at least 1 search term.",
                onControlRendered: ctrl => { ctrlRegex = ctrl; },
                onValidate: (ctrl, results) => {
                    // Ensure a pattern exists
                    if ((results.value || "").trim().length == 0) {
                        // Invalidate the extensions
                        results.isValid = false;
                        results.invalidMessage = "A regex pattern is required.";
                    } else {
                        // Validate the regex pattern
                        try {
                            // Create a new regex
                            new RegExp(results.value);
                        } catch (ex) {
                            // Invalidate the regex pattern
                            results.isValid = false;
                            results.invalidMessage = "The regex pattern is not valid.";
                        }
                    }

                    // Return the results
                    return results;
                }
            },
            {
                label: "File Types",
                name: "FileTypes",
                className: "mb-3",
                type: Components.FormControlTypes.TextField,
                value: fileExt
            }
        ];
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, auditOnly: boolean, isSearch: boolean, onClose: () => void) {
        // Render the summary
        this._dashboard = new Dashboard({
            el,
            navigation: {
                title: "Search Documents",
                showFilter: false,
                items: [{
                    text: "New Search",
                    className: "btn-outline-light",
                    isButton: true,
                    onClick: () => {
                        // Call the close event
                        onClose();
                    }
                }, {
                    text: "Bulk Label",
                    className: isSearch ? "d-none" : "btn-outline-light ms-2",
                    isButton: true,
                    onClick: () => {
                        // Get the drive items
                        let driveItems = [];
                        this._items.forEach(item => { driveItems.push(item._driveItem); });

                        // Show the label form
                        SensitivityLabels.showLabelFilesForm(driveItems, responses => {
                            // Update the items
                        });
                    }
                }],
                itemsEnd: [{
                    text: "Export to CSV",
                    className: "btn-outline-light me-2",
                    isButton: true,
                    onClick: () => {
                        // Export the CSV
                        new ExportCSV("searchDocs.csv", CSVFields, this._items);
                    }
                }]
            },
            table: {
                rows: this._items,
                onRendering: dtProps => {
                    dtProps.columnDefs = [
                        {
                            "targets": 3,
                            "orderable": false,
                            "searchable": false
                        }
                    ];

                    // Order by the 1st column by default; ascending
                    dtProps.order = [[0, "asc"]];

                    // Return the properties
                    return dtProps;
                },
                columns: [
                    {
                        name: "",
                        title: "File Name",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            el.innerHTML = `
                                <small class="text-muted">File Name: </small>${item.Title + (isSearch ? "." + item.FileExtension : "")}
                                <br/>
                                <small class="text-muted">Path: </small>${item.FileUrl || item.Path}
                            `;
                        }
                    },
                    {
                        name: "",
                        title: isSearch ? "Last Modified" : "Sensitivity Label",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            if (isSearch) {
                                // Set the last modified time
                                el.innerHTML = item.LastModifiedTime ? moment(item.LastModifiedTime).format(Strings.TimeFormat) : "";
                            } else {
                                // Set the sensitivity label
                                el.innerHTML = item.SensitivityLabel || "";
                            }
                        }
                    },
                    {
                        name: "",
                        title: isSearch ? "Highlighted Summary" : "Pattern Matches",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            // Declare a span element
                            let span = document.createElement("span");

                            // See if we were searching by keyword
                            if (isSearch) {
                                // Return the plain text if less than 50 chars
                                if (el.innerHTML.length < 50) {
                                    span.innerHTML = item.HitHighlightedSummary;
                                } else {
                                    // Truncate to the last white space character in the text after 50 chars and add an ellipsis
                                    span.innerHTML = item.HitHighlightedSummary.substring(0, 50).replace(/\s([^\s]*)$/, '') + '&#8230';

                                    // Add a tooltip containing the text
                                    Components.Tooltip({
                                        content: "<small>" + item.HitHighlightedSummary + "</small>",
                                        target: span
                                    });
                                }
                            } else {
                                span.innerHTML = item.RegexPattern;
                            }

                            // Append the span
                            el.appendChild(span);
                        }
                    },
                    {
                        className: "text-end",
                        name: "",
                        title: "",
                        onRenderCell: (el, col, item: ISearchItem, rowIdx) => {
                            let btnDelete: Components.IButton = null;

                            // Render the buttons
                            let tooltips = Components.TooltipGroup({ el });

                            // See if this is from search
                            if (isSearch) {
                                // Add a view button
                                tooltips.add({
                                    content: "Click to view the document.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        text: "View File",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        isSmall: true,
                                        onClick: () => {
                                            // View the file
                                            window.open(Documents.isWopi(`${item.Title}.${item.FileExtension}`) ? item.SPWebUrl + "/_layouts/15/WopiFrame.aspx?sourcedoc=" + item.Path + "&action=view" : item.Path, "_blank");
                                        }
                                    }
                                });
                            } else {
                                // Add a view button                                
                                tooltips.add({
                                    content: "Click to view the document.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        text: "View File",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        isSmall: true,
                                        onClick: () => {
                                            // View the file
                                            window.open(item.ViewUrl, "_blank");
                                        }
                                    }
                                });
                            }

                            // Add a download button
                            tooltips.add({
                                content: "Click to download the document.",
                                btnProps: {
                                    className: "pe-2 py-1",
                                    text: "Download",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    isSmall: true,
                                    onClick: () => {
                                        // Download the document
                                        window.open(`${item.SPWebUrl}/_layouts/15/download.aspx?SourceUrl=${item.Path}`, "_blank");
                                    }
                                }
                            });

                            // Add the option to delete
                            if (!auditOnly) {
                                // Add a sensitivity label button
                                tooltips.add({
                                    content: "Click to set the sensitivity label on the file.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        text: "Set Label",
                                        isSmall: true,
                                        onClick: () => {
                                            SensitivityLabels.showLabelFileForm(item._driveItem, label => {
                                                // Update the row cell
                                                item.SensitivityLabel = label;
                                                this._dashboard.updateCell(rowIdx, 1, item);
                                            });
                                        }
                                    }
                                });

                                // Add a delete button
                                tooltips.add({
                                    content: "Click to delete the document.",
                                    btnProps: {
                                        assignTo: btn => { btnDelete = btn; },
                                        className: "pe-2 py-1",
                                        text: "Delete",
                                        type: Components.ButtonTypes.OutlineDanger,
                                        isSmall: true,
                                        onClick: () => {
                                            // Confirm the deletion of the group
                                            if (confirm("Are you sure you want to delete this document?")) {
                                                // Disable this button
                                                btnDelete.disable();

                                                // Delete the document
                                                this.deleteDocument(item);
                                            }
                                        }
                                    }
                                });
                            }
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
        this._items = [];
        this._loadOneDrive = values["LoadOneDrive"] == "true";

        // Show a loading dialog
        LoadingDialog.setHeader("Searching Site");
        LoadingDialog.setBody("Searching the content on this site...");
        LoadingDialog.show();

        // Get the form values
        let fileExt = values["FileTypes"] ? values["FileTypes"].split(' ') : null;
        let searchTerms = (values["SearchTerms"] || "").split(' ');
        let searchType = (values["SearchType"] || "");

        // Set the regex patterns
        let regexPatterns = [];
        (values["RegexPattern"]?.split(' ') || []).forEach(pattern => {
            regexPatterns.push(new RegExp(pattern));
        });

        // Set the query
        let query: Types.Microsoft.Office.Server.Search.REST.SearchRequest = {
            Querytext: `${searchTerms.join(" OR ")} IsDocument: true path: ${this._loadOneDrive ? DataSource.OneDriveWeb.Url : DataSource.SiteContext.SiteFullUrl}`,
            RowLimit: 500,
            SelectProperties: {
                results: [
                    "Author", "FileExtension", "HitHighlightedSummary", "LastModifiedTime",
                    "ListId", "Path", "SPSiteUrl", "SPWebUrl", "Title", "WebId"
                ]
            }
        };

        // See if file extensions exist
        if (fileExt) {
            // Set the filter
            query.RefinementFilters = {
                results: [`fileExtension:or("${fileExt.join('", "')}")`]
            };
        }

        // Set the url
        let url = this._loadOneDrive ? ContextInfo.siteAbsoluteUrl : DataSource.SiteContext.SiteFullUrl;

        // See if we are doing a keyword search
        if (searchType == "Keyword") {
            // Search for the content
            Search.postQuery({
                getAllItems: true,
                url,
                targetInfo: { requestDigest: this._loadOneDrive ? ContextInfo.formDigestValue : DataSource.SiteContext.FormDigestValue },
                query
            }).then(search => {
                // Clear the element
                while (el.firstChild) { el.removeChild(el.firstChild); }

                // Set the items
                this._items = search.results;

                // Render the summary
                this.renderSummary(el, auditOnly, true, onClose);

                // Hide the sub-nav
                this._elSubNav.classList.add("d-none");

                // Hide the loading dialog
                LoadingDialog.hide();
            });
        } else {
            // Show a loading dialog
            LoadingDialog.setHeader("Searching Site");
            LoadingDialog.setBody("Loading the libraries...");
            LoadingDialog.show();

            // Clear the element
            while (el.firstChild) { el.removeChild(el.firstChild); }

            // Render the summary
            this.renderSummary(el, auditOnly, false, onClose);

            // Hide the loading dialog
            LoadingDialog.hide();

            // Determine the webs to target
            let siteItems: Components.IDropdownItem[] = null;
            if (this._loadOneDrive) {
                siteItems = [{ text: DataSource.OneDriveWeb.Url, value: DataSource.OneDriveWeb.Id }] as any;
            } else {
                siteItems = values["TargetWeb"] && values["TargetWeb"]["value"] ? [values["TargetWeb"]] as any : DataSource.SiteItems;
            }

            // Parse the webs
            let counter = 0;
            Helper.Executor(siteItems, siteItem => {
                // See if we are stopping this process
                if (this._stopFl) { return; }

                // Update the status
                this._elSubNav.children[0].innerHTML = `Searching Site ${++counter} of ${siteItems.length}`;
                this._elSubNav.children[1].innerHTML = "Getting the libraries for this web...";

                // Return a promise
                return new Promise(resolve => {
                    // Set the default filter
                    let filter = `Hidden eq false and BaseTemplate eq ${SPTypes.ListTemplateType.DocumentLibrary} or BaseTemplate eq ${SPTypes.ListTemplateType.MySiteDocumentLibrary} or BaseTemplate eq ${SPTypes.ListTemplateType.PageLibrary}`;

                    // See if a list was specified
                    if (values["TargetList"]) {
                        // Set the filter
                        filter = `Title eq '${values["TargetList"]}'`;
                    }

                    // Get the drives for this web
                    v2.drives({
                        siteId: this._loadOneDrive ? DataSource.OneDriveSite.Id : DataSource.Site.Id,
                        webId: this._loadOneDrive ? DataSource.OneDriveWeb.Id : DataSource.Web.Id
                    }).execute(drives => {
                        // Update the dialog
                        this._elSubNav.children[1].innerHTML = "Loading the files for the libraries...";

                        // Analyze the libraries
                        return this.analyzeLibraries(siteItem.value, siteItem.text, drives.results, fileExt, regexPatterns);
                    });
                });
            }).then(() => {
                // Hide the sub-nav
                this._elSubNav.classList.add("d-none");
            });
        }
    }
}