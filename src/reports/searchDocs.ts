import { CanvasForm, Dashboard, Documents, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, Search, SPTypes, Types, Web, v2 } from "gd-sprest-bs";
import { Workbook } from "exceljs";
import { loadAsync } from "jszip";
import { extractRawText } from "mammoth";
import * as moment from "moment";
import { PDFParse } from "pdf-parse";
import { DataSource } from "../ds";
import { M365Groups } from "../m365Groups";
import { IRegexPattern } from "../regexPatternsDialog";
import Strings from "../strings";
import { ExportCSV } from "./exportCSV";
import { SensitivityLabels } from "./sensitivityLabels";
import { ViewPermissions } from "./viewPermissions";

export interface ISearchItem {
    _driveItem?: Types.Microsoft.Graph.driveItem;
    Author: string;
    DriveId: string;
    ErrorExtractingContent?: boolean;
    ErrorMessage?: string;
    FileExtension: string;
    FileUrl: string;
    HasUniquePermissions: boolean;
    HitHighlightedSummary?: string;
    ItemId: number;
    LastModifiedTime: string;
    ListId: string;
    Overshared: string;
    Path: string;
    Permissions: Types.SP.RoleAssignmentOData[];
    RegexPatterns?: string;
    SensitivityLabel?: string;
    SensitivityLabelId?: string;
    SPSiteUrl: string;
    SPWebUrl: string;
    Title: string;
    ViewUrl: string;
    WebId: string;
}

const CSVFields = [
    "Author", "FileExtension", "HitHighlightedSummary", "RegexPatterns", "LastModifiedTime",
    "SensitivityLabel", "SensitivityLabelId", "ListId", "Path", "SPSiteUrl", "SPWebUrl", "Title", "WebId"
]

// The valid file extensions for regex patterns
const FileExtensions = ["csv", "docx", "pptx", "pdf", "txt", "xlsx"];

export class SearchDocs {
    private static _dashboard: Dashboard = null;
    private static _elSubNav: HTMLElement = null;
    private static _items: ISearchItem[] = [];
    private static _itemErrors: ISearchItem[] = [];
    private static _loadOneDrive: boolean = false;
    private static _stopFl: boolean = false;

    // Analyzes the file
    private static analyzeFile(item: Types.Microsoft.Graph.driveItem, driveUrl: string, listId: string, webUrl: string, webId: string, regexPatterns: RegExp[]) {
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
                        case "pptx":
                            // Load the file
                            loadAsync(buffer).then(zip => {
                                let allText: string[] = [];

                                // Get the slides
                                let slides = Object.keys(zip.files)
                                    .filter(file => /^ppt\/slides\/slide\d+\.xml$/i.test(file))
                                    .sort((a, b) => {
                                        let numA = parseInt(a.match(/\d+/)?.[0] ?? "0", 10);
                                        let numB = parseInt(b.match(/\d+/)?.[0] ?? "0", 10);
                                        return numA - numB;
                                    });

                                // Parse the slides
                                Helper.Executor(slides, slide => {
                                    // Return a promise
                                    return new Promise(resolve => {
                                        // Get the text
                                        zip.file(slide)?.async("text").then(xml => {
                                            // Ensure xml exist
                                            if (!xml) { return; }

                                            // Extract the text
                                            const matches = xml.match(/<a:t[^>]*>.*?<\/a:t>/g) || [];
                                            allText.push(matches
                                                .map(m =>
                                                    m
                                                        .replace(/<a:t[^>]*>/, "")
                                                        .replace(/<\/a:t>/, "")
                                                )
                                                .join(" ")
                                            );

                                            // Check the next slide
                                            resolve(null);
                                        }).catch(resolve);
                                    });
                                }).then(() => {
                                    // Resolve the text
                                    resolve(allText.join("\n"));
                                });
                            }).catch(reject);
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
            }).items(item.id)["content"]().execute((buffer: any) => {
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
                            DriveId: item.parentReference.driveId,
                            ErrorExtractingContent: false,
                            FileExtension: item.file["fileExtension"].substring(1),
                            FileUrl: item.parentReference.path.split("/root:").pop() + "/" + item.name,
                            HasUniquePermissions: item?.listItem["HasUniquePermissions"],
                            ItemId: item?.listItem["Id"],
                            LastModifiedTime: item.fileSystemInfo.lastModifiedDateTime,
                            ListId: listId,
                            Overshared: item.listItem["RoleAssignments"] ? (ViewPermissions.isOvershared(item.listItem["RoleAssignments"].results) ? "Yes" : "No") : "",
                            Path: driveUrl + item.parentReference.path.split("/root:").pop(),
                            Permissions: item?.listItem["RoleAssignments"]?.results,
                            RegexPatterns: patterns.join(", "),
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
                    let itemInfo: ISearchItem = {
                        _driveItem: item,
                        Author: item.createdBy.user["email"],
                        DriveId: item.parentReference.driveId,
                        ErrorExtractingContent: true,
                        ErrorMessage: "Unable to extract the text from the file. It may be protected.",
                        FileExtension: item.file["fileExtension"].substring(1),
                        FileUrl: item.parentReference.path.split("/root:").pop() + "/" + item.name,
                        HasUniquePermissions: item?.listItem["HasUniquePermissions"],
                        ItemId: item?.listItem["Id"],
                        LastModifiedTime: item.fileSystemInfo.lastModifiedDateTime,
                        ListId: listId,
                        Overshared: item.listItem["RoleAssignments"] ? (ViewPermissions.isOvershared(item.listItem["RoleAssignments"].results) ? "Yes" : "No") : "",
                        Path: driveUrl + item.parentReference.path.split("/root:").pop(),
                        Permissions: item?.listItem["RoleAssignments"]?.results,
                        SensitivityLabel: item.sensitivityLabel?.displayName,
                        SensitivityLabelId: item.sensitivityLabel?.id,
                        SPSiteUrl: item.parentReference.path,
                        SPWebUrl: webUrl,
                        Title: item.name,
                        ViewUrl: item.webUrl,
                        WebId: webId,
                    }

                    // Add it to the dashboard
                    this._itemErrors.push(itemInfo);

                    // Resolve the request
                    resolve(null);
                });
            }, (ex) => {
                let itemInfo: ISearchItem = {
                    _driveItem: item,
                    Author: item.createdBy.user["email"],
                    DriveId: item.parentReference.driveId,
                    ErrorExtractingContent: true,
                    ErrorMessage: "Unable to download the file. It may be protected.",
                    FileExtension: item.file["fileExtension"].substring(1),
                    FileUrl: item.parentReference.path.split("/root:").pop() + "/" + item.name,
                    HasUniquePermissions: item?.listItem["HasUniquePermissions"],
                    ItemId: item?.listItem["Id"],
                    LastModifiedTime: item.fileSystemInfo.lastModifiedDateTime,
                    ListId: listId,
                    Overshared: item.listItem["RoleAssignments"] ? (ViewPermissions.isOvershared(item.listItem["RoleAssignments"].results) ? "Yes" : "No") : "",
                    Path: driveUrl + item.parentReference.path.split("/root:").pop(),
                    Permissions: item?.listItem["RoleAssignments"]?.results,
                    SensitivityLabel: item.sensitivityLabel?.displayName,
                    SensitivityLabelId: item.sensitivityLabel?.id,
                    SPSiteUrl: item.parentReference.path,
                    SPWebUrl: webUrl,
                    Title: item.name,
                    ViewUrl: item.webUrl,
                    WebId: webId,
                }

                // Add it to the dashboard
                this._itemErrors.push(itemInfo);

                // Resolve the request
                resolve(null);
            });
        });
    }

    // Analyzes the library
    private static analyzeLibrary(webId: string, webUrl: string, library: Types.SP.ListOData, drives: Types.Microsoft.Graph.drive[], folderId: string, loadPermissions: boolean, fileExt: string[], regexPatterns: RegExp[]) {
        // Return a promise
        return new Promise(resolve => {
            // Get the drive for this library
            let drive = drives.find(drive => {
                return drive.name == library.Title || drive.webUrl.endsWith(library.RootFolder.ServerRelativeUrl);
            });

            // Ensure a drive exists, otherwise check the next library
            if (drive == null) {
                resolve(null);
                return;
            }

            // Set the completed event
            let onCompleted = () => {
                // Clear the callback events
                ContextInfo.clearRateLimitCallbacks();

                // Resolve the request
                resolve(null);
            };

            // Subscribe to the rate info event and set the time to sleep before completing the request
            let sleepTime = 0;
            ContextInfo.onRateLimitDetected(rateInfo => {
                // See if we have dropped below a threshold
                if (rateInfo.remaining < Strings.RateLimitThreshold) {
                    // Set the sleep time
                    sleepTime = rateInfo.reset * 1000;

                    // Show a loading dialog
                    LoadingDialog.setHeader("Throttling Detected");
                    LoadingDialog.setBody(`Throttling has been detected. Pausing requests for ${rateInfo.reset} seconds before sending next request...`);
                    LoadingDialog.show();

                    // Wait for the specified time and reset the value
                    setTimeout(() => {
                        // Clear the sleep time and hide the dialog
                        sleepTime = 0;
                        LoadingDialog.hide();
                    }, sleepTime);
                }
            });

            // File counters for processing
            let filesLoaded = 0;
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

                // See if the queue is empty
                if (filesToProcess.length == 0) {
                    // See if we have completed processing the files
                    if (processingCounter == 0 && (this._stopFl || allFilesLoaded)) {
                        // Stop the process
                        worker.stop();

                        // Call the event
                        onCompleted ? onCompleted() : null;
                    }

                    // Do nothing
                    return;
                }

                // Increment the # of files being processed
                processingCounter++;

                // Processes the file
                let processFile = () => {
                    // Get the file
                    let file = filesToProcess.splice(0, 1).pop();

                    // Wait for the specified sleep time to avoid throttling
                    setTimeout(() => {
                        // Analyze the file
                        this.analyzeFile(file, file.parentReference["driveUrl"], library.Id, webUrl, webId, regexPatterns).then(() => {
                            // Update the dialog
                            this._elSubNav.children[0].innerHTML = `Analyzing Library: ${library.Title} - Processed ${++processedCounter} of ${filesLoaded}`;
                            this._elSubNav.children[1].innerHTML = `[${processingCounter} Processing] File Labelled: ${file.name}`;

                            // Decrement the # of files being processed
                            processingCounter--;
                        });
                    }, sleepTime);
                }

                // Process the file
                processFile();
            }, 100);

            // Parse the file extensions to target
            let filters = [];
            (fileExt || FileExtensions).forEach(ext => {
                filters.push(`substringof('.${ext}', FileLeafRef)`);
            });

            // Set the status
            this._elSubNav.children[0].innerHTML = `Analyzing Library: ${library.Title}`;
            this._elSubNav.children[1].innerHTML = `Loading the files for this library...`;

            // Load the files for this drive
            let allFilesLoaded = false;
            return DataSource.loadFiles(webId, webUrl, drive.id, drive.name, folderId, loadPermissions, item => {
                // Ensure the file extension is valid
                if ((fileExt || FileExtensions).indexOf(item.file["fileExtension"].substring(1)) < 0) { return; }

                // Add the drive reference
                item.parentReference["driveUrl"] = drive.webUrl;

                // Add the file to process
                filesToProcess.push(item);
                filesLoaded++;

                // Ensure the process is running
                worker.start();
            }).then(() => {
                // Set the flag
                allFilesLoaded = true;

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
    static getFormFields(fileExt: string = "", keywords: string = "", regexPatterns: string = "", regexOnly: boolean = false): Components.IFormControlProps[] {
        let ctrlPermissions: Components.IFormControl;
        let ctrlRegex: Components.IFormControl;
        let ctrlRegexPatterns: Components.IFormControl;
        let ctrlSearchTerms: Components.IFormControl;

        // Set the items
        let items: Components.IDropdownItem[] = regexOnly ? [{ text: "Regex Pattern", value: "RegexPatterns" }] : [
            { text: "Keyword", value: "Keyword" },
            { text: "Regex Pattern", value: "RegexPatterns", isSelected: true }
        ]

        // Convert the regex patterns to an object
        let regexItems: Components.IDropdownItem[] = [{ text: "Select a Category" }];
        try {
            (JSON.parse(regexPatterns) as IRegexPattern[]).forEach(pattern => {
                // Append the item
                regexItems.push({
                    data: pattern,
                    text: pattern.title
                });
            });
        } catch {
            // Clear the items
            regexItems = [];
        }

        // Return the properties
        return [
            {
                label: "Search Type",
                name: "SearchType",
                className: "mb-3",
                type: Components.FormControlTypes.Dropdown,
                required: true,
                items,
                onChange: (item) => {
                    // See which one is selected
                    if (item.value == "Keyword") {
                        // Show/Hide the controls
                        ctrlRegex.hide();
                        ctrlRegexPatterns.hide();
                        ctrlPermissions.hide();
                        ctrlSearchTerms.show();
                    } else {
                        // Show/Hide the controls
                        ctrlRegex.show();
                        regexItems.length > 0 ? ctrlRegexPatterns.show() : null;
                        ctrlPermissions.show();
                        ctrlSearchTerms.hide();
                    }
                }
            } as Components.IFormControlPropsDropdown,
            {
                label: "Search Terms",
                name: "SearchTerms",
                className: "mb-3 d-none",
                description: "Enter the search terms using quotes for phrases [Ex: movie \"social media\" show]",
                type: Components.FormControlTypes.TextField,
                required: true,
                value: keywords,
                errorMessage: "You must enter at least 1 search term.",
                onControlRendered: ctrl => { ctrlSearchTerms = ctrl; }
            },
            {
                label: "Regex Categories",
                name: "RegexCategories",
                className: `mb-3 ${regexItems.length == 0 ? "d-none" : ""}`,
                description: "Populates the regex patterns based on the selected category.",
                type: Components.FormControlTypes.Dropdown,
                items: regexItems,
                onControlRendered: ctrl => { ctrlRegexPatterns = ctrl; },
                onChange: (item) => {
                    // See if an item was selected
                    if (item?.data) {
                        // Set the patterns
                        ctrlRegex.setValue((item.data as IRegexPattern).patterns.join("\r\n"));

                        // Clear the selection
                        ctrlRegexPatterns.dropdown.setValue(null);
                    }
                }
            } as Components.IFormControlPropsDropdown,
            {
                label: "Regex Patterns",
                name: "RegexPatterns",
                className: "mb-3",
                description: "Enter the regular expression pattern on each line.",
                type: Components.FormControlTypes.TextArea,
                required: true,
                rows: 10,
                errorMessage: "You must enter at least 1 search term.",
                onControlRendered: ctrl => { ctrlRegex = ctrl; },
                onValidate: (ctrl, results) => {
                    // Ensure a pattern exists
                    if ((results.value || "").trim().length == 0) {
                        // Invalidate the extensions
                        results.isValid = false;
                        results.invalidMessage = "A regex pattern is required.";
                    } else {
                        let errors = [];

                        // Parse the patterns
                        results.value.split(/\r?\n/).forEach(pattern => {
                            // Validate the regex pattern
                            try {
                                // Create a new regex
                                new RegExp(results.value);
                            } catch (ex) {
                                // Add the error
                                errors.push(pattern);
                            }
                        });

                        // See if errors exist
                        if (errors.length > 0) {
                            // Invalidate the regex pattern
                            results.isValid = false;
                            results.invalidMessage = "The regex pattern is not valid.";
                        }
                    }

                    // Return the results
                    return results;
                }
            } as Components.IFormControlPropsTextField,
            {
                label: "Load Permissions",
                name: "LoadPermissions",
                type: Components.FormControlTypes.Switch,
                value: true,
                onControlRendered: ctrl => { ctrlPermissions = ctrl; }
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

    // Refreshes the item
    private static refreshItem(searchItem: ISearchItem): PromiseLike<ISearchItem> {
        // Return a promise
        return new Promise(resolve => {
            let web = this._loadOneDrive ? Web.getOneDrive() : Web(searchItem.SPWebUrl, { requestDigest: DataSource.SiteContext.FormDigestValue });

            // Get the list item
            web.Lists(searchItem["ListTitle"]).Items(searchItem["ItemId"]).query({
                Expand: ["RoleAssignments/Member/Users", "RoleAssignments/RoleDefinitionBindings"],
                Select: ["Id", "HasUniqueRoleAssignments"]
            }).execute((item) => {
                // Update the item
                searchItem.HasUniquePermissions = item.HasUniqueRoleAssignments;
                searchItem.Overshared = ViewPermissions.isOvershared(item.RoleAssignments.results as any) ? "Yes" : "No";
                searchItem.Permissions = item.RoleAssignments.results as any;

                // Find the item
                for (let i = 0; i < this._items.length; i++) {
                    // See if this is the item
                    let item = this._items[i];
                    if (item["listItem"]["ItemId"] == searchItem["ItemId"] && item.ListId == searchItem.ListId) {
                        // Update the item
                        this._items[i] = searchItem;
                        break;
                    }
                }

                // Resolve the request
                resolve(searchItem);
            });
        });
    }

    // Renders the errors
    private static renderErrors(searchType: "Search" | "Regex" | "Library") {
        let isLibrary = searchType === "Library";
        if (isLibrary) {
            // Clear the canvas
            CanvasForm.clear();
            CanvasForm.setSize(Components.OffcanvasSize.Medium2);
        } else {
            // Clear the modal
            Modal.clear();
            Modal.setType(Components.ModalTypes.Large);
        }

        // Set the header
        (isLibrary ? CanvasForm : Modal).setHeader("Document Errors");

        // Render a dashboard
        new Dashboard({
            el: isLibrary ? CanvasForm.BodyElement : Modal.BodyElement,
            navigation: {
                title: "M365 Group Errors",
                showFilter: false,
                showSearch: false,
                itemsEnd: [{
                    text: "Export to CSV",
                    className: "btn-outline-light me-2",
                    isButton: true,
                    onClick: () => {
                        // Export the CSV
                        new ExportCSV("searchDocs.csv", CSVFields, this._itemErrors);
                    }
                }]
            },
            table: {
                rows: this._itemErrors,
                columns: [
                    {
                        name: "Title",
                        title: "Filename",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            el.innerHTML = `
                                <small class="text-muted">File Name: </small>${item.Title}
                                <br/>
                                <small class="text-muted">Path: </small>${item.FileUrl || item.Path}
                            `;
                        }
                    },
                    {
                        name: "SensitivityLabel",
                        title: "Sensitivity Label"
                    },
                    {
                        name: "ErrorMessage",
                        title: "Error Message"
                    }
                ]
            }
        });

        // Render the footer
        Components.Button({
            el: isLibrary ? CanvasForm.BodyElement : Modal.FooterElement,
            text: "Close",
            type: Components.ButtonTypes.OutlinePrimary,
            onClick: () => {
                isLibrary ? CanvasForm.hide() : Modal.hide();
            }
        });

        // Show the errors
        isLibrary ? CanvasForm.show() : Modal.show();
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, auditOnly: boolean, hidePermissions: boolean, searchType: "Search" | "Regex" | "Library", onClose: () => void) {
        let isSearch = searchType === "Search" ? true : false;

        // Render the summary
        this._dashboard = new Dashboard({
            el,
            navigation: {
                title: "Search Documents",
                showFilter: false,
                items: [{
                    text: "New Search",
                    className: "btn-outline-light" + (searchType === "Library" ? " d-none" : ""),
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
                }, {
                    text: "Errors",
                    className: "btn-outline-light ms-2",
                    isButton: true,
                    onClick: () => {
                        // Display the errors
                        this.renderErrors(searchType);
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
                            "targets": 5,
                            "orderable": false,
                            "searchable": false
                        }
                    ];

                    // See if we are hiding the permissions
                    if (hidePermissions) {
                        // Hide the Permissions column
                        dtProps.columnDefs.push({
                            "targets": [3, 4],
                            "visible": false
                        });
                    }

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
                                span.innerHTML = item.RegexPatterns;
                            }

                            // Append the span
                            el.appendChild(span);
                        }
                    },
                    {
                        name: "",
                        title: "Overshared",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            let isOvershared = item.Overshared === "Yes" ? true : false;

                            // Set the order info
                            el.setAttribute("data-order", item.Overshared);

                            // Make the badge display in the middle
                            el.style.verticalAlign = "middle";

                            // Render a badge
                            let badge = Components.Badge({
                                el,
                                className: "me-2",
                                content: isOvershared ? "Overshared" : item.Overshared,
                                type: isOvershared ? Components.BadgeTypes.Danger : Components.BadgeTypes.Secondary,
                                isPill: true
                            });

                            // See if this is overshared
                            if (isOvershared) {
                                // Render a tooltip
                                Components.Tooltip({
                                    target: badge.el,
                                    content: `The file has been flagged as overshared because it's shared with the following groups:<br/>${ViewPermissions.getOversharedGroups(item.Permissions).join("<br/>")}`
                                });
                            }
                        }
                    },
                    {
                        name: "",
                        title: "Permissions",
                        onRenderCell: (el, col, item: ISearchItem) => {
                            let adGroups = 0;
                            let m365Groups = 0;
                            let siteGroups = 0;
                            let users = 0;

                            // Ensure permissions exist
                            if (item.Permissions == null) { return; }

                            // Parse the permissions
                            item.Permissions.forEach(role => {
                                // See if this is a user
                                switch (role.Member.PrincipalType) {
                                    case SPTypes.PrincipalTypes.User:
                                        users++;
                                        break;
                                    case SPTypes.PrincipalTypes.SharePointGroup:
                                        siteGroups++;
                                        break;
                                    default:
                                        let groupId = M365Groups.getGroupId(role.Member.LoginName);
                                        groupId ? m365Groups++ : adGroups++;
                                        break;
                                }
                            });

                            // Output the permission information
                            el.innerHTML = `
                                <b>Unique Permissions: </b>${item.HasUniquePermissions ? "Yes" : "No"}
                                <br/>
                                <b># of Users: </b>${users}
                                <br/>
                                <b># of Site Groups: </b>${siteGroups}
                                <br/>
                                <b># of AD Groups: </b>${adGroups}
                                <br/>
                                <b># of M365 Groups: </b>${m365Groups}
                                <br/>
                            `;
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

                                // Ensure permissions exist
                                if (item.Permissions) {
                                    // Add the view permissions button
                                    tooltips.add({
                                        content: "Click to view the permissions for this document.",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            text: "View Permissions",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // View the permissions for the document
                                                ViewPermissions.show(item);
                                            }
                                        }
                                    });
                                }

                                // See if the file is overshared
                                if (item.Overshared === "Yes") {
                                    tooltips.add({
                                        content: "Click to remove the groups that are flagging this file as overshared.",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            text: "Secure File",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Remove the overshared groups from the permissions
                                                ViewPermissions.removeOversharedGroups(item, () => {
                                                    // Refresh the item
                                                    this.refreshItem(item).then(updatedItem => {
                                                        // Update this row
                                                        this._dashboard.updateRow(rowIdx, updatedItem);
                                                    });
                                                });
                                            }
                                        }
                                    });
                                }

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
        this._itemErrors = [];
        this._loadOneDrive = values["LoadOneDrive"] == "true";

        // Show a loading dialog
        LoadingDialog.setHeader("Searching Site");
        LoadingDialog.setBody("Searching the content on this site...");
        LoadingDialog.show();

        // Get the form values
        let fileExt = values["FileTypes"] ? values["FileTypes"].split(' ') : null;
        let searchTerms = (values["SearchTerms"] || "").split(' ');
        let searchType = (values["SearchType"] || "");
        let targetFolder = values["TargetFolder"];
        let loadPermissions = values["LoadPermissions"] as any == true ? true : false;

        // Set the regex patterns
        let regexPatterns = [];
        (values["RegexPatterns"]?.split(/\r?\n/) || []).forEach(pattern => {
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
                this.renderSummary(el, auditOnly, true, "Search", onClose);

                // Hide the sub-nav
                this._elSubNav.classList.add("d-none");

                // Hide the loading dialog
                LoadingDialog.hide();
            });
        } else {
            // Clear the element
            while (el.firstChild) { el.removeChild(el.firstChild); }

            // Render the summary
            this.renderSummary(el, auditOnly, !loadPermissions, values["TargetList"] ? "Library" : "Regex", onClose);

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

                    // Get the libraries for this site
                    let web = this._loadOneDrive ? Web.getOneDrive() : Web(siteItem.text, { requestDigest: DataSource.SiteContext.FormDigestValue });
                    web.Lists().query({
                        Filter: filter,
                        Expand: ["RootFolder"],
                        GetAllItems: true,
                        Select: ["Id", "Title", "RootFolder/ServerRelativeUrl"],
                        Top: 5000
                    }).execute(libs => {
                        // Update the dialog
                        this._elSubNav.children[1].innerHTML = "Loading the files for the libraries...";

                        // Get the drives for this web
                        v2.drives({
                            siteId: this._loadOneDrive ? DataSource.OneDriveSite.Id : DataSource.Site.Id,
                            webId: this._loadOneDrive ? DataSource.OneDriveWeb.Id : DataSource.Web.Id
                        }).execute(drives => {
                            // Update the dialog
                            this._elSubNav.children[1].innerHTML = "Loading the files for the libraries...";

                            // Process the libraries
                            Helper.Executor(libs.results, lib => {
                                // Return a promise
                                return new Promise(resolveLib => {
                                    // Analyze the library
                                    this.analyzeLibrary(siteItem.value, siteItem.text, lib, drives.results, targetFolder, loadPermissions, fileExt, regexPatterns).then(resolveLib);
                                });
                            }).then(resolve);
                        });
                    });
                });
            }).then(() => {
                // Hide the sub-nav
                this._elSubNav.classList.add("d-none");

                // Get the error button
                let elNav = el.querySelector("#navigation .navbar-nav");

                // See if no errors exist
                if (this._itemErrors.length == 0) {
                    // Remove the last button
                    elNav.querySelector("li:last-child").remove();
                } else {
                    // Update the error text
                    elNav.querySelector("li:last-child > a").innerHTML = `${this._itemErrors.length} Errors`;
                }
            });

            // Hide the loading dialog
            LoadingDialog.hide();
        }
    }

    // Searches a library for agents
    static searchLibrary(auditOnly: boolean, values: { [key: string]: string }) {
        // Clear a modal form
        Modal.clear();
        Modal.setHeader("Search Documents");
        Modal.setType(Components.ModalTypes.Full);

        // Run the report
        this.run(Modal.BodyElement, auditOnly, values, () => { });

        // Render the footer
        Components.ButtonGroup({
            el: Modal.FooterElement,
            className: "mt-3",
            buttons: [
                {
                    text: "Close",
                    type: Components.ButtonTypes.OutlineSecondary,
                    onClick: () => {
                        // Hide the form
                        Modal.hide();
                    }
                }
            ]
        });

        // Show the form
        Modal.show();
    }

    // Stops the report
    static stop() { this._stopFl = true; }
}