import { Dashboard, Documents, LoadingDialog, Modal } from "dattatable";
import { Components, Helper, SPTypes, Types, Web, v2 } from "gd-sprest-bs";
import { fileEarmarkText } from "gd-sprest-bs/build/icons/svgs/fileEarmarkText";
import { DataSource } from "../ds";
import Strings from "../strings";
import { ExportCSV } from "./exportCSV";

interface ISensitivityLabelItem {
    Author: string;
    File: Types.Microsoft.Graph.driveItem;
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

interface ISetSensitivityLabelResponse {
    errorFl: boolean;
    error?: any;
    fileName: string;
    message: string;
    url: string;
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

const CSVSensitivityLabelResponseFields = [
    "errorFl", "url", "fileName", "message"
]

export class SensitivityLabels {
    private static _dashboard: Dashboard = null;
    private static _elSubNav: HTMLElement = null;
    private static _filterLabels: string[] = [];
    private static _items: ISensitivityLabelItem[] = [];
    private static _loadOneDrive: boolean = false;
    private static _stopFl: boolean = false;

    // Gets the form fields to display
    static getFormFields(): Components.IFormControlProps[] {
        return [
            {
                name: "SearchType",
                className: "my-3",
                type: Components.FormControlTypes.MultiSwitch,
                required: true,
                errorMessage: "A selection is required.",
                items: [
                    {
                        name: "WithLabels",
                        label: "Find all files with a label",
                        isSelected: true
                    },
                    {
                        name: "WithoutLabels",
                        label: "Find all files without a label"
                    }
                ]
            } as Components.IFormControlPropsMultiSwitch,
            {
                name: "FilterLabel",
                className: "mb-3",
                type: Components.FormControlTypes.MultiDropdownCheckbox,
                items: DataSource.SensitivityLabelItems.slice(1),
                label: "Find Files with Label:",
                placeholder: "Select Label(s)",
                description: "Filter results for specific sensitivity label(s)."
            } as Components.IFormControlPropsMultiDropdownCheckbox
        ];
    }

    // Analyzes the libraries
    private static analyzeLibraries(webId: string, webUrl: string, libraries: Types.SP.ListOData[], drives: Types.Microsoft.Graph.drive[], withLabelsFl, withoutLabelsFl) {
        // Return a promise
        return new Promise(resolve => {
            let counter = 0;
            let siteText = this._elSubNav.children[0].innerHTML;

            // Parse the libraries
            Helper.Executor(libraries, lib => {
                let fileItems: ISensitivityLabelItem[] = [];

                // See if we are stopping this process
                if (this._stopFl) { return; }

                // Update the dialog
                this._elSubNav.children[0].innerHTML = `${siteText} [Analyzing Library ${++counter} of ${libraries.length}]: ${lib.Title}`;

                // Get the drive for this library
                let drive = drives.find(drive => {
                    return drive.name == lib.Title || drive.webUrl.endsWith(lib.RootFolder.ServerRelativeUrl);
                });

                // Ensure a drive exists, otherwise check the next library
                if (drive == null) { return; }

                // Return a promise
                return new Promise(resolve => {
                    // Update the dialog
                    this._elSubNav.children[1].innerHTML = `Analyzing the files for this library...`;

                    // Get the files for this library
                    let filesProcessed = 0;
                    DataSource.loadFiles(webId, drive.id, null, (file: Types.Microsoft.Graph.driveItem) => {
                        let hasLabel = file.sensitivityLabel && file.sensitivityLabel.displayName ? true : false;

                        // Update the dialog
                        this._elSubNav.children[1].innerHTML = `Analyzing the files for this library. Files Analyzed: ${++filesProcessed}`;

                        // Add the file, based on the flags
                        if ((withLabelsFl && hasLabel) || (withoutLabelsFl && !hasLabel)) {
                            // See if we are filter for a label
                            if (this._filterLabels.length > 0) {
                                // See if this is a target label
                                if (this._filterLabels.indexOf(file.sensitivityLabel.id) < 0) { return; }
                            }

                            let fileInfo = file.name.split('.');
                            let folderPath = file.parentReference.path.split('root:')[1];

                            // Append the data
                            let fileItem = {
                                Author: file.createdBy.user["email"] || file.createdBy.user["displayName"],
                                File: file,
                                FileExtension: fileInfo[fileInfo.length - 1],
                                FileName: file.name,
                                ListId: lib.Id,
                                ListTitle: lib.Title,
                                Path: `${lib.RootFolder.ServerRelativeUrl}${folderPath}/${file.name}`,
                                SensitivityLabel: file.sensitivityLabel.displayName,
                                SensitivityLabelId: file.sensitivityLabel.id,
                                WebId: webId,
                                WebUrl: webUrl
                            };

                            // Save a reference to the item
                            this._items.push(fileItem);
                            fileItems.push(fileItem);

                            // See if we have hit 100 items
                            if (fileItems.length >= 100) {
                                // Add the items to the datatable
                                this._dashboard.Datatable.addRow(fileItems);
                                fileItems = [];
                            }
                        }

                        // Return the stop flag
                        return this._stopFl;
                    }).then(() => {
                        // See if items exist
                        if (fileItems.length > 0) {
                            // Add the items to the datatable
                            this._dashboard.Datatable.addRow(fileItems);
                        }

                        // Check the next library
                        resolve(null);
                    });
                });
            }).then(resolve);
        });
    }

    // Labels file
    private static labelFile(file: Types.Microsoft.Graph.driveItem, label: string, labelId: string, justification: string, responses: ISetSensitivityLabelResponse[]): PromiseLike<ISetSensitivityLabelResponse[]> {
        // Update the sensitivity label for this file
        return new Promise(resolve => {
            // Update the sensitivity label
            // The action source allowed values are:
            // * Manual - The label was applied manually by a user.
            // * Automatic - The label was applied automatically based on pre-defined rules or conditions.
            // * Default - The label was applied as the default label for a document library or location.
            // * Policy - The label was applied as part of a policy configuration.
            // The assignment method allowed values are:
            // * Standard - The label was applied manually by a user or through a standard process.
            // * Privileged - The label was applied by an administrator or through a privileged operation.
            // * Auto - The label was applied automatically by the system based on predefined rules or policies, such as auto-classification of sensitive content.
            // * UnknownFutureValue - A placeholder for future use, indicating an unknown or unsupported assignment method.
            v2.drive({
                driveId: file.parentReference.driveId,
                siteId: file.parentReference.siteId,
                targetInfo: { keepalive: true }
            }).items(file.id).setSensitivityLabel("Manual", "Privileged", labelId, justification).execute(
                // Success
                () => {
                    // Add the response
                    let response: ISetSensitivityLabelResponse = {
                        errorFl: false,
                        fileName: file.name,
                        message: `The file was successfully labelled: ${label}.`,
                        url: file.webUrl
                    };
                    responses.push(response);
                    this._dashboard.Datatable.addRow(response);

                    // Resolve the request
                    resolve(responses);
                },

                // Error
                err => {
                    // Get the error
                    let error = null;
                    try { error = JSON.parse(err["response"])?.error?.message; }
                    catch { error = err; }

                    // Add the response
                    let response: ISetSensitivityLabelResponse = {
                        errorFl: true,
                        error: err,
                        fileName: file.name,
                        message: `There was an error tagging this file: ${error}`,
                        url: file.webUrl
                    };
                    responses.push(response);
                    this._dashboard.Datatable.addRow(response);

                    // Resolve the request
                    resolve(responses);
                }
            );
        });
    }

    // Labels files for a specified folder
    private static labelFilesInFolder(webId: string, listName: string, driveId: string, folderId: string, fileExtensions: string[], label: Components.IDropdownItem, replaceLabel: Components.IDropdownItem, overrideLabelFl: boolean, justification: string) {
        let responses: ISetSensitivityLabelResponse[] = [];

        // Show the responses
        this.showResponses(responses, () => {
            // Stop the worker process
            worker.stop();
        });

        // Update the dialog
        this._elSubNav.children[0].innerHTML = `Loading files from library: ${listName}`;

        // Process the labels as we load the files
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

            // Get the file to process
            let file = filesToProcess.splice(0, 1)[0];

            // See if the file extensions are provided
            if (fileExtensions && fileExtensions.length > 0) {
                let analyzeFile = false

                // Loop through the file extensions
                fileExtensions.forEach(fileExt => {
                    // Set the flag if there is match
                    if (file.name?.toLowerCase().endsWith(fileExt.toLowerCase())) { analyzeFile = true; }
                });

                // See if we are not analyzing the file
                if (!analyzeFile) {
                    // Add a response
                    let response: ISetSensitivityLabelResponse = {
                        errorFl: false,
                        fileName: file.name,
                        message: `Skipping file, file extension not listed to process.`,
                        url: file.webUrl
                    };
                    responses.push(response);
                    this._dashboard.Datatable.addRow(response);

                    // Update the dialog
                    this._elSubNav.children[1].innerHTML = `[Processed ${++processedCounter} of ${fileCounter}] ${response.message}`;

                    // Check the next file
                    return;
                }
            }

            // See if we are replacing a label
            if (replaceLabel) {
                // Set the flag
                overrideLabelFl = true;

                // See if file is not labelled with the target
                if (file.sensitivityLabel?.id != label.value) {
                    // Add a response
                    let response: ISetSensitivityLabelResponse = {
                        errorFl: false,
                        fileName: file.name,
                        message: `Skipping file, replacing '${label.text}' label, file label: '${file.sensitivityLabel.displayName}.`,
                        url: file.webUrl
                    };
                    responses.push(response);
                    this._dashboard.Datatable.addRow(response);

                    // Update the dialog
                    this._elSubNav.children[1].innerHTML = `[Processed ${++processedCounter} of ${fileCounter}] ${response.message}`;

                    // Check the next file
                    return;
                }
            }

            // Set the label for this file
            let fileLabel = replaceLabel || label;

            // See if this file is already has this label
            // or
            // See if the file has a label and we aren't overriding them 
            if (file.sensitivityLabel?.id == fileLabel.value || (file.sensitivityLabel?.id && !overrideLabelFl)) {
                // Add a response
                let response: ISetSensitivityLabelResponse = {
                    errorFl: false,
                    fileName: file.name,
                    message: `Skipping file, it's already labelled: '${file.sensitivityLabel.displayName}'.`,
                    url: file.webUrl
                };
                responses.push(response);
                this._dashboard.Datatable.addRow(response);

                // Check the next file
                return;
            }

            // Increment the # of files being processed
            processingCounter++;

            // Update the dialog
            this._elSubNav.children[1].innerHTML = `[Processing ${processedCounter} of ${fileCounter}] Labelling File: ${file.name}`;

            // Labels the file
            let labelFile = () => {
                // Label the file
                this.labelFile(file, fileLabel.text, fileLabel.value, justification, responses).then(() => {
                    // Update the dialog
                    this._elSubNav.children[1].innerHTML = `[Processed ${++processedCounter} of ${fileCounter}] File Labelled: ${file.name}`;

                    // Decrement the # of files being processed
                    processingCounter--;
                });
            }

            // See if we are sending multiple requests
            if (Strings.MaxRequests > 1) {
                // Set a timeout to avoid throttling
                setTimeout(labelFile, 250 * processingCounter);
            } else {
                // Label the file
                labelFile();
            }
        }, 100);

        // Set the completed event
        let onCompleted = () => {
            // Clear the sub-nav
            this._elSubNav.classList.add("d-none");
        };

        // Load the files for this drive
        DataSource.loadFiles(webId, driveId, folderId, file => {
            // Add the file to process
            filesToProcess.push(file);

            // Update the dialog
            this._elSubNav.children[0].innerHTML = `Loading files from library: ${listName} [Files Loaded: ${++fileCounter}]`;

            // Ensure the process is running
            worker.start();

            // Return the stop flag
            return this._stopFl;
        }).then(() => {
            // Update the dialog
            this._elSubNav.children[0].innerHTML = `Library: ${listName} [Files Loaded: ${++fileCounter}]`;

            // Ensure the process is running
            worker.start();
        });
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, auditOnly: boolean, onClose: () => void) {
        // Render the summary
        this._dashboard = new Dashboard({
            el,
            navigation: {
                title: "Sensitivity Labels",
                showFilter: false,
                items: [{
                    text: "New Search",
                    className: "btn-outline-light",
                    isButton: true,
                    onClick: () => {
                        // Set the stop flag
                        this._stopFl = true;

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
                        name: "Path",
                        title: "File Info",
                        onRenderCell: (el, col, item: ISensitivityLabelItem) => {
                            // Show the file info
                            el.innerHTML = `
                                <b>Created By: </b>${item.Author}
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
                        onRenderCell: (el, col, row: ISensitivityLabelItem, rowIdx) => {
                            // Render the buttons
                            let tooltips = Components.TooltipGroup({
                                el,
                                tooltips: [
                                    {
                                        content: "View Document",
                                        btnProps: {
                                            className: "pe-2 py-1",
                                            iconClassName: "mx-1",
                                            iconType: fileEarmarkText,
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

                            // Ensure we can make updates
                            if (!auditOnly) {
                                // Add the label button
                                tooltips.add({
                                    content: "Sets the label for the file.",
                                    btnProps: {
                                        className: "pe-2 py-1",
                                        text: "Set Label",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the form to label the file
                                            this.showLabelFileForm(row.File, label => {
                                                // Update the row cell
                                                this._dashboard.updateCell(rowIdx, 3, label);
                                            });
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
        let data: IWebItem[] = [];
        this._loadOneDrive = values["LoadOneDrive"] == "true";
        this._stopFl = false;

        // Clear the items
        this._items = [];

        // Set the flags
        let withLabelsFl = false;
        let withoutLabelsFl = false;
        (values["SearchType"] as any as Components.ICheckboxGroupItem[]).forEach(item => {
            withLabelsFl = item.name == "WithLabels" ? true : withLabelsFl;
            withoutLabelsFl = item.name == "WithoutLabels" ? true : withoutLabelsFl;
        });

        // Set the labels filter
        this._filterLabels = [];
        let labels = values["FilterLabel"] as any as Components.IDropdownItem[];
        (labels || []).forEach(label => {
            this._filterLabels.push(label.value);
        });

        // Show a loading dialog
        LoadingDialog.setHeader("Searching Site");
        LoadingDialog.setBody("Loading the libraries...");
        LoadingDialog.show();

        // Clear the element
        while (el.firstChild) { el.removeChild(el.firstChild); }

        // Render the summary
        this.renderSummary(el, auditOnly, onClose);

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
                // Get the libraries for this site
                let web = this._loadOneDrive ? Web.getOneDrive() : Web(siteItem.text, { requestDigest: DataSource.SiteContext.FormDigestValue });
                web.Lists().query({
                    Filter: `Hidden eq false and BaseTemplate eq ${SPTypes.ListTemplateType.DocumentLibrary} or BaseTemplate eq ${SPTypes.ListTemplateType.MySiteDocumentLibrary} or BaseTemplate eq ${SPTypes.ListTemplateType.PageLibrary}`,
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
                    this._elSubNav.children[1].innerHTML = "Loading the files for the libraries...";

                    // Get the drives for this web
                    v2.drives({
                        siteId: this._loadOneDrive ? DataSource.OneDriveSite.Id : DataSource.Site.Id,
                        webId: this._loadOneDrive ? DataSource.OneDriveWeb.Id : DataSource.Web.Id
                    }).execute(drives => {
                        // Analyze the libraries
                        this.analyzeLibraries(siteItem.value, siteItem.text, libs.results, drives.results, withLabelsFl, withoutLabelsFl).then(resolve);
                    });
                });
            });
        }).then(() => {
            // Hide the sub-nav
            this._elSubNav.classList.add("d-none");
        });
    }

    // Shows the form to label a file
    private static showLabelFileForm(file: Types.Microsoft.Graph.driveItem, onUpdate: (label: string) => void) {
        // Set the modal header
        Modal.clear();
        Modal.setHeader("Set Sensitivity Label");

        // Set the form
        let form = Components.Form({
            el: Modal.BodyElement,
            groupClassName: "mb-3",
            controls: [
                {
                    name: "SensitivityLabel",
                    label: "Select Sensitivity Label:",
                    description: "This will set any file that isn't currently labelled.",
                    errorMessage: "A sensitivity label is required.",
                    items: DataSource.SensitivityLabelItems,
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    onValidate: (ctrl, results) => {
                        // Ensure a selection exists
                        results.isValid = results.value && results.value.text ? true : false;
                        return results;
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "Justification",
                    label: "Justification:",
                    description: "Your organization requires justification to change this label.",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    items: [
                        { text: "Previous label no longer applies" },
                        { text: "Previous label was incorrect" },
                        { text: "Other" }
                    ],
                    onChange: (item) => {
                        let ctrlTextbox = form.getControl("JustificationOther");

                        // See if we are showing it
                        if (item.text == "Other") {
                            // Show it
                            ctrlTextbox.show();
                        } else {
                            // Hide it
                            ctrlTextbox.hide();
                        }
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "JustificationOther",
                    label: "Explain Justification:",
                    description: "Do not enter sensitive information",
                    className: "d-none",
                    type: Components.FormControlTypes.TextField,
                    errorMessage: "A justification is required.",
                    onValidate: (ctrl, results) => {
                        let item = form.getValues()["Justification"] as Components.IDropdownItem;

                        // See if we are expecting a justification
                        if (item.text == "Other") {
                            // Set the falg
                            results.isValid = results.value ? true : false;
                        }

                        // Return the results
                        return results;
                    }
                } as Components.IFormControlPropsTextField,
            ]
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Sets the default sensitivity label to the selected option.",
                    btnProps: {
                        text: "Update",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let values = form.getValues();
                                let label: Components.IDropdownItem = values["SensitivityLabel"];

                                // Update the justification
                                let justification = values["Justification"].text;
                                justification = justification == "Other" ? values["JustificationOther"] : justification;

                                // Show a loading dialog
                                LoadingDialog.setHeader("Setting Label");
                                LoadingDialog.setBody("Updating the label for this file.");
                                LoadingDialog.show();

                                // Label the file
                                this.labelFile(file, label.text, label.value, justification, []).then(responses => {
                                    // See if it was successful
                                    if (!responses[0].errorFl) {
                                        // Call the event
                                        onUpdate(label.text);
                                    } else {
                                        // Set the error
                                        let ctrl = form.getControl("SensitivityLabel");
                                        ctrl.updateValidation(ctrl.el, {
                                            isValid: false,
                                            invalidMessage: responses[0].message
                                        });
                                    }

                                    // Hide the dialogs
                                    LoadingDialog.hide();
                                    Modal.hide();
                                });
                            }
                        }
                    }
                },
                {
                    content: "Closes the dialog.",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Close the modal
                            Modal.hide();
                        }
                    }
                }
            ]
        });

        // Show the modal
        Modal.show();
    }

    // Shows the form for a library
    static showLibraryForm(webId: string, listName: string, driveId: string, defaultLabel: string, folders: Components.IDropdownItem[], disableSensitivityLabelOverride: boolean, fileTypes: string) {
        let defaultItem: Components.IDropdownItem = { text: "All Folders", value: "" };

        // Set the modal header
        Modal.clear();
        Modal.setHeader("Set Sensitivity Label");

        // Set the form
        let form = Components.Form({
            el: Modal.BodyElement,
            groupClassName: "mb-3",
            onRendered: () => {
                // Set the folders to trigger the change events for the sub-folders
                form.getControl("ListFolder").dropdown.setItems([defaultItem].concat(folders));
            },
            controls: [
                {
                    name: "SensitivityLabel",
                    label: "Select Sensitivity Label:",
                    description: "This will set any file that isn't currently labelled.",
                    errorMessage: "A sensitivity label is required.",
                    items: DataSource.SiteSensitivityLabelItems,
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    value: defaultLabel,
                    onValidate: (ctrl, results) => {
                        // Ensure a selection exists
                        results.isValid = results.value && results.value.text ? true : false;
                        return results;
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "OverrideLabel",
                    label: "Override Label?",
                    description: "If a label already exists, it will attempt to override it.",
                    isDisabled: disableSensitivityLabelOverride,
                    type: Components.FormControlTypes.Switch,
                    value: false
                },
                {
                    name: "ReplaceLabel",
                    label: "Replace Sensitivity Label:",
                    description: "This will replace the selected label with this label.",
                    items: DataSource.SensitivityLabelItems,
                    type: Components.FormControlTypes.Dropdown,
                    value: defaultLabel,
                    onChange: (item) => {
                        // See if an item was selected
                        if (item && item.value) {
                            // Hide the override label option
                            form.getControl("OverrideLabel").hide();
                        } else {
                            // Show the override label option
                            form.getControl("OverrideLabel").show();
                        }
                    },
                    onValidate: (ctrl, results) => {
                        let item = results.value as Components.IDropdownItem;;

                        // See if a value exists
                        if (item?.value) {
                            // Get the source value
                            let sourceItem = form.getControl("SensitivityLabel").getValue() as Components.IDropdownItem;
                            if (sourceItem?.value == item.value) {
                                // Set the validation
                                results.isValid = false;
                                results.invalidMessage = "The replace label cannot be the same as the sensitivity label.";
                            }
                        }

                        // Return the results
                        return results;
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "FileTypes",
                    label: "File Types",
                    type: Components.FormControlTypes.TextField,
                    value: fileTypes
                },
                {
                    name: "Justification",
                    label: "Justification:",
                    description: "Your organization requires justification to change this label.",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    items: [
                        { text: "Previous label no longer applies" },
                        { text: "Previous label was incorrect" },
                        { text: "Other" }
                    ],
                    onChange: (item) => {
                        let ctrlTextbox = form.getControl("JustificationOther");

                        // See if we are showing it
                        if (item.text == "Other") {
                            // Show it
                            ctrlTextbox.show();
                        } else {
                            // Hide it
                            ctrlTextbox.hide();
                        }
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "JustificationOther",
                    label: "Explain Justification:",
                    description: "Do not enter sensitive information",
                    className: "d-none",
                    type: Components.FormControlTypes.TextField,
                    errorMessage: "A justification is required.",
                    onValidate: (ctrl, results) => {
                        let item = form.getValues()["Justification"] as Components.IDropdownItem;

                        // See if we are expecting a justification
                        if (item.text == "Other") {
                            // Set the falg
                            results.isValid = results.value ? true : false;
                        }

                        // Return the results
                        return results;
                    }
                } as Components.IFormControlPropsTextField,
                {
                    name: "ListFolder",
                    label: "Select a Folder:",
                    description: "Targets a specific folder to tag, otherwise will apply to all files in the library.",
                    type: Components.FormControlTypes.Dropdown,
                    onChange: (item) => {
                        // Clear the sub-folder
                        let subFolder = form.getControl("ListSubFolder1");
                        subFolder.dropdown.setItems([defaultItem]);

                        // See if a folder is selected
                        if (item && item.value) {
                            // Show the folder
                            subFolder.show();

                            // Set the dropdown items
                            subFolder.dropdown.setItems([defaultItem]);

                            // Load the folders
                            DataSource.loadFolders(webId, driveId, item.value).then(items => {
                                // Set the dropdown items
                                subFolder.dropdown.setItems([defaultItem].concat(items));
                            });
                        } else {
                            subFolder.hide();
                        }
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "ListSubFolder1",
                    label: "Select a Sub-Folder:",
                    description: "Targets a specific folder to tag, otherwise will apply to all files in this folder.",
                    type: Components.FormControlTypes.Dropdown,
                    onChange: (item) => {
                        // Clear the sub-folder
                        let subFolder = form.getControl("ListSubFolder2");
                        subFolder.dropdown.setItems([]);

                        // See if a folder is selected
                        if (item && item.value) {
                            // Show the folder
                            subFolder.show();

                            // Set the dropdown items
                            subFolder.dropdown.setItems([defaultItem]);

                            // Load the folders
                            DataSource.loadFolders(webId, driveId, item.value).then(items => {
                                // Set the dropdown items
                                subFolder.dropdown.setItems([defaultItem].concat(items));
                            });
                        } else {
                            // Hide the folder
                            subFolder.hide();
                        }
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "ListSubFolder2",
                    label: "Select a Sub-Folder:",
                    description: "Targets a specific folder to tag, otherwise will apply to all files in this folder.",
                    type: Components.FormControlTypes.Dropdown,
                    onChange: (item) => {
                        // Clear the sub-folder
                        let subFolder = form.getControl("ListSubFolder3");
                        subFolder.dropdown.setItems([]);

                        // See if a folder is selected
                        if (item && item.value) {
                            // Show the folder
                            subFolder.show();

                            // Set the dropdown items
                            subFolder.dropdown.setItems([defaultItem]);

                            // Load the folders
                            DataSource.loadFolders(webId, driveId, item.value).then(items => {
                                // Set the dropdown items
                                subFolder.dropdown.setItems([defaultItem].concat(items));
                            });
                        } else {
                            // Hide the folder
                            subFolder.hide();
                        }
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "ListSubFolder3",
                    label: "Select a Sub-Folder:",
                    description: "Targets a specific folder to tag, otherwise will apply to all files in this folder.",
                    type: Components.FormControlTypes.Dropdown,
                    onChange: (item) => {
                        // Clear the sub-folder
                        let subFolder = form.getControl("ListSubFolder4");
                        subFolder.dropdown.setItems([]);

                        // See if a folder is selected
                        if (item && item.value) {
                            // Show the folder
                            subFolder.show();

                            // Set the dropdown items
                            subFolder.dropdown.setItems([defaultItem]);

                            // Load the folders
                            DataSource.loadFolders(webId, driveId, item.value).then(items => {
                                // Set the dropdown items
                                subFolder.dropdown.setItems([defaultItem].concat(items));
                            });
                        } else {
                            // Hide the folder
                            subFolder.hide();
                        }
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "ListSubFolder4",
                    label: "Select a Sub-Folder:",
                    description: "Targets a specific folder to tag, otherwise will apply to all files in this folder.",
                    type: Components.FormControlTypes.Dropdown,
                    onChange: (item) => {
                        // Clear the sub-folder
                        let subFolder = form.getControl("ListSubFolder5");
                        subFolder.dropdown.setItems([]);

                        // See if a folder is selected
                        if (item && item.value) {
                            // Show the folder
                            subFolder.show();

                            // Set the dropdown items
                            subFolder.dropdown.setItems([defaultItem]);

                            // Load the folders
                            DataSource.loadFolders(webId, driveId, item.value).then(items => {
                                // Set the dropdown items
                                subFolder.dropdown.setItems([defaultItem].concat(items));
                            });
                        } else {
                            // Hide the folder
                            subFolder.hide();
                        }
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "ListSubFolder5",
                    label: "Select a Sub-Folder:",
                    description: "Targets a specific folder to tag, otherwise will apply to all files in this folder.",
                    type: Components.FormControlTypes.Dropdown,
                    onControlRendered: ctrl => { ctrl.hide(); },
                    onChange: (item) => {
                        let subFolder = form.getControl("ListSubFolder5");

                        // See if items exist
                        if (item) {
                            // Show the folder
                            subFolder.show();
                        } else {
                            // Hide the folder
                            subFolder.hide();
                        }
                    }
                } as Components.IFormControlPropsDropdown
            ]
        });

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Sets the default sensitivity label to the selected option.",
                    btnProps: {
                        text: "Update",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let values = form.getValues();
                                let fileExtensions: string[] = values["FileTypes"] ? values["FileTypes"].trim().split(' ') : [];
                                let label: Components.IDropdownItem = values["SensitivityLabel"];
                                let replaceLabel: Components.IDropdownItem = values["ReplaceLabel"];
                                let overrideLabelFl: boolean = values["OverrideLabel"];

                                // Set the target folder
                                let folder = values["ListFolder"].data;
                                let subFolder1 = values["ListSubFolder1"]?.data;
                                let subFolder2 = values["ListSubFolder2"]?.data;
                                let subFolder3 = values["ListSubFolder3"]?.data;
                                let subFolder4 = values["ListSubFolder4"]?.data;
                                let subFolder5 = values["ListSubFolder5"]?.data;
                                let targetFolder = subFolder5 || subFolder4 || subFolder3 || subFolder2 || subFolder1 || folder;

                                // Update the justification
                                let justification = values["Justification"].text;
                                justification = justification == "Other" ? values["JustificationOther"] : justification;

                                // Label the files
                                this.labelFilesInFolder(webId, listName, driveId, targetFolder?.id, fileExtensions, label, replaceLabel?.value ? replaceLabel : null, overrideLabelFl, justification);
                            }
                        }
                    }
                },
                {
                    content: "Closes the dialog.",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Close the modal
                            Modal.hide();
                        }
                    }
                }
            ]
        });

        // Show the modal
        Modal.show();
    }

    // Shows the responses for setting the sensitivity labels
    private static showResponses(responses: ISetSensitivityLabelResponse[], onClose?: () => void) {
        // Clear the modal and set the header
        Modal.clear();
        Modal.setType(Components.ModalTypes.Full);
        Modal.setHeader("Set Sensitivity Label Results");

        // Set the close event
        Modal.setCloseEvent(onClose);

        // Set the dashboard
        this._dashboard = new Dashboard({
            el: Modal.BodyElement,
            navigation: {
                title: "Senstivity Labels",
                showFilter: false
            },
            table: {
                rows: responses,
                columns: [
                    {
                        name: "errorFl",
                        title: "Error?"
                    },
                    {
                        name: "fileName",
                        title: "File Name",
                        onRenderCell: (el, col, item: ISetSensitivityLabelResponse) => {
                            // Clear the element
                            el.innerHTML = "";

                            // Render a link to the file
                            Components.Button({
                                el,
                                text: item.fileName,
                                href: item.url,
                                target: "_blank",
                                type: Components.ButtonTypes.OutlineLink
                            });
                        }
                    },
                    {
                        name: "message",
                        title: "Status Message"
                    }
                ]
            }
        });

        // Set the sub-nav element
        this._elSubNav = Modal.BodyElement.querySelector("#sub-navigation");
        this._elSubNav.classList.remove("d-none");
        this._elSubNav.classList.add("my-2");
        this._elSubNav.innerHTML = `<div class="h6"></div><div></div>`;

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Exports the report as a csv.",
                    btnProps: {
                        text: "Export to CSV",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Export the CSV
                            new ExportCSV("senstivity_labels.csv", CSVSensitivityLabelResponseFields, this._items);
                        }
                    }
                },
                {
                    content: "Closes the dialog.",
                    btnProps: {
                        text: "Close",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            // Close the modal
                            Modal.hide();
                        }
                    }
                }
            ]
        });

        // Show the modal
        Modal.show();
    }

    // Stops the report
    static stop() { this._stopFl = true; }
}