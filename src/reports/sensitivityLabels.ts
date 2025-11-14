import { Dashboard, DataTable, Documents, LoadingDialog, Modal } from "dattatable";
import { Components, Helper, SPTypes, Types, Web, v2 } from "gd-sprest-bs";
import { fileEarmarkText } from "gd-sprest-bs/build/icons/svgs/fileEarmarkText";
import { DataSource } from "../ds";
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
    private static _items: ISensitivityLabelItem[] = [];

    // Gets the form fields to display
    static getFormFields(): Components.IFormControlProps[] {
        return [
            {
                name: "SearchType",
                className: "mb-3",
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
            } as Components.IFormControlPropsMultiSwitch
        ];
    }

    // Analyzes the libraries
    private static analyzeLibraries(webId: string, webUrl: string, libraries: Types.SP.ListOData[], withLabelsFl, withoutLabelsFl) {
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
                            let hasLabel = file.sensitivityLabel && file.sensitivityLabel.displayName ? true : false;

                            // Add the file, based on the flags
                            if ((withLabelsFl && hasLabel) || (withoutLabelsFl && !hasLabel)) {
                                let fileInfo = file.name.split('.');
                                let folderPath = file.parentReference.path.split('root:')[1];

                                // Append the data
                                this._items.push({
                                    Author: file.createdBy.user["email"],
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

    // Labels file
    private static labelFile(file: Types.Microsoft.Graph.driveItem, overrideLabelFl: boolean, label: string, labelId: string, justification: string, responses: ISetSensitivityLabelResponse[]): PromiseLike<ISetSensitivityLabelResponse[]> {
        // See if this file has a sensitivity label
        if (file.sensitivityLabel?.id && !overrideLabelFl) {
            // Add a response
            responses.push({
                errorFl: false,
                fileName: file.name,
                message: `Skipping file, it's already labelled: '${file.sensitivityLabel.displayName}'.`,
                url: file.webUrl
            });
            return;
        }

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
                siteId: file.parentReference.siteId
            }).items(file.id).setSensitivityLabel("Manual", "Privileged", labelId, justification).execute(
                // Success
                () => {
                    // Add the response
                    responses.push({
                        errorFl: false,
                        fileName: file.name,
                        message: `The file was successfully labelled: ${label}.`,
                        url: file.webUrl
                    });

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
                    responses.push({
                        errorFl: true,
                        error: err,
                        fileName: file.name,
                        message: `There was an error tagging this file: ${error}`,
                        url: file.webUrl
                    });

                    // Resolve the request
                    resolve(responses);
                }
            );
        });
    }

    // Labels files for a specified folder
    private static labelFilesInFolder(webId: string, listName: string, folder: Types.SP.Folder, label: Components.IDropdownItem, overrideLabelFl: boolean, justification: string) {
        let responses: ISetSensitivityLabelResponse[] = [];

        // Show a loading dialog
        LoadingDialog.setHeader("Loading Items");
        LoadingDialog.setBody("Loading the files for this library...");
        LoadingDialog.show();

        // Load the files for this drive
        DataSource.loadFiles(webId, listName, folder).then(files => {
            // Update the loading dialog
            LoadingDialog.setBody("Applying the sensitivity labels to the files...");

            // Parse the files
            let counter = 0;
            Helper.Executor(files, file => {
                // Update the loading dialog
                LoadingDialog.setBody(`Processing ${++counter} of ${files.length} files...`);

                // Label the file
                return this.labelFile(file, overrideLabelFl, label.text, label.value, justification, responses);
            }).then(() => {
                // Show the responses
                this.showResponses(responses);

                // Hide the dialog
                LoadingDialog.hide();
            });
        });
    }

    // Renders the search summary
    private static renderSummary(el: HTMLElement, auditOnly: boolean, onClose: () => void) {
        // Render the summary
        let dt = new Dashboard({
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
                                                dt.updateCell(rowIdx, 3, label);
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
    }

    // Runs the report
    static run(el: HTMLElement, auditOnly: boolean, values: { [key: string]: string }, onClose: () => void) {
        let data: IWebItem[] = [];

        // Clear the items
        this._items = [];

        // Set the flags
        let withLabelsFl = false;
        let withoutLabelsFl = false;
        (values["SearchType"] as any as Components.ICheckboxGroupItem[]).forEach(item => {
            withLabelsFl = item.name == "WithLabels" ? true : withLabelsFl;
            withoutLabelsFl = item.name == "WithoutLabels" ? true : withoutLabelsFl;
        });

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
                    this.analyzeLibraries(siteItem.value, siteItem.text, libs.results, withLabelsFl, withoutLabelsFl).then(resolve);
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

    // Shows the form for labeling a list
    static setDefaultSensitivityLabelForFiles(webId: string, listName: string, defaultLabel: string, folders: Components.IDropdownItem[], disableSensitivityLabelOverride: boolean) {
        // Set the modal header
        Modal.clear();
        Modal.setHeader("Set Default Sensitivity Label");

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
                    value: defaultLabel,
                    onValidate: (ctrl, results) => {
                        // Ensure a selection exists
                        results.isValid = results.value && results.value.text ? true : false;
                        return results;
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "ListFolder",
                    label: "Select a Folder:",
                    description: "Targets a specific folder to tag, otherwise will apply to all files in the library.",
                    type: Components.FormControlTypes.Dropdown,
                    items: folders
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
                                let folder = values["ListFolder"].data;
                                let label: Components.IDropdownItem = values["SensitivityLabel"];
                                let overrideLabelFl: boolean = values["OverrideLabel"];

                                // Update the justification
                                let justification = values["Justification"].text;
                                justification = justification == "Other" ? values["JustificationOther"] : justification;

                                // Label the files
                                this.labelFilesInFolder(webId, listName, folder, label, overrideLabelFl, justification);
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
                                this.labelFile(file, true, label.text, label.value, justification, []).then(responses => {
                                    // See if it was successful
                                    if (!responses[0].errorFl) {
                                        // Call the event
                                        onUpdate(label.text);

                                        // Hide the dialogs
                                        LoadingDialog.hide();
                                        Modal.hide();
                                    }
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

    // Shows the responses for setting the sensitivity labels
    private static showResponses(responses: ISetSensitivityLabelResponse[]) {
        // Clear the modal and set the header
        Modal.clear();
        Modal.setType(Components.ModalTypes.Full);
        Modal.setHeader("Set Sensitivity Label Results");

        // Set the body
        new DataTable({
            el: Modal.BodyElement,
            rows: responses,
            columns: [
                {
                    name: "errorFl",
                    title: "Error?"
                },
                {
                    name: "",
                    title: "File Name",
                    onRenderCell: (el, col, item: ISetSensitivityLabelResponse) => {
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
        });

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
}