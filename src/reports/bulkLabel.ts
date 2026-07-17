import { Dashboard, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, Types, Web, v2 } from "gd-sprest-bs";
import { DataSource } from "../ds";
import Strings from "../strings";
import { ExportCSV } from "./exportCSV";
import { ISensitivityLabelItem } from "./sensitivityLabels";

export interface ISetSensitivityLabelResponse {
    errorFl: boolean;
    error?: any;
    fileName: string;
    message: string;
    url: string;
}

const CSVSensitivityLabelResponseFields = [
    "errorFl", "url", "fileName", "message"
]

/**
 * Bulk Label Form
 */
export class BulkLabel {
    private static _dashboard: Dashboard = null;
    private static _elSubNav: HTMLElement = null;
    private static _items: ISensitivityLabelItem[] = [];
    private static _stopFl: boolean = false;

    // Labels a file
    static labelFile(file: Types.Microsoft.Graph.driveItem, label: string, labelId: string, justification: string, responses: ISetSensitivityLabelResponse[]): PromiseLike<ISetSensitivityLabelResponse> {
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

                    // Resolve the request
                    resolve(response);
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

                    // Resolve the request
                    resolve(response);
                }
            );
        });
    }

    // Labels files
    static labelFiles(driveItems: Types.Microsoft.Graph.driveItem[], label: string, labelId: string, justification: string): PromiseLike<ISetSensitivityLabelResponse[]> {
        // Return a promise
        return new Promise(resolve => {
            let responses: ISetSensitivityLabelResponse[] = [];

            // Show the responses
            this.showResponses(responses, () => {
                // Stop the worker process
                worker.stop();
            });

            // Update the dialog
            this._elSubNav.children[0].innerHTML = `Processing ${driveItems.length} files...`;

            // Process the labels as we load the files
            let fileCounter = 0;
            let filesToProcess: Types.Microsoft.Graph.driveItem[] = driveItems;
            let processingCounter = 0;
            let processedCounter = 0;

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

                // See if this file is already has this label
                if (file.sensitivityLabel?.id == labelId) {
                    // Add a response
                    let response: ISetSensitivityLabelResponse = {
                        errorFl: false,
                        fileName: file.name,
                        message: `Skipping file, it's already labelled: '${file.sensitivityLabel.displayName}'.`,
                        url: file.webUrl
                    };
                    responses.push(response);
                    this._dashboard.Datatable.addRow(response);

                    // Update the dialog
                    this._elSubNav.children[1].innerHTML = `[Processed ${++processedCounter} of ${fileCounter}] File Already Labelled: ${response.fileName}`;

                    // Check the next file
                    return;
                }

                // Increment the # of files being processed
                processingCounter++;

                // Update the dialog
                this._elSubNav.children[1].innerHTML = `[Processing ${processedCounter} of ${fileCounter}] Labelling File: ${file.name}`;

                // Labels the file
                let labelFile = () => {
                    // Wait for the specified sleep time to avoid throttling
                    setTimeout(() => {
                        // Label the file
                        this.labelFile(file, label, labelId, justification, responses).then((response) => {
                            // Add the response
                            responses.push(response);
                            this._dashboard.Datatable.addRow(response);

                            // Update the dialog
                            this._elSubNav.children[1].innerHTML = `[Processed ${++processedCounter} of ${fileCounter}] File Labelled: ${file.name}`;

                            // Decrement the # of files being processed
                            processingCounter--;
                        });
                    }, sleepTime);
                }

                // Label the file
                labelFile();
            }, 100);

            // Set the completed event
            let onCompleted = () => {
                // Clear the sub-nav
                this._elSubNav.classList.add("d-none");

                // Clear the callback events
                ContextInfo.clearRateLimitCallbacks();

                // Resolve the request
                resolve(responses);
            };

            // Ensure the process is running
            worker.start();
        });
    }

    // Labels files for a specified folder
    private static labelFilesInFolder(webId: string, webUrl: string, listName: string, driveId: string, folderId: string, fileExtensions: string[], label: Components.IDropdownItem, replaceLabel: Components.IDropdownItem, overrideLabelFl: boolean, justification: string) {
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
                    this._elSubNav.children[1].innerHTML = `[Processed ${++processedCounter} of ${fileCounter}] Skipping File: ${response.fileName}`;

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
                    this._elSubNav.children[1].innerHTML = `[Processed ${++processedCounter} of ${fileCounter}] File Already Labelled: ${response.fileName}`;

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

                // Update the dialog
                this._elSubNav.children[1].innerHTML = `[Processed ${++processedCounter} of ${fileCounter}] File Already Labelled: ${response.fileName}`;

                // Check the next file
                return;
            }

            // Increment the # of files being processed
            processingCounter++;

            // Update the dialog
            this._elSubNav.children[1].innerHTML = `[Processing ${processedCounter} of ${fileCounter}] Labelling File: ${file.name}`;

            // Labels the file
            let labelFile = () => {
                // Wait for the specified sleep time to avoid throttling
                setTimeout(() => {
                    // Label the file
                    this.labelFile(file, fileLabel.text, fileLabel.value, justification, responses).then((response) => {
                        // Add the response
                        responses.push(response);
                        this._dashboard.Datatable.addRow(response);

                        // Update the dialog
                        this._elSubNav.children[1].innerHTML = `[Processed ${++processedCounter} of ${fileCounter}] File Labelled: ${file.name}`;

                        // Decrement the # of files being processed
                        processingCounter--;
                    });
                }, sleepTime);
            }

            // Label the file
            labelFile();
        }, 100);

        // Set the completed event
        let onCompleted = () => {
            // Clear the sub-nav
            this._elSubNav.classList.add("d-none");

            // Clear the callback events
            ContextInfo.clearRateLimitCallbacks();
        };

        // Load the files for this drive
        DataSource.loadFiles(webId, webUrl, driveId, listName, folderId, false, file => {
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

    // Sets the default label for a library
    private static setDefaultLabel(webUrl: string, listName: string, labelId: string): PromiseLike<boolean> {
        // Return a promise
        return new Promise(resolve => {
            // Show a loading dialog
            LoadingDialog.setHeader("Updating List");
            LoadingDialog.setBody("This dialog will close after the list is updated...");
            LoadingDialog.show();

            // Restore the permissions
            Web(webUrl, { requestDigest: DataSource.SiteContext.FormDigestValue })
                .Lists(listName).update({
                    DefaultSensitivityLabelForLibrary: labelId
                }).execute(() => {
                    // Resolve the request
                    resolve(true);

                    // Hide the dialog
                    LoadingDialog.hide();
                }, () => {
                    // Resolve the request
                    resolve(false);

                    // Hide the dialog
                    LoadingDialog.hide();
                });
        });
    }

    // Shows the form for a library
    static showLibraryForm(webId: string, webUrl: string, listName: string, driveId: string, defaultLabel: string, folders: Components.IDropdownItem[], disableSensitivityLabelOverride: boolean, fileTypes: string, onDefaultLabelSet: (labelId: string) => void) {
        let defaultItem: Components.IDropdownItem = { text: "All Folders", value: "" };

        // Set the modal header
        Modal.clear();
        Modal.setHeader("Sensitivity Labels");

        // Updates the controls based on the selected tab
        let updateControls = (tabName: string) => {
            // Clear the validation on the form
            form.clearValidation();

            // See which tab was selected
            switch (tabName) {
                case "Bulk Label":
                    // Show/Hide controls
                    form.getControl("FileTypes").show();
                    form.getControl("Justification").show();
                    form.getControl("ListFolder").show();
                    form.getControl("OverrideLabel").show();
                    form.getControl("ReplaceLabel").hide();
                    form.getControl("SensitivityLabel").setDescription("This will apply the selected label to the file if it currently is not labelled.");
                    break;
                case "Default Label":
                    form.getControl("FileTypes").hide();
                    form.getControl("Justification").hide();
                    form.getControl("JustificationOther").hide();
                    form.getControl("ListFolder").hide();
                    form.getControl("ListSubFolder1").hide();
                    form.getControl("ListSubFolder2").hide();
                    form.getControl("ListSubFolder3").hide();
                    form.getControl("ListSubFolder4").hide();
                    form.getControl("ListSubFolder5").hide();
                    form.getControl("OverrideLabel").hide();
                    form.getControl("ReplaceLabel").hide();
                    form.getControl("SensitivityLabel").setDescription("Select a label to be the default sensitivity label for this library.");
                    break;
                case "Replace Label":
                    // Show/Hide controls
                    form.getControl("FileTypes").show();
                    form.getControl("Justification").show();
                    form.getControl("ListFolder").show();
                    form.getControl("OverrideLabel").hide();
                    form.getControl("ReplaceLabel").show();
                    form.getControl("SensitivityLabel").setDescription("Select a label to search for and replace with the selection below.");
                    break;
            }

            // See if it's not the default label
            if (tabName != "Default Label") {
                let values = form.getValues();

                // See if we are showing the other justification control
                if ((values["Justification"] as Components.IDropdownItem)?.text == "Other") {
                    form.getControl("JustificationOther").show();
                }

                // See if we are showing the sub-folder controls
                if ((values["ListFolder"] as Components.IDropdownItem)?.value) {
                    form.getControl("ListSubFolder1").show();
                }
                if ((values["ListSubFolder1"] as Components.IDropdownItem)?.value) {
                    form.getControl("ListSubFolder2").show();
                }
                if ((values["ListSubFolder2"] as Components.IDropdownItem)?.value) {
                    form.getControl("ListSubFolder3").show();
                }
                if ((values["ListSubFolder3"] as Components.IDropdownItem)?.value) {
                    form.getControl("ListSubFolder4").show();
                }
                if ((values["ListSubFolder4"] as Components.IDropdownItem)?.value) {
                    form.getControl("ListSubFolder5").show();
                }
            }
        }

        // Render tabs
        let nav = Components.Nav({
            el: Modal.BodyElement,
            isTabs: true,
            isPills: true,
            onClick: (newTab) => {
                // Update the controls
                updateControls(newTab.tabName);
            },
            items: [
                {
                    title: "Default Label",
                    isActive: true
                },
                {
                    title: "Bulk Label"
                },
                {
                    title: "Replace Label"
                }
            ]
        });

        // Set the form
        let form = Components.Form({
            el: Modal.BodyElement,
            groupClassName: "my-3",
            onRendered: () => {
                // Set the folders to trigger the change events for the sub-folders
                form.getControl("ListFolder").dropdown.setItems([defaultItem].concat(folders));

                // Update the controls
                updateControls(nav.activeTab.tabName);
            },
            controls: [
                {
                    name: "SensitivityLabel",
                    label: "Select Sensitivity Label:",
                    description: "This will apply the selected label to the file if it currently is not labelled.",
                    errorMessage: "A selection is required.",
                    items: DataSource.SensitivityLabelItems,
                    type: Components.FormControlTypes.Dropdown,
                    value: defaultLabel,
                    onValidate: (ctrl, results) => {
                        // See if we require a selection
                        if (nav.activeTab.tabName == "Default Label") {
                            // Allow a null option to clear the setting
                            results.isValid = true;
                        } else {
                            // Ensure a selection exists
                            results.isValid = results.value && results.value.text ? true : false;
                        }

                        // Return the results
                        return results;
                    }
                } as Components.IFormControlPropsDropdown,
                {
                    name: "OverrideLabel",
                    label: "Override Label?",
                    description: "If you select to enable \"Override Label\", then SAT will attempt to force a label update on each file. This may fail under specific conditions.",
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
                    onValidate: (ctrl, results) => {
                        // See if we require a selection
                        if (nav.activeTab.tabName == "Default Label") {
                            // Ignore this control option
                            results.isValid = true;
                        } else {
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
                    errorMessage: "A justification is required.",
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
                    },
                    onValidate: (ctrl, results) => {
                        // See if we require a selection
                        if (nav.activeTab.tabName == "Default Label") {
                            // Ignore this control option
                            results.isValid = true;
                        } else {
                            // Ensure a value exists
                            results.isValid = results.value && results.value.text ? true : false;
                        }

                        // Return the results
                        return results;
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
                        // See if we require a selection
                        if (nav.activeTab.tabName == "Default Label") {
                            // Ignore this control option
                            results.isValid = true;
                        } else {
                            let item = form.getValues()["Justification"] as Components.IDropdownItem;

                            // See if we are expecting a justification
                            if (item.text == "Other") {
                                // Set the flag
                                results.isValid = results.value?.trim() ? true : false;
                            }
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

                                // Call the appropriate function based on the selection
                                switch (nav.activeTab.tabName) {
                                    case "Default Label":
                                        // Set the default label
                                        this.setDefaultLabel(webUrl, listName, label.value).then(successFl => {
                                            // See if it was successful
                                            if (successFl) {
                                                // Call the event to update the list
                                                onDefaultLabelSet(label.value);
                                            }
                                        });
                                        break;
                                    case "Bulk Label":
                                        // Label the files
                                        this.labelFilesInFolder(webId, webUrl, listName, driveId, targetFolder?.id, fileExtensions, label, null, overrideLabelFl, justification);
                                        break;
                                    case "Replace Label":
                                        // Label the files
                                        this.labelFilesInFolder(webId, webUrl, listName, driveId, targetFolder?.id, fileExtensions, label, replaceLabel, true, justification);
                                        break;
                                }
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
    static showResponses(responses: ISetSensitivityLabelResponse[], onClose?: () => void) {
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
                title: "Sensitivity Labels",
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
                            new ExportCSV("sensitivity_labels.csv", CSVSensitivityLabelResponseFields, this._items);
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
    static stop() {
        // Set the stop flag
        this._stopFl = true;
    }
}