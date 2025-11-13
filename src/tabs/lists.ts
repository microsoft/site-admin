import { Dashboard, DataTable, Modal, LoadingDialog } from "dattatable";
import { Components, Helper, SPTypes, Types, Web, v2 } from "gd-sprest-bs";
import { ExportCSV } from "../reports/exportCSV";
import { IAppProps } from "../app";
import { DataSource } from "../ds";
import { IDropdownItem } from "gd-sprest-bs/src/components/components";

interface IList {
    DefaultSensitivityLabel: string;
    HasUniqueRoleAssignments: boolean;
    IncludedInSearch: boolean;
    ItemCount: number;
    ListId: string;
    ListName: string;
    ListTemplateType: number;
    ListTemplate: string;
    ListUrl: string;
    ListViewUrl: string;
    WebId: string;
    WebUrl: string;
}

interface ISetSensitivityLabelResponse {
    errorFl: boolean;
    error?: any;
    fileName: string;
    message: string;
    url: string;
}

const CSVFields = [
    "WebUrl", "WebId", "ListName", "ListTemplateType", "ListTemplate", "ListUrl", "ListViewUrl", "HasUniqueRoleAssignments", "ItemCount", "IncludedInSearch", "DefaultSensitivityLabel"
]
const CSVSensitivityLabelResponseFields = [
    "errorFl", "url", "fileName", "message"
]

/**
 * Lists Tab
 */
export class ListsTab {
    private _appProps: IAppProps = null;
    private _el: HTMLElement = null;
    private _form: Components.IForm = null;
    private _items: IList[] = [];
    private _webs: Components.IDropdownItem[] = null;

    // List Template Types
    private static _listTemplates: { [key: number]: string } = {};

    // Constructor
    constructor(el: HTMLElement, appProps: IAppProps) {
        this._el = el;
        this._appProps = appProps;

        // Render the tab
        this.render();
    }

    // Analyzes a list
    private analyzeList(webUrl: string, webId: string, list: Types.SP.ListOData) {
        // Add a row for this entry
        this._items.push({
            DefaultSensitivityLabel: list.DefaultSensitivityLabelForLibrary,
            HasUniqueRoleAssignments: list.HasUniqueRoleAssignments,
            IncludedInSearch: !list.NoCrawl,
            ItemCount: list.ItemCount,
            ListId: list.Id,
            ListName: list.Title,
            ListTemplate: ListsTab._listTemplates[list.BaseTemplate],
            ListTemplateType: list.BaseTemplate,
            ListUrl: list.RootFolder.ServerRelativeUrl,
            ListViewUrl: list.DefaultViewUrl,
            WebId: webId,
            WebUrl: webUrl
        });
    }

    // Loads the root folders for the list
    private loadFolders(item: IList): PromiseLike<Components.IDropdownItem[]> {
        // Return a promise
        return new Promise(resolve => {
            let items: Components.IDropdownItem[] = [{ text: "All Files in Library", value: null }];

            // Load the folders for this list
            Web(item.WebUrl).Lists().getById(item.ListId).RootFolder().Folders().query({ Expand: ["ListItemAllFields"], OrderBy: ["Name"] }).execute(folders => {
                // Parse the folders
                folders.results.forEach(folder => {
                    // Ensure an item exists for this folder
                    if (folder.ListItemAllFields?.Id > 0) {
                        // Add the item
                        items.push({
                            data: folder,
                            text: folder.Name,
                            value: folder.Name
                        });
                    }
                });

                // Resolve the request
                resolve(items);
            }, () => {
                // Shouldn't happen but we'll render a blank list
                resolve(items);
            });
        });
    }

    // Loads the lists
    private loadLists(showHiddenLists: boolean = false, webItems?: IDropdownItem[]) {
        // Show a loading dialog
        LoadingDialog.setHeader("Searching Lists");
        LoadingDialog.setBody("Searching the site...");
        LoadingDialog.show();

        // Clear the items
        this._items = [];

        // Parse all the webs
        let counter = 0;
        let webs = webItems || DataSource.SiteItems;
        Helper.Executor(webs, siteItem => {
            // Set the loading dialog element
            let elLoadingDialog = document.createElement("div");
            elLoadingDialog.innerHTML = `<span>Search ${++counter} of ${webs.length}...</span><br/><span></span>`;
            let elStatus = elLoadingDialog.childNodes[2] as HTMLElement;

            // Update the loading dialog
            LoadingDialog.setBody(elLoadingDialog);

            // Return a promise
            return new Promise(resolve => {
                // Set the filter
                let Filter = showHiddenLists ? "" : "Hidden eq false";

                // Get the list templates
                let web = Web(siteItem.text, { requestDigest: DataSource.SiteContext.FormDigestValue });
                web.ListTemplates().execute(templates => {
                    // Parse the templates
                    for (let i = 0; i < templates.results.length; i++) {
                        let template = templates.results[i];

                        // Skip the template if we already set it
                        if (ListsTab._listTemplates[template.ListTemplateTypeKind]) { continue; }

                        // Set the template
                        ListsTab._listTemplates[template.ListTemplateTypeKind] = template.Name;
                    }
                });

                // Get the lists
                web.Lists().query({
                    Filter,
                    Expand: ["DefaultViewFormUrl", "RootFolder"],
                    Select: [
                        "BaseTemplate", "DefaultSensitivityLabelForLibrary", "Id", "NoCrawl",
                        "ItemCount", "Title", "HasUniqueRoleAssignments", "RootFolder/ServerRelativeUrl"
                    ]
                }).execute(lists => {
                    let ctrList = 0;

                    // Parse the lists
                    Helper.Executor(lists.results, list => {
                        // Update the status
                        elStatus.innerHTML = `Analyzing List ${++ctrList} of ${lists.results.length}...`;

                        // Analyze the list
                        return this.analyzeList(siteItem.text, siteItem.value, list);
                    }).then(resolve);
                }, true);
            });
        }).then(() => {
            // Render the summary
            this.render();

            // Hide the loading dialog
            LoadingDialog.hide();
        });
    }

    // Renders the tab
    private render() {
        // Clear the element
        while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

        // Render a dashboard
        let dt = new Dashboard({
            el: this._el,
            navigation: {
                title: "Lists/Libraries",
                showFilter: false,
                items: [{
                    text: "New Search",
                    className: "btn-outline-light",
                    isButton: true,
                    onClick: () => {
                        // Show the form
                        this.showForm();
                    }
                }],
                itemsEnd: [{
                    text: "Export to CSV",
                    className: "btn-outline-light me-2",
                    isButton: true,
                    onClick: () => {
                        // Export the CSV
                        new ExportCSV("listPermissions.csv", CSVFields, this._items);
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
                    dtProps.order = [[0, "asc"]];

                    // Return the properties
                    return dtProps;
                },
                columns: [
                    {
                        name: "WebUrl",
                        title: "Url"
                    },
                    {
                        name: "",
                        title: "List Info",
                        onRenderCell: (el, col, item: IList) => {
                            // Render the list information
                            el.innerHTML = `
                                <span><b>Name: </b>${item.ListName}</span>
                                <br/>
                                <span><b>Type: </b>${item.ListTemplate}</span>
                            `;
                        }
                    },
                    {
                        name: "ItemCount",
                        title: "# of Items"
                    },
                    {
                        name: "DefaultSensitivityLabel",
                        title: "Default Sensitivity Label",
                        onRenderCell: (el, col, item: IList) => {
                            // See if a value exists
                            if (item.DefaultSensitivityLabel) {
                                // Set the value
                                el.innerHTML = DataSource.getSensitivityLabel(item.DefaultSensitivityLabel);
                            }
                        }
                    },
                    {
                        name: "IncludedInSearch",
                        title: "Included In Search"
                    },
                    {
                        name: "HasUniqueRoleAssignments",
                        title: "Has Unique Permissions"
                    },
                    {
                        className: "text-end",
                        name: "",
                        title: "",
                        onRenderCell: (el, col, item: IList, rowIdx) => {
                            let isLibrary = item.ListTemplateType == SPTypes.ListTemplateType.DocumentLibrary ||
                                item.ListTemplateType == SPTypes.ListTemplateType.WebPageLibrary

                            // Render the buttons
                            let tooltips = Components.TooltipGroup({
                                el,
                                tooltips: [
                                    {
                                        content: "Click to view the list.",
                                        btnProps: {
                                            text: "View List",
                                            type: Components.ButtonTypes.OutlinePrimary,
                                            onClick: () => {
                                                // Show the list/library
                                                window.open(item.ListUrl, "_blank");
                                            }
                                        }
                                    }
                                ]
                            });

                            // Add the options to make changes
                            if (!this._appProps.auditOnly) {
                                let tooltip: Components.ITooltip = null;
                                tooltips.add({
                                    content: "Click to set the default sensitivity label.",
                                    btnProps: {
                                        isDisabled: !isLibrary || !DataSource.HasSensitivityLabels,
                                        text: "Default Label",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the form
                                            this.setDefaultSensitivityLabel(item);
                                        }
                                    }
                                });
                                tooltips.add({
                                    content: "Click to set the default sensitivity label for any files that aren't currently labelled.",
                                    btnProps: {
                                        isDisabled: !isLibrary || !DataSource.HasSensitivityLabels,
                                        text: "Label Files",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Load the folders for this list
                                            this.loadFolders(item).then(folders => {
                                                // Show the form
                                                this.setDefaultSensitivityLabelForFiles(item, folders);
                                            });
                                        }
                                    }
                                });
                                tooltips.add({
                                    assignTo: obj => { tooltip = obj; },
                                    content: `Click to ${item.IncludedInSearch ? "remove" : "add"} the content from the search index.`,
                                    btnProps: {
                                        text: "Update Search",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Show the form
                                            this.setListSearch(item, () => {
                                                // Flip the flag
                                                item.IncludedInSearch = !item.IncludedInSearch;
                                                tooltip.setContent(`Click to ${item.IncludedInSearch ? "remove" : "add"} the content from the search index.`);

                                                // Update the row cell
                                                dt.Datatable.datatable.cell({ row: rowIdx, column: 4 }).data(item.IncludedInSearch).draw();
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

    // Reverts the item permissions
    private setDefaultSensitivityLabel(item: IList) {
        // Set the modal header
        Modal.clear();
        Modal.setHeader("Set Default Sensitivity Label");

        // Set the form
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [
                {
                    name: "SensitivityLabel",
                    label: "Select Sensitivity Label:",
                    description: "This will set the default sensitivity label for this library.",
                    items: DataSource.SensitivityLabelItems,
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    value: item.DefaultSensitivityLabel
                } as Components.IFormControlPropsDropdown
            ]
        });

        // Set the footmer
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
                                let labelId = form.getValues()["SensitivityLabel"].value;

                                // Show a loading dialog
                                LoadingDialog.setHeader("Updating List");
                                LoadingDialog.setBody("This dialog will close after the list is updated...");
                                LoadingDialog.show();

                                // Set the logic to run after the update completes
                                let onComplete = () => {
                                    // Hide the dialogs
                                    LoadingDialog.hide();
                                    Modal.hide();
                                }

                                // Restore the permissions
                                Web(item.WebUrl, { requestDigest: DataSource.SiteContext.FormDigestValue })
                                    .Lists().getById(item.ListId).update({
                                        DefaultSensitivityLabelForLibrary: labelId
                                    }).execute(() => {
                                        // Update the item
                                        item.DefaultSensitivityLabel = labelId;

                                        // Update the data table
                                        // TODO

                                        // Run the complete logic
                                        onComplete();
                                    }, onComplete);
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

    // Reverts the item permissions
    private setDefaultSensitivityLabelForFiles(item: IList, folders: Components.IDropdownItem[]) {
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
                    value: item.DefaultSensitivityLabel,
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
                    isDisabled: this._appProps.disableSensitivityLabelOverride,
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
                                let responses: ISetSensitivityLabelResponse[] = [];

                                // Show a loading dialog
                                LoadingDialog.setHeader("Loading Items");
                                LoadingDialog.setBody("Loading the files for this library...");
                                LoadingDialog.show();

                                // Update the justification
                                let justification = values["Justification"].text;
                                justification = justification == "Other" ? values["JustificationOther"] : justification;

                                // Load the files for this drive
                                DataSource.loadFiles(item.WebId, item.ListName, folder).then(files => {
                                    // Update the loading dialog
                                    LoadingDialog.setBody("Applying the sensitivity labels to the files...");

                                    // Parse the files
                                    let counter = 0;
                                    Helper.Executor(files, file => {
                                        // Update the loading dialog
                                        LoadingDialog.setBody(`Processing ${++counter} of ${files.length} files...`);

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
                                            }).items(file.id).setSensitivityLabel("Manual", "Privileged", label.value, justification).execute(
                                                // Success
                                                () => {
                                                    // Add the response
                                                    responses.push({
                                                        errorFl: false,
                                                        fileName: file.name,
                                                        message: `The file was successfully labelled: ${label.text}.`,
                                                        url: file.webUrl
                                                    });

                                                    // Check the next file
                                                    resolve(null);
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

                                                    // Check the next file
                                                    resolve(null);
                                                }
                                            );
                                        });
                                    }).then(() => {
                                        // Show the responses
                                        this.showResponses(responses);

                                        // Hide the dialog
                                        LoadingDialog.hide();
                                    });
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

    // Adds/Removes the list content from the search index
    private setListSearch(item: IList, onUpdate: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader(item.IncludedInSearch ? "Remove From Search" : "Add To Search");

        // Set the body
        Modal.setBody(item.IncludedInSearch ? "Confirm to remove the contents of this list from the search index." : "Confirm to add the contents of this list to the search index.");

        // Set the footer
        Components.TooltipGroup({
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: item.IncludedInSearch ? "Remove the list from search" : "Add the list to search",
                    btnProps: {
                        text: "Update",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Show a loading dialog
                            LoadingDialog.setHeader("Update List");
                            LoadingDialog.setBody("This will close after the list has been updated...");
                            LoadingDialog.show();

                            // Update the list
                            Web(item.WebUrl, { requestDigest: DataSource.SiteContext.FormDigestValue }).Lists().getById(item.ListId).update({
                                NoCrawl: item.IncludedInSearch
                            }).execute(() => {
                                // Call the update event
                                onUpdate();

                                // Close the dialog
                                LoadingDialog.hide();
                                Modal.hide();
                            }, () => {
                                // Error updating the list
                                console.error("Error updating the search index for the list...");
                                LoadingDialog.hide();
                                Modal.hide();
                            });
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

    // Sets the webs
    setWebs(webs: Components.IDropdownItem[]) {
        this._webs = webs;

        // See if the form exists
        if (this._form) {
            // Get the control
            let ctrl = this._form.getControl("SelectedWebs");

            // Set the items
            ctrl.dropdown.setItems(this._webs);

            // Enable the control
            ctrl.dropdown.enable();
        }
    }

    // Shows the load form
    private showForm() {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Load Lists");

        // Set the form
        let ddlWebs: Components.IFormControl;
        let form = Components.Form({
            el: Modal.BodyElement,
            onRendered: () => {
                // Hide the selected webs
                ddlWebs.hide();
            },
            controls: [
                {
                    name: "SearchAll",
                    type: Components.FormControlTypes.Switch,
                    label: "Search All Webs?",
                    description: "Select this option to search all webs in this site.",
                    value: true,
                    onChange: (item => {
                        // Show or hide the control
                        item ? ddlWebs.hide() : ddlWebs.show();
                    })
                } as Components.IFormControlPropsSwitch,
                {
                    name: "SelectedWebs",
                    type: Components.FormControlTypes.MultiDropdownCheckbox,
                    isDisabled: this._webs == null,
                    items: this._webs,
                    label: "Selected Web(s):",
                    placeholder: "Select a web",
                    onControlRendered: ctrl => { ddlWebs = ctrl; },
                    onValidate: (ctrl, results) => {
                        // See if we are not searching all the webs
                        let values = form.getValues();
                        if (values["SearchAll"] == false) {
                            // Ensure a value exists
                            results.isValid = results.value ? true : false;
                            results.invalidMessage = "A selection is required.";
                        }

                        // Return the results
                        return results;
                    }
                } as Components.IFormControlPropsMultiDropdownCheckbox,
                {
                    name: "ShowHiddenLists",
                    type: Components.FormControlTypes.Switch,
                    label: "Show Hidden Lists?",
                    description: "Select this option to include hidden lists.",
                    value: false
                }
            ]
        });

        // Set the footer
        Components.TooltipGroup({
            buttonType: Components.ButtonTypes.OutlinePrimary,
            el: Modal.FooterElement,
            tooltips: [
                {
                    content: "Closes the form.",
                    btnProps: {
                        text: "Close",
                        onClick: () => {
                            // Hide the modal
                            Modal.hide();
                        }
                    }
                },
                {
                    content: "Runs the lists report.",
                    btnProps: {
                        text: "Run",
                        onClick: () => {
                            // Ensure the form is valid
                            if (form.isValid()) {
                                let values = form.getValues();

                                // See if we are searching all webs
                                if (values["SearchAll"]) {
                                    // Load the lists
                                    this.loadLists(true);
                                } else {
                                    // Load the lists
                                    this.loadLists(false, values["SelectedWebs"]);
                                }

                                // Close the dialog
                                Modal.hide();
                            }
                        }
                    }
                }
            ]
        });

        // Show the form
        Modal.show();
    }

    // Shows the responses for setting the sensitivity labels
    private showResponses(responses: ISetSensitivityLabelResponse[]) {
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